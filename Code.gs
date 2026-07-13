/**
 * サイクリングコミュニティ イベント企画・連絡ツール
 * Google Apps Script バックエンド
 *
 * スクリプトプロパティ設定（必須）：
 *   ADMIN_PASSWORD  : 管理者パスワード（任意の文字列）
 *   SHEET_ID        : Google スプレッドシートのID
 *
 * スクリプトプロパティ設定（推奨）：
 *   PASSWORD_PEPPER : パスワードハッシュ化用の共通ペッパー文字列（任意の長いランダム文字列）
 *
 * 会員登録の流れ：
 *   1. メールアドレス・ニックネームを入力し、利用規約・免責事項に同意して登録
 *   2. サーバーが仮パスワードを生成し、MailAppでメール送信（SHA-256+salt+pepperでハッシュ化して保存）
 *   3. 仮パスワードでログイン → mustChangePassword=true が返るのでパスワード変更を強制
 *   4. 変更後、通常利用可能
 *
 * members シート列構成：
 *   A: memberId  B: email  C: passwordHash  D: salt  E: nickname
 *   F: mustChangePassword(true/false)  G: agreedAt  H: registeredAt  I: updatedAt
 *
 * ⚠️ 旧スキーマ（ニックネーム＋パスワードのみ）からの移行時は、
 *    メールアドレスが存在しない既存会員は再登録が必要です。
 */

// ============================================================
// エントリーポイント
// ============================================================

/**
 * POSTリクエストのエントリーポイント
 * Content-Type: text/plain で受信し、JSON としてパース
 */
function doPost(e) {
  // CORS ヘッダーを付与したレスポンスを返すためのヘルパー
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    // リクエストボディを JSON パース
    var body = JSON.parse(e.postData.contents);
    var action = body.action;

    var result;

    switch (action) {
      case 'register':
        result = registerMember(body);
        break;
      case 'login':
        result = loginMember(body);
        break;
      case 'changePassword':
        result = changePassword(body);
        break;
      case 'updateNickname':
        result = updateNickname(body);
        break;
      case 'adminLogin':
        result = adminLogin(body);
        break;
      case 'getEvents':
        result = getEvents(body);
        break;
      case 'getEventDetail':
        result = getEventDetail(body);
        break;
      case 'createEvent':
        result = createEvent(body);
        break;
      case 'updateEvent':
        result = updateEvent(body);
        break;
      case 'deleteEvent':
        result = deleteEvent(body);
        break;
      case 'joinEvent':
        result = joinEvent(body);
        break;
      case 'cancelJoin':
        result = cancelJoin(body);
        break;
      case 'postMessage':
        result = postMessage(body);
        break;
      case 'deleteMessage':
        result = deleteMessage(body);
        break;
      case 'adminDeleteMember':
        result = adminDeleteMember(body);
        break;
      case 'adminGetAllMembers':
        result = adminGetAllMembers(body);
        break;
      case 'adminGetAllMessages':
        result = adminGetAllMessages(body);
        break;
      default:
        result = { status: 'error', message: '不明なアクションです: ' + action };
    }

    output.setContent(JSON.stringify(result));
  } catch (err) {
    output.setContent(JSON.stringify({
      status: 'error',
      message: 'サーバーエラーが発生しました: ' + err.message
    }));
  }

  return output;
}

/**
 * OPTIONSリクエスト（プリフライト）への対応
 * ※ GAS では doOptions は標準サポートされないため doGet で代替
 */
function doGet(e) {
  var output = ContentService.createTextOutput('OK');
  output.setMimeType(ContentService.MimeType.TEXT);
  return output;
}

// ============================================================
// スプレッドシート取得ユーティリティ
// ============================================================

/**
 * スプレッドシートオブジェクトを取得する
 * SHEET_ID はスクリプトプロパティから取得
 */
function getSpreadsheet() {
  var sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!sheetId) {
    throw new Error('SHEET_ID がスクリプトプロパティに設定されていません');
  }
  return SpreadsheetApp.openById(sheetId);
}

/**
 * 指定シート名のシートオブジェクトを取得する
 */
function getSheet(sheetName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('シート「' + sheetName + '」が見つかりません');
  }
  return sheet;
}

// ============================================================
// UUID 生成
// ============================================================

/**
 * シンプルなUUIDを生成する（v4相当）
 */
function generateUUID() {
  return Utilities.getUuid();
}

// ============================================================
// パスワードハッシュ化
// ============================================================

/**
 * パスワードをSHA-256（salt + pepper付き）でハッシュ化し、16進文字列で返す
 * @param {string} password - 平文パスワード
 * @param {string} salt - 会員ごとのランダムsalt（membersシートに保存）
 * @returns {string} SHA-256ハッシュ値（64文字の16進文字列）
 */
function hashPassword(password, salt) {
  var pepper = PropertiesService.getScriptProperties().getProperty('PASSWORD_PEPPER') || '';
  var combined = String(password) + String(salt || '') + pepper;
  var digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    combined,
    Utilities.Charset.UTF_8
  );
  var hex = '';
  for (var i = 0; i < digest.length; i++) {
    var byte = digest[i];
    if (byte < 0) byte += 256; // 符号付きバイトを補正
    var h = byte.toString(16);
    if (h.length === 1) h = '0' + h;
    hex += h;
  }
  return hex;
}

/**
 * ランダムな仮パスワードを生成する（紛らわしい文字は除外）
 */
function generateTempPassword() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  var length = 10;
  var pw = '';
  for (var i = 0; i < length; i++) {
    pw += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return pw;
}

// ============================================================
// レート制限（CacheService）
// ============================================================

/**
 * 指定キーの試行回数をチェックし、上限内であればカウントアップする
 * @returns {boolean} 上限内なら true、上限超過なら false
 */
function checkRateLimit(key, maxAttempts, windowSeconds) {
  var cache = CacheService.getScriptCache();
  var countStr = cache.get(key);
  var count = countStr ? parseInt(countStr, 10) : 0;
  if (count >= maxAttempts) return false;
  cache.put(key, String(count + 1), windowSeconds);
  return true;
}

/**
 * レート制限カウンタをリセットする（ログイン成功時など）
 */
function clearRateLimit(key) {
  CacheService.getScriptCache().remove(key);
}

// ============================================================
// メール送信
// ============================================================

/**
 * 仮パスワードを会員にメール送信する
 */
function sendTempPasswordEmail(email, nickname, tempPassword) {
  var subject = '【サイクリングコミュニティ】仮パスワードのお知らせ';
  var body =
    nickname + ' 様\n\n' +
    'サイクリングコミュニティへご登録いただきありがとうございます。\n' +
    '以下の仮パスワードでログインし、初回ログイン時に画面の案内に従って\n' +
    '新しいパスワードへの変更をお願いいたします。\n\n' +
    '仮パスワード： ' + tempPassword + '\n\n' +
    'ログインページ： https://kenken6291.github.io/cycling-community/\n\n' +
    '※このメールにお心当たりがない場合は、お手数ですが破棄してください。\n\n' +
    '-----------------------------------\n' +
    'サイクリングコミュニティ運営';
  MailApp.sendEmail(email, subject, body);
}

// ============================================================
// トークン管理
// ============================================================

/**
 * ログイントークンを生成してスクリプトプロパティに保存する
 * トークンキー: TOKEN_{memberId}
 * 有効期限: 24時間（ミリ秒）
 */
function generateToken(memberId) {
  var token = Utilities.getUuid();
  var expiry = new Date().getTime() + (24 * 60 * 60 * 1000); // 24時間後
  var tokenData = JSON.stringify({ token: token, expiry: expiry });
  PropertiesService.getScriptProperties().setProperty('TOKEN_' + memberId, tokenData);
  return token;
}

/**
 * トークンを検証する
 * @returns {boolean} 有効な場合 true
 */
function verifyToken(memberId, token) {
  if (!memberId || !token) return false;
  var stored = PropertiesService.getScriptProperties().getProperty('TOKEN_' + memberId);
  if (!stored) return false;
  try {
    var data = JSON.parse(stored);
    if (data.token !== token) return false;
    if (new Date().getTime() > data.expiry) {
      // 期限切れ：削除する
      PropertiesService.getScriptProperties().deleteProperty('TOKEN_' + memberId);
      return false;
    }
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * 管理者トークンを生成してスクリプトプロパティに保存する
 */
function generateAdminToken() {
  var token = Utilities.getUuid();
  var expiry = new Date().getTime() + (24 * 60 * 60 * 1000);
  var tokenData = JSON.stringify({ token: token, expiry: expiry });
  PropertiesService.getScriptProperties().setProperty('ADMIN_TOKEN', tokenData);
  return token;
}

/**
 * 管理者トークンを検証する
 */
function verifyAdminToken(token) {
  if (!token) return false;
  var stored = PropertiesService.getScriptProperties().getProperty('ADMIN_TOKEN');
  if (!stored) return false;
  try {
    var data = JSON.parse(stored);
    if (data.token !== token) return false;
    if (new Date().getTime() > data.expiry) {
      PropertiesService.getScriptProperties().deleteProperty('ADMIN_TOKEN');
      return false;
    }
    return true;
  } catch (e) {
    return false;
  }
}

// ============================================================
// 日時フォーマット
// ============================================================

/**
 * 現在日時を "YYYY-MM-DD HH:mm:ss" 形式の文字列で返す
 */
function nowString() {
  var d = new Date();
  var pad = function(n) { return n < 10 ? '0' + n : '' + n; };
  return d.getFullYear() + '-' +
    pad(d.getMonth() + 1) + '-' +
    pad(d.getDate()) + ' ' +
    pad(d.getHours()) + ':' +
    pad(d.getMinutes()) + ':' +
    pad(d.getSeconds());
}

// ============================================================
// 会員登録 (register)
// ============================================================

/**
 * 新規会員登録（メールアドレス＋ニックネーム。パスワードは仮発行してメール送信）
 * @param {Object} body - { email, nickname, agreedToTerms }
 */
function registerMember(body) {
  var email = (body.email || '').trim().toLowerCase();
  var nickname = (body.nickname || '').trim();
  var agreedToTerms = !!body.agreedToTerms;

  // バリデーション
  var emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!email || !emailPattern.test(email)) {
    return { status: 'error', message: '有効なメールアドレスを入力してください' };
  }
  if (!nickname || nickname.length < 2 || nickname.length > 20) {
    return { status: 'error', message: 'ニックネームは2〜20文字で入力してください' };
  }
  if (!agreedToTerms) {
    return { status: 'error', message: '利用規約・免責事項への同意が必要です' };
  }

  // 登録試行のレート制限（同一メールアドレスからの連続登録を防止）
  if (!checkRateLimit('reg_' + email, 3, 3600)) {
    return { status: 'error', message: '登録の試行回数が上限に達しました。しばらくしてから再度お試しください' };
  }

  var sheet = getSheet('members');
  var data = sheet.getDataRange().getValues();

  // メールアドレス・ニックネームの重複チェック（1行目はヘッダーのためスキップ）
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase() === email) {
      return { status: 'error', message: 'このメールアドレスはすでに登録されています' };
    }
    if (data[i][4] === nickname) {
      return { status: 'error', message: 'そのニックネームはすでに使用されています' };
    }
  }

  // 仮パスワードを生成し、salt+pepper付きSHA-256でハッシュ化して保存
  var memberId = generateUUID();
  var tempPassword = generateTempPassword();
  var salt = generateUUID();
  var passwordHash = hashPassword(tempPassword, salt);
  var now = nowString();

  sheet.appendRow([
    memberId,       // A: memberId
    email,          // B: email
    passwordHash,   // C: passwordHash
    salt,           // D: salt
    nickname,       // E: nickname
    true,           // F: mustChangePassword
    now,            // G: agreedAt
    now,            // H: registeredAt
    now             // I: updatedAt
  ]);

  // 仮パスワードをメール送信。失敗した場合は登録自体をロールバックする
  try {
    sendTempPasswordEmail(email, nickname, tempPassword);
  } catch (mailErr) {
    sheet.deleteRow(sheet.getLastRow());
    return { status: 'error', message: 'メール送信に失敗しました。メールアドレスをご確認のうえ、再度お試しください' };
  }

  return {
    status: 'ok',
    data: { message: '仮パスワードをメールで送信しました。メールをご確認のうえログインしてください' }
  };
}

// ============================================================
// ログイン (login)
// ============================================================

/**
 * 会員ログイン
 * @param {Object} body - { email, password }
 */
function loginMember(body) {
  var email = (body.email || '').trim().toLowerCase();
  var password = (body.password || '').trim();

  if (!email || !password) {
    return { status: 'error', message: 'メールアドレスとパスワードを入力してください' };
  }

  // ログイン試行のレート制限（総当たり攻撃対策）
  var rateLimitKey = 'login_' + email;
  if (!checkRateLimit(rateLimitKey, 5, 900)) {
    return { status: 'error', message: 'ログイン試行回数が上限に達しました。しばらくしてから再度お試しください' };
  }

  var sheet = getSheet('members');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase() === email) {
      var salt = data[i][3];
      var hashedInput = hashPassword(password, salt);
      if (data[i][2] === hashedInput) {
        var memberId = data[i][0];
        var token = generateToken(memberId);
        clearRateLimit(rateLimitKey);
        return {
          status: 'ok',
          data: {
            memberId: memberId,
            nickname: data[i][4],
            email: data[i][1],
            token: token,
            mustChangePassword: data[i][5] === true || String(data[i][5]).toLowerCase() === 'true'
          }
        };
      }
      break; // メールアドレスは一致したがパスワードが違う場合もこの後は共通の汎用エラーを返す
    }
  }

  // アカウントの有無を区別しない汎用エラー（アカウント列挙対策）
  return { status: 'error', message: 'メールアドレスまたはパスワードが違います' };
}

// ============================================================
// パスワード変更 (changePassword)
// ============================================================

/**
 * パスワードを変更する（本人のみ）。初回ログイン時の強制変更にも使用。
 * @param {Object} body - { memberId, token, currentPassword, newPassword }
 */
function changePassword(body) {
  if (!verifyToken(body.memberId, body.token)) {
    return { status: 'error', message: 'ログインが必要です。再度ログインしてください' };
  }

  var currentPassword = (body.currentPassword || '').trim();
  var newPassword = (body.newPassword || '').trim();

  if (!currentPassword || !newPassword) {
    return { status: 'error', message: '現在のパスワードと新しいパスワードを入力してください' };
  }
  if (newPassword.length < 4) {
    return { status: 'error', message: '新しいパスワードは4文字以上で入力してください' };
  }

  var sheet = getSheet('members');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === body.memberId) {
      var salt = data[i][3];
      var hashedCurrent = hashPassword(currentPassword, salt);
      if (data[i][2] !== hashedCurrent) {
        return { status: 'error', message: '現在のパスワードが正しくありません' };
      }

      var newSalt = generateUUID();
      var newHash = hashPassword(newPassword, newSalt);
      sheet.getRange(i + 1, 3).setValue(newHash);      // C: passwordHash
      sheet.getRange(i + 1, 4).setValue(newSalt);      // D: salt
      sheet.getRange(i + 1, 6).setValue(false);        // F: mustChangePassword
      sheet.getRange(i + 1, 9).setValue(nowString());  // I: updatedAt

      return { status: 'ok', data: { message: 'パスワードを変更しました' } };
    }
  }

  return { status: 'error', message: '会員情報が見つかりません' };
}

// ============================================================
// ニックネーム変更 (updateNickname)
// ============================================================

/**
 * ニックネームを変更する（本人のみ）
 * @param {Object} body - { memberId, token, nickname }
 */
function updateNickname(body) {
  if (!verifyToken(body.memberId, body.token)) {
    return { status: 'error', message: 'ログインが必要です。再度ログインしてください' };
  }

  var newNickname = (body.nickname || '').trim();
  if (!newNickname || newNickname.length < 2 || newNickname.length > 20) {
    return { status: 'error', message: 'ニックネームは2〜20文字で入力してください' };
  }

  var sheet = getSheet('members');
  var data = sheet.getDataRange().getValues();

  // 重複チェック（自分以外）
  for (var i = 1; i < data.length; i++) {
    if (data[i][4] === newNickname && data[i][0] !== body.memberId) {
      return { status: 'error', message: 'そのニックネームはすでに使用されています' };
    }
  }

  for (var j = 1; j < data.length; j++) {
    if (data[j][0] === body.memberId) {
      sheet.getRange(j + 1, 5).setValue(newNickname);   // E: nickname
      sheet.getRange(j + 1, 9).setValue(nowString());   // I: updatedAt
      return { status: 'ok', data: { message: 'ニックネームを変更しました', nickname: newNickname } };
    }
  }

  return { status: 'error', message: '会員情報が見つかりません' };
}

// ============================================================
// 管理者ログイン (adminLogin)
// ============================================================

/**
 * 管理者ログイン
 * @param {Object} body - { password }
 */
function adminLogin(body) {
  var password = (body.password || '').trim();
  var adminPassword = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');

  if (!adminPassword) {
    return { status: 'error', message: 'ADMIN_PASSWORD がスクリプトプロパティに設定されていません' };
  }

  if (password !== adminPassword) {
    return { status: 'error', message: '管理者パスワードが違います' };
  }

  var token = generateAdminToken();
  return {
    status: 'ok',
    data: { adminToken: token, message: '管理者ログイン成功' }
  };
}

// ============================================================
// イベント一覧取得 (getEvents)
// ============================================================

/**
 * イベント一覧を取得する（認証不要）
 * @param {Object} body - { filter: { difficulty, month, rinko } }（省略可）
 */
function getEvents(body) {
  var sheet = getSheet('events');
  var data = sheet.getDataRange().getValues();

  // participantsシートで各イベントの参加者数を集計
  var participantsSheet = getSheet('participants');
  var participantsData = participantsSheet.getDataRange().getValues();
  var participantCount = {};
  for (var pi = 1; pi < participantsData.length; pi++) {
    var eid = participantsData[pi][1]; // イベントID
    participantCount[eid] = (participantCount[eid] || 0) + 1;
  }

  var events = [];
  var filter = body.filter || {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue; // 空行スキップ

    var eventId     = row[0];
    var difficulty  = row[9];  // 列J: 難易度
    var eventDate   = row[4];  // 列E: 開催日
    var rinko       = row[13]; // 列N: 輪行区間

    // フィルタリング
    if (filter.difficulty && filter.difficulty !== difficulty) continue;
    if (filter.rinko && filter.rinko === 'あり' && rinko !== 'あり') continue;
    if (filter.month) {
      var dateStr = String(eventDate);
      var monthStr = dateStr.substring(0, 7); // "YYYY-MM"
      if (monthStr !== filter.month) continue;
    }

    events.push({
      eventId:        eventId,
      organizerId:    row[1],
      organizerName:  row[2],
      title:          row[3],
      eventDate:      row[4],
      location:       row[5],
      routeOverview:  row[6],
      distance:       row[7],
      elevation:      row[8],
      difficulty:     row[9],
      capacity:       row[10],
      deadline:       row[11],
      notes:          row[12],
      rinko:          row[13],
      contact:        row[14],
      createdAt:      row[15],
      updatedAt:      row[16],
      participantCount: participantCount[eventId] || 0
    });
  }

  // 開催日の昇順でソート
  events.sort(function(a, b) {
    return String(a.eventDate).localeCompare(String(b.eventDate));
  });

  return { status: 'ok', data: { events: events } };
}

// ============================================================
// イベント詳細取得 (getEventDetail)
// ============================================================

/**
 * イベント詳細・参加者一覧・コメント一覧を取得する
 * @param {Object} body - { eventId }
 */
function getEventDetail(body) {
  var eventId = body.eventId;
  if (!eventId) {
    return { status: 'error', message: 'イベントIDが指定されていません' };
  }

  // イベント情報を取得
  var eventsSheet = getSheet('events');
  var eventsData = eventsSheet.getDataRange().getValues();
  var eventRow = null;
  for (var i = 1; i < eventsData.length; i++) {
    if (eventsData[i][0] === eventId) {
      eventRow = eventsData[i];
      break;
    }
  }
  if (!eventRow) {
    return { status: 'error', message: 'イベントが見つかりません' };
  }

  var event = {
    eventId:        eventRow[0],
    organizerId:    eventRow[1],
    organizerName:  eventRow[2],
    title:          eventRow[3],
    eventDate:      eventRow[4],
    location:       eventRow[5],
    routeOverview:  eventRow[6],
    distance:       eventRow[7],
    elevation:      eventRow[8],
    difficulty:     eventRow[9],
    capacity:       eventRow[10],
    deadline:       eventRow[11],
    notes:          eventRow[12],
    rinko:          eventRow[13],
    contact:        eventRow[14],
    createdAt:      eventRow[15],
    updatedAt:      eventRow[16]
  };

  // 参加者一覧を取得
  var participantsSheet = getSheet('participants');
  var participantsData = participantsSheet.getDataRange().getValues();
  var participants = [];
  for (var pi = 1; pi < participantsData.length; pi++) {
    if (participantsData[pi][1] === eventId) {
      participants.push({
        joinId:     participantsData[pi][0],
        eventId:    participantsData[pi][1],
        memberId:   participantsData[pi][2],
        nickname:   participantsData[pi][3],
        joinedAt:   participantsData[pi][4],
        message:    participantsData[pi][5]
      });
    }
  }

  // コメント一覧を取得
  var messagesSheet = getSheet('messages');
  var messagesData = messagesSheet.getDataRange().getValues();
  var messages = [];
  for (var mi = 1; mi < messagesData.length; mi++) {
    if (messagesData[mi][1] === eventId) {
      messages.push({
        messageId:  messagesData[mi][0],
        eventId:    messagesData[mi][1],
        memberId:   messagesData[mi][2],
        nickname:   messagesData[mi][3],
        content:    messagesData[mi][4],
        postedAt:   messagesData[mi][5]
      });
    }
  }

  // 投稿日時昇順でソート
  messages.sort(function(a, b) {
    return String(a.postedAt).localeCompare(String(b.postedAt));
  });

  return {
    status: 'ok',
    data: { event: event, participants: participants, messages: messages }
  };
}

// ============================================================
// イベント作成 (createEvent)
// ============================================================

/**
 * 新規イベントを作成する（ログイン必要）
 * @param {Object} body - { memberId, token, event: { ... } }
 */
function createEvent(body) {
  // 認証チェック
  if (!verifyToken(body.memberId, body.token)) {
    return { status: 'error', message: 'ログインが必要です。再度ログインしてください' };
  }

  var ev = body.event || {};

  // 必須項目バリデーション
  if (!ev.title || !ev.eventDate || !ev.location) {
    return { status: 'error', message: 'タイトル・開催日・集合場所は必須です' };
  }

  // 主催者のニックネームを会員マスタから取得
  var membersSheet = getSheet('members');
  var membersData = membersSheet.getDataRange().getValues();
  var organizerNickname = '';
  for (var i = 1; i < membersData.length; i++) {
    if (membersData[i][0] === body.memberId) {
      organizerNickname = membersData[i][4];
      break;
    }
  }
  if (!organizerNickname) {
    return { status: 'error', message: '会員情報が見つかりません' };
  }

  var eventId = generateUUID();
  var now = nowString();

  var sheet = getSheet('events');
  sheet.appendRow([
    eventId,                          // A: イベントID
    body.memberId,                    // B: 主催者会員ID
    organizerNickname,                // C: 主催者ニックネーム
    ev.title || '',                   // D: タイトル
    ev.eventDate || '',               // E: 開催日
    ev.location || '',                // F: 集合場所
    ev.routeOverview || '',           // G: ルート概要
    ev.distance || '',                // H: 距離
    ev.elevation || '',               // I: 獲得標高
    ev.difficulty || '初心者向け',    // J: 難易度
    ev.capacity || '',                // K: 定員
    ev.deadline || '',                // L: 募集締切日
    ev.notes || '',                   // M: 詳細・備考
    ev.rinko || 'なし',               // N: 輪行区間
    ev.contact || '',                 // O: 連絡先
    now,                              // P: 作成日時
    now                               // Q: 更新日時
  ]);

  return {
    status: 'ok',
    data: { message: 'イベントを作成しました', eventId: eventId }
  };
}

// ============================================================
// イベント編集 (updateEvent)
// ============================================================

/**
 * イベントを編集する（主催者本人または管理者のみ）
 * @param {Object} body - { memberId, token, adminToken, eventId, event: { ... } }
 */
function updateEvent(body) {
  var isAdmin = verifyAdminToken(body.adminToken);
  var isMember = verifyToken(body.memberId, body.token);

  if (!isAdmin && !isMember) {
    return { status: 'error', message: 'ログインが必要です。再度ログインしてください' };
  }

  var sheet = getSheet('events');
  var data = sheet.getDataRange().getValues();

  // 対象行を検索
  var targetRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === body.eventId) {
      // 管理者でなければ主催者本人かチェック
      if (!isAdmin && data[i][1] !== body.memberId) {
        return { status: 'error', message: '編集権限がありません' };
      }
      targetRow = i + 1; // シートは1インデックス
      break;
    }
  }

  if (targetRow === -1) {
    return { status: 'error', message: 'イベントが見つかりません' };
  }

  var ev = body.event || {};
  var now = nowString();

  // 各列を更新（B・C列=主催者情報は変更しない）
  sheet.getRange(targetRow, 4).setValue(ev.title || data[targetRow-1][3]);
  sheet.getRange(targetRow, 5).setValue(ev.eventDate || data[targetRow-1][4]);
  sheet.getRange(targetRow, 6).setValue(ev.location || data[targetRow-1][5]);
  sheet.getRange(targetRow, 7).setValue(ev.routeOverview !== undefined ? ev.routeOverview : data[targetRow-1][6]);
  sheet.getRange(targetRow, 8).setValue(ev.distance !== undefined ? ev.distance : data[targetRow-1][7]);
  sheet.getRange(targetRow, 9).setValue(ev.elevation !== undefined ? ev.elevation : data[targetRow-1][8]);
  sheet.getRange(targetRow, 10).setValue(ev.difficulty || data[targetRow-1][9]);
  sheet.getRange(targetRow, 11).setValue(ev.capacity !== undefined ? ev.capacity : data[targetRow-1][10]);
  sheet.getRange(targetRow, 12).setValue(ev.deadline !== undefined ? ev.deadline : data[targetRow-1][11]);
  sheet.getRange(targetRow, 13).setValue(ev.notes !== undefined ? ev.notes : data[targetRow-1][12]);
  sheet.getRange(targetRow, 14).setValue(ev.rinko || data[targetRow-1][13]);
  sheet.getRange(targetRow, 15).setValue(ev.contact !== undefined ? ev.contact : data[targetRow-1][14]);
  sheet.getRange(targetRow, 17).setValue(now); // Q: 更新日時

  return { status: 'ok', data: { message: 'イベントを更新しました' } };
}

// ============================================================
// イベント削除 (deleteEvent)
// ============================================================

/**
 * イベントを削除する（主催者本人または管理者のみ）
 * 関連する参加申込・コメントも削除する
 * @param {Object} body - { memberId, token, adminToken, eventId }
 */
function deleteEvent(body) {
  var isAdmin = verifyAdminToken(body.adminToken);
  var isMember = verifyToken(body.memberId, body.token);

  if (!isAdmin && !isMember) {
    return { status: 'error', message: 'ログインが必要です。再度ログインしてください' };
  }

  // イベントシートから削除
  var eventsSheet = getSheet('events');
  var eventsData = eventsSheet.getDataRange().getValues();
  var targetRow = -1;

  for (var i = 1; i < eventsData.length; i++) {
    if (eventsData[i][0] === body.eventId) {
      if (!isAdmin && eventsData[i][1] !== body.memberId) {
        return { status: 'error', message: '削除権限がありません' };
      }
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    return { status: 'error', message: 'イベントが見つかりません' };
  }

  eventsSheet.deleteRow(targetRow);

  // 関連する参加申込を削除
  deleteRowsByEventId('participants', body.eventId);

  // 関連するコメントを削除
  deleteRowsByEventId('messages', body.eventId);

  return { status: 'ok', data: { message: 'イベントを削除しました' } };
}

/**
 * 指定シートからイベントIDに紐づく行をすべて削除する
 */
function deleteRowsByEventId(sheetName, eventId) {
  var sheet = getSheet(sheetName);
  var data = sheet.getDataRange().getValues();

  // 後ろから削除（行番号ずれを防ぐ）
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === eventId) {
      sheet.deleteRow(i + 1);
    }
  }
}

// ============================================================
// 参加申込 (joinEvent)
// ============================================================

/**
 * イベントに参加申込する（ログイン必要）
 * @param {Object} body - { memberId, token, eventId, message }
 */
function joinEvent(body) {
  if (!verifyToken(body.memberId, body.token)) {
    return { status: 'error', message: 'ログインが必要です。再度ログインしてください' };
  }

  // イベント存在チェック
  var eventsSheet = getSheet('events');
  var eventsData = eventsSheet.getDataRange().getValues();
  var eventExists = false;
  for (var i = 1; i < eventsData.length; i++) {
    if (eventsData[i][0] === body.eventId) {
      eventExists = true;
      break;
    }
  }
  if (!eventExists) {
    return { status: 'error', message: 'イベントが見つかりません' };
  }

  // 重複申込チェック
  var participantsSheet = getSheet('participants');
  var participantsData = participantsSheet.getDataRange().getValues();
  for (var pi = 1; pi < participantsData.length; pi++) {
    if (participantsData[pi][1] === body.eventId && participantsData[pi][2] === body.memberId) {
      return { status: 'error', message: 'すでに参加申込済みです' };
    }
  }

  // 主催者本人の申込チェック
  for (var ei = 1; ei < eventsData.length; ei++) {
    if (eventsData[ei][0] === body.eventId && eventsData[ei][1] === body.memberId) {
      return { status: 'error', message: '自分が主催するイベントには参加申込できません' };
    }
  }

  // ニックネームを取得
  var membersSheet = getSheet('members');
  var membersData = membersSheet.getDataRange().getValues();
  var nickname = '';
  for (var mi = 1; mi < membersData.length; mi++) {
    if (membersData[mi][0] === body.memberId) {
      nickname = membersData[mi][4];
      break;
    }
  }

  var joinId = generateUUID();
  var now = nowString();
  participantsSheet.appendRow([
    joinId,
    body.eventId,
    body.memberId,
    nickname,
    now,
    body.message || ''
  ]);

  return { status: 'ok', data: { message: '参加申込が完了しました', joinId: joinId } };
}

// ============================================================
// 参加キャンセル (cancelJoin)
// ============================================================

/**
 * 参加申込をキャンセルする（本人のみ）
 * @param {Object} body - { memberId, token, eventId }
 */
function cancelJoin(body) {
  if (!verifyToken(body.memberId, body.token)) {
    return { status: 'error', message: 'ログインが必要です。再度ログインしてください' };
  }

  var sheet = getSheet('participants');
  var data = sheet.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === body.eventId && data[i][2] === body.memberId) {
      sheet.deleteRow(i + 1);
      return { status: 'ok', data: { message: '参加をキャンセルしました' } };
    }
  }

  return { status: 'error', message: '参加申込が見つかりません' };
}

// ============================================================
// コメント投稿 (postMessage)
// ============================================================

/**
 * コメントを投稿する（ログイン必要）
 * @param {Object} body - { memberId, token, eventId, content }
 */
function postMessage(body) {
  if (!verifyToken(body.memberId, body.token)) {
    return { status: 'error', message: 'ログインが必要です。再度ログインしてください' };
  }

  var content = (body.content || '').trim();
  if (!content || content.length > 500) {
    return { status: 'error', message: 'コメントは1〜500文字で入力してください' };
  }

  // イベント存在チェック
  var eventsSheet = getSheet('events');
  var eventsData = eventsSheet.getDataRange().getValues();
  var eventExists = false;
  for (var i = 1; i < eventsData.length; i++) {
    if (eventsData[i][0] === body.eventId) {
      eventExists = true;
      break;
    }
  }
  if (!eventExists) {
    return { status: 'error', message: 'イベントが見つかりません' };
  }

  // ニックネームを取得
  var membersSheet = getSheet('members');
  var membersData = membersSheet.getDataRange().getValues();
  var nickname = '';
  for (var mi = 1; mi < membersData.length; mi++) {
    if (membersData[mi][0] === body.memberId) {
      nickname = membersData[mi][4];
      break;
    }
  }

  var messageId = generateUUID();
  var now = nowString();
  var messagesSheet = getSheet('messages');
  messagesSheet.appendRow([
    messageId,
    body.eventId,
    body.memberId,
    nickname,
    content,
    now
  ]);

  return { status: 'ok', data: { message: 'コメントを投稿しました', messageId: messageId } };
}

// ============================================================
// コメント削除 (deleteMessage)
// ============================================================

/**
 * コメントを削除する（投稿者本人または管理者のみ）
 * @param {Object} body - { memberId, token, adminToken, messageId }
 */
function deleteMessage(body) {
  var isAdmin = verifyAdminToken(body.adminToken);
  var isMember = verifyToken(body.memberId, body.token);

  if (!isAdmin && !isMember) {
    return { status: 'error', message: 'ログインが必要です。再度ログインしてください' };
  }

  var sheet = getSheet('messages');
  var data = sheet.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === body.messageId) {
      if (!isAdmin && data[i][2] !== body.memberId) {
        return { status: 'error', message: '削除権限がありません' };
      }
      sheet.deleteRow(i + 1);
      return { status: 'ok', data: { message: 'コメントを削除しました' } };
    }
  }

  return { status: 'error', message: 'コメントが見つかりません' };
}

// ============================================================
// 管理者：会員削除 (adminDeleteMember)
// ============================================================

/**
 * 会員を削除する（管理者のみ）
 * @param {Object} body - { adminToken, targetMemberId }
 */
function adminDeleteMember(body) {
  if (!verifyAdminToken(body.adminToken)) {
    return { status: 'error', message: '管理者権限がありません' };
  }

  var sheet = getSheet('members');
  var data = sheet.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === body.targetMemberId) {
      sheet.deleteRow(i + 1);
      // 対応するトークンも削除
      PropertiesService.getScriptProperties().deleteProperty('TOKEN_' + body.targetMemberId);
      return { status: 'ok', data: { message: '会員を削除しました' } };
    }
  }

  return { status: 'error', message: '会員が見つかりません' };
}

// ============================================================
// 管理者：全会員一覧取得 (adminGetAllMembers)
// ============================================================

/**
 * 全会員一覧を取得する（管理者のみ）
 * @param {Object} body - { adminToken }
 */
function adminGetAllMembers(body) {
  if (!verifyAdminToken(body.adminToken)) {
    return { status: 'error', message: '管理者権限がありません' };
  }

  var sheet = getSheet('members');
  var data = sheet.getDataRange().getValues();
  var members = [];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    members.push({
      memberId:   data[i][0],
      email:      data[i][1],
      // パスワード・saltは返さない
      nickname:   data[i][4],
      mustChangePassword: data[i][5] === true || String(data[i][5]).toLowerCase() === 'true',
      registeredAt: data[i][7]
    });
  }

  return { status: 'ok', data: { members: members } };
}

// ============================================================
// 管理者：全コメント一覧取得 (adminGetAllMessages)
// ============================================================

/**
 * 全コメント一覧を取得する（管理者のみ）
 * @param {Object} body - { adminToken }
 */
function adminGetAllMessages(body) {
  if (!verifyAdminToken(body.adminToken)) {
    return { status: 'error', message: '管理者権限がありません' };
  }

  var sheet = getSheet('messages');
  var data = sheet.getDataRange().getValues();
  var messages = [];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    messages.push({
      messageId:  data[i][0],
      eventId:    data[i][1],
      memberId:   data[i][2],
      nickname:   data[i][3],
      content:    data[i][4],
      postedAt:   data[i][5]
    });
  }

  // 投稿日時降順でソート
  messages.sort(function(a, b) {
    return String(b.postedAt).localeCompare(String(a.postedAt));
  });

  return { status: 'ok', data: { messages: messages } };
}
