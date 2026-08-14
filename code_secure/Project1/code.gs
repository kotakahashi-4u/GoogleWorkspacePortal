/**
 * ====================================================================
 *  Google Workspace Portal
 * 
 *  Copyright (c) 2026 4U, Inc. All Rights Reserved.
 *  Released under the MIT License.
 *  https://corp-4u.com/
 * ====================================================================
 */

// ==========================================
// DB Proxy API ヘルパークラス
// ==========================================

/**
 * Project 2 (DB Proxy API) と安全に通信を行うためのヘルパークラス
 */
class DbProxy {
  /**
   * 実行ユーザーのOAuthトークンを付与してAPIを呼び出す
   * @param {string} action - 実行するAPIのアクション名
   * @param {Object} [payload] - APIへ送信するデータ
   * @returns {any} APIから返却されたデータ
   */
  static call(action, payload) {
    const apiUrl = PropertiesService.getScriptProperties().getProperty('API_URL');
    if (!apiUrl) throw new Error("システムエラー: API_URLが設定されていません。プロジェクトの設定から追加してください。");

    const token = ScriptApp.getOAuthToken();
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { "Authorization": "Bearer " + token },
      payload: JSON.stringify({
        action: action,
        user_email: Session.getActiveUser().getEmail(),
        payload: payload || {}
      }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(apiUrl, options);
    if (response.getResponseCode() !== 200) {
      throw new Error("DB_API通信エラー: " + response.getContentText());
    }
    
    const resJson = JSON.parse(response.getContentText());
    if (!resJson.success) throw new Error(resJson.error);
    return resJson.data;
  }
}

// ==========================================
// キャッシュクリア・初期化処理
// ==========================================

/**
 * デバッグおよび強制リセット用：すべてのキャッシュをクリアする
 * ユーザー単位およびスクリプト全体（全社共通）のキャッシュを破棄します。
 */
function clearAllCachesForDebug() {
  const userCache = CacheService.getUserCache();
  userCache.remove('FULL_USER_CONFIG_CACHE');
  userCache.remove('USER_GROUPS_CACHE');

  const scriptCache = CacheService.getScriptCache();
  scriptCache.remove('MASTER_LINKS_CACHE');
  scriptCache.remove('CLOUD_SEARCH_SETTING_CACHE');
  scriptCache.remove('WF_MASTER_DATA_CACHE');
  scriptCache.remove('PORTAL_MASTER_DATA_CACHE');
}

/**
 * プロキシAPIからマスタデータを取得（キャッシュ対応）
 */
function fetchCachedPortalMasterData() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('PORTAL_MASTER_DATA_CACHE');
  if (cached) {
    try { return JSON.parse(cached); } catch(e){}
  }
  
  const data = DbProxy.call('getPortalMasterData');
  // 3時間(10800秒) キャッシュする
  try { cache.put('PORTAL_MASTER_DATA_CACHE', JSON.stringify(data), 10800); } catch(e){}
  return data;
}

// ==========================================
// 多言語対応（i18n）基盤
// ==========================================

/**
 * 実行ユーザーの言語設定を取得する
 * @returns {string} 'ja' または 'en'
 */
function getUserLocale() {
  try {
    const locale = Session.getActiveUserLocale();
    return (locale && locale.startsWith('ja')) ? 'ja' : 'en';
  } catch (e) {
    return 'ja'; // 取得できない場合は日本語をデフォルトとする
  }
}

/**
 * バックエンド用の多言語メッセージ辞書
 * ※フロントエンド側の翻訳は別途HTML側で実装します
 */
const BACKEND_MSG = {
  ja: {
    errDbInit: "データベース用のスプレッドシートへのアクセス、または初期化に失敗しました。\nスクリプトプロパティ「LINK_SPREADSHEET_ID」が正しく設定され、対象シートへのアクセス権限があることを確認してください。\n【エラー詳細】: ",
    errInvalidInput: "入力データが不正です。必須項目を確認してください。",
    errInvalidInputAssignee: "担当者が設定されていないステップがあります。",
    errCommentRequired: "却下・差戻しの場合はコメント（理由）を入力してください。",
    errUpdateConflict: "他のユーザーが更新中です。",
    errServerBusy: "サーバーが混み合っています。少し待ってから再試行してください。",
    errWorkflowFolderIdMissing: "システムエラー: 管理者に WORKFLOW_FOLDER_ID の設定を依頼してください。",
    errTicketNotFound: "チケットが見つかりません。",
    errRouteInfoNotFound: "経路情報が見つかりません。",
    errTaskNoPermission: "このタスクを実行する権限がありません。",
    errAppAlreadyCompleted: "この申請はすでに完了または却下されています。",
    errProcessFailed: "処理に失敗しました: ",
    errUnauthorizedGroup: "指定されたグループへのアクセスまたは投稿権限がありません。"
  },
  en: {
    errDbInit: "Failed to access or initialize the database spreadsheet.\nPlease ensure the script property 'LINK_SPREADSHEET_ID' is set correctly and you have access permissions.\n[Error Details]: ",
    errInvalidInput: "Invalid input data. Please check required fields.",
    errInvalidInputAssignee: "There are steps with no assignee set.",
    errCommentRequired: "A comment (reason) is required when rejecting or returning.",
    errUpdateConflict: "Another user is currently updating. Please try again.",
    errServerBusy: "Server is busy. Please try again shortly.",
    errWorkflowFolderIdMissing: "System Error: Please ask the administrator to set WORKFLOW_FOLDER_ID.",
    errTicketNotFound: "Ticket not found.",
    errRouteInfoNotFound: "Route information not found.",
    errTaskNoPermission: "You do not have permission to execute this task.",
    errAppAlreadyCompleted: "This application has already been completed or rejected.",
    errProcessFailed: "Process failed: ",
    errUnauthorizedGroup: "You do not have permission to access or post to the specified group."
  }
};

/**
 * 指定したキーの翻訳メッセージを取得する
 * @param {string} key - メッセージキー
 * @returns {string} 翻訳されたメッセージ
 */
function getMsg(key) {
  const lang = getUserLocale();
  return BACKEND_MSG[lang][key] || BACKEND_MSG['en'][key] || key;
}


// ==========================================
// 描画処理
// ==========================================

/**
 * Webアプリにアクセスがあった際の初期表示（エントリーポイント）
 * @param {Object} e - イベントオブジェクト
 * @returns {HtmlOutput} 評価済みのHTML
 */
function doGet(e) {
  const html = HtmlService.createTemplateFromFile('index').evaluate();
  html.setTitle('GoogleWorkspaceポータル');
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  // モバイル端末でのレスポンシブ対応
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1, user-scalable=no');
  
  return html;
}

/**
 * ポータル起動用の初期データを一括取得する
 * 通信オーバーヘッドを最小化するため、画面描画に必要な情報を1回の通信でまとめて返却します。
 * @returns {Object} ポータル初期化に必要なデータの集合
 */
function getInitialPortalData() {
  // API経由でマスタデータと状態を一括取得（キャッシュにより高速化）
  const masterData = fetchCachedPortalMasterData();
  
  if (masterData.isMaintenanceMode) {
    return { isMaintenanceMode: true, userLocale: getUserLocale() };
  }

  return {
    isMaintenanceMode: false,
    userLocale: getUserLocale(),
    masterLinks: masterData.masterLinks,
    userGroups: getUserGroups(),
    cloudSearchEnabled: masterData.cloudSearchEnabled,
    userConfig: getUserConfig(),
    archiveSchedule: masterData.archiveSchedule
  };
}

// ==========================================
// ユーザー設定（レイアウト・テーマ）の管理
// ==========================================

// ポータルの初期設定値定義
const DEFAULT_CONFIG = {
  themeColor: '#4285F4',
  layout: [
    { id: 'widget-workflow', type: 'workflow', title: '承認・確認待ちタスク', width: '25' },
    { id: 'widget-bulletin-1', type: 'bulletin', title: '社内掲示板', width: '50', groupEmail: '' },
    { id: 'widget-bcp', type: 'bcp', title: '防災・運行・緊急連絡', width: '25' },
    { id: 'widget-gmail', type: 'gmail', title: '未読メール', width: '25' },
    { id: 'widget-drive', type: 'drive', title: '最近のファイル', width: '25' },
    { id: 'widget-calendar', type: 'calendar', title: 'マイスケジュール', width: '50' }
  ],
  selectedLinks: [],
  bcpWeatherArea1: '120000', // 東京
  bcpWeatherArea2: '130000', // 大阪
  bcpTransitLinks: [
    { id: 'transit_jr_east', title: 'JR東日本', url: 'https://traininfo.jreast.co.jp/train_info/kanto.aspx', isCustom: false },
    { id: 'transit_tokyo_metro', title: '東京メトロ', url: 'https://www.tokyometro.jp/unkou/', isCustom: false }
  ]
};

/**
 * 実行ユーザーの個人設定を取得する
 * ユーザープロパティから取得後、未設定の項目にはデフォルト値を補完します。
 * 取得結果はキャッシュに保存し、次回以降の読み込みを高速化します。
 * @returns {Object} ユーザーごとの設定オブジェクト
 */
function getUserConfig() {
  const cache = CacheService.getUserCache();
  const cachedFullConfig = cache.get('FULL_USER_CONFIG_CACHE');
  if (cachedFullConfig) return JSON.parse(cachedFullConfig);

  const props = PropertiesService.getUserProperties();
  const configStr = props.getProperty('portalConfig');
  let config = configStr ? JSON.parse(configStr) : JSON.parse(JSON.stringify(DEFAULT_CONFIG));
  
  // 後方互換性および新規項目の補完
  if (!config.selectedLinks) config.selectedLinks = [];
  if (!config.bcpWeatherArea1) config.bcpWeatherArea1 = DEFAULT_CONFIG.bcpWeatherArea1;
  if (!config.bcpWeatherArea2) config.bcpWeatherArea2 = DEFAULT_CONFIG.bcpWeatherArea2;
  if (!config.bcpTransitLinks) config.bcpTransitLinks = DEFAULT_CONFIG.bcpTransitLinks;

  // 掲示板設定のマイグレーション
  if (config.bulletinGroup !== undefined) {
    const bWidget = config.layout.find(w => w.type === 'bulletin');
    if (bWidget && bWidget.groupEmail === undefined) bWidget.groupEmail = config.bulletinGroup;
    delete config.bulletinGroup;
  }

  // 必須ウィジェットの欠損を補完するロジック
  let needsSave = false;
  const defaultTypes = DEFAULT_CONFIG.layout.map(w => w.type);
  const currentTypes = config.layout.map(w => w.type);

  const initialLength = config.layout.length;
  config.layout = config.layout.filter(w => w.type === 'bulletin' || defaultTypes.includes(w.type));
  if (config.layout.length !== initialLength) needsSave = true;

  DEFAULT_CONFIG.layout.forEach(defW => {
    if (defW.type !== 'bulletin' && !currentTypes.includes(defW.type)) {
      config.layout.push(JSON.parse(JSON.stringify(defW)));
      needsSave = true;
    }
  });

  if (needsSave) {
    props.setProperty('portalConfig', JSON.stringify(config));
  }
  
  config.userEmail = Session.getActiveUser().getEmail();
  config.isAdmin = isAdminUser();

  // ユーザーのGoogleカレンダーリストを取得し、埋め込み用パラメータを生成
  let calParams = '';
  try {
    if (typeof Calendar !== 'undefined') {
      const calendarList = Calendar.CalendarList.list().items;
      if (calendarList) {
        calendarList.forEach(cal => {
          if (cal.selected) calParams += `&src=${encodeURIComponent(cal.id)}&color=${encodeURIComponent(cal.backgroundColor)}`;
        });
      }
    } else {
      calParams = `&src=${encodeURIComponent(config.userEmail)}`;
    }
  } catch(e) {
    calParams = `&src=${encodeURIComponent(config.userEmail)}`;
  }
  
  config.calendarParams = calParams;

  try { cache.put('FULL_USER_CONFIG_CACHE', JSON.stringify(config), 21600); } catch(e){}
  
  return config;
}

/**
 * ユーザーの個人設定を保存する
 * @param {Object} config - 保存する設定オブジェクト
 * @returns {boolean} 処理が成功したかどうか
 */
function saveUserConfig(config) {
  try {
    const configStr = JSON.stringify(config);
    const props = PropertiesService.getUserProperties();
    props.setProperty('portalConfig', configStr);
    
    // 次回アクセス時に再計算させるため、キャッシュを破棄
    try { CacheService.getUserCache().remove('FULL_USER_CONFIG_CACHE'); } catch(e){}
    
    return true;
  } catch (error) {
    throw error;
  }
}

// ==========================================
// 各種データ取得API
// ==========================================

/**
 * 実行ユーザーのGmail受信トレイから最新メールを取得する
 * 無限スクロール（ページネーション）に対応
 * @param {number} [offset=0] - 取得開始位置
 * @param {number} [limit=30] - 取得件数上限
 * @returns {Array<Object>} メールのスレッド情報の配列
 */
function getInboxEmails(offset = 0, limit = 30) {
  try {
    const threads = GmailApp.search('label:inbox', offset, limit);
    return threads.map(t => {
      const msg = t.getMessages()[t.getMessageCount() - 1];
      let bodyText = msg.getPlainBody() || '';
      if (bodyText.length > 250) bodyText = bodyText.substring(0, 250) + '...';

      let rawFrom = msg.getFrom();
      let fromName = rawFrom.includes('<') ? (rawFrom.split('<')[0].trim() || rawFrom.match(/<([^>]+)>/)[1]) : rawFrom;

      const decThreadId = BigInt('0x' + t.getId()).toString();
      const decMsgId = BigInt('0x' + msg.getId()).toString();

      return {
        id: msg.getId(), 
        subject: msg.getSubject().trim() || '(件名なし)',
        from: fromName,
        date: msg.getDate().toLocaleDateString('ja-JP', { month: 'short', day: 'numeric' }),
        link: `https://mail.google.com/mail/u/0/?fs=1&th=%23thread-f:${decThreadId}%7Cmsg-f:${decMsgId}&search=all&tf=cv`,
        isUnread: t.isUnread(),
        bodyPreview: bodyText
      };
    });
  } catch (e) {
    console.error("Gmail取得エラー:", e);
    return [];
  }
}

/**
 * Google Driveからファイルやフォルダの一覧を取得する
 * @param {string} tabType - 'recent' (最近), 'mydrive' (マイドライブ), 'shared' (共有ドライブ)
 * @param {string} folderId - 展開する対象フォルダのID (nullの場合はルート)
 * @returns {Object} 取得したアイテムリストとカレント階層情報のオブジェクト
 */
function getDriveItems(tabType, folderId) {
  try {
    const MAX_FOLDERS = 50; 
    const MAX_FILES = 50;   
    let items = [];
    let parentId = null;
    
    // 共有ドライブルートの取得 (Drive API v2 必須)
    if (tabType === 'shared' && !folderId) {
      if (typeof Drive === 'undefined') return { items: [], error: '「サービス ＋」から「Google Drive API」を追加してください。' };
      const driveList = Drive.Drives.list({ pageSize: 50 });
      const sharedDrives = driveList.drives || driveList.items || [];
      sharedDrives.forEach(d => { items.push({ id: d.id, name: d.name, type: 'drive', mimeType: '', url: '', date: '' }); });
      return { items: items, parentId: null, currentName: '共有ドライブ' };
    }
    
    // 最近使用したアイテムの取得
    if (tabType === 'recent' && !folderId) {
      const files = DriveApp.getFiles(); 
      while (files.hasNext() && items.length < MAX_FILES) {
        const f = files.next();
        items.push({ id: f.getId(), name: f.getName(), type: 'file', mimeType: f.getMimeType(), url: f.getUrl(), date: Utilities.formatDate(f.getLastUpdated(), "JST", "MM/dd") });
      }
      return { items: items, parentId: null, currentName: '最近使用したアイテム' };
    }
    
    let folder = !folderId ? DriveApp.getRootFolder() : DriveApp.getFolderById(folderId);
    
    // パンくず（戻る）用の親フォルダIDを取得
    if (folderId) {
      const parents = folder.getParents();
      if (parents.hasNext()) parentId = parents.next().getId();
      else if (tabType === 'shared') parentId = 'ROOT_SHARED'; 
    }
    
    let folderCount = 0;
    const folders = folder.getFolders();
    while (folders.hasNext() && folderCount < MAX_FOLDERS) {
       const f = folders.next();
       items.push({ id: f.getId(), name: f.getName(), type: 'folder', mimeType: '', url: f.getUrl(), date: Utilities.formatDate(f.getLastUpdated(), "JST", "MM/dd") });
       folderCount++;
    }
    
    let fileCount = 0;
    const files = folder.getFiles();
    while (files.hasNext() && fileCount < MAX_FILES) {
       const f = files.next();
       items.push({ id: f.getId(), name: f.getName(), type: 'file', mimeType: f.getMimeType(), url: f.getUrl(), date: Utilities.formatDate(f.getLastUpdated(), "JST", "MM/dd") });
       fileCount++;
    }
    return { items: items, parentId: parentId, currentName: folder.getName() };
  } catch (e) {
    return { items: [], error: e.toString() };
  }
}

// ==========================================
// 社内掲示板（Googleグループ連携）
// ==========================================

/**
 * 実行ユーザーが所属しているGoogleグループのメールアドレス一覧を取得する
 * @returns {Array} グループアドレスの配列
 */
function getUserGroups() {
  const cache = CacheService.getUserCache();
  const cached = cache.get('USER_GROUPS_CACHE');
  if (cached) return JSON.parse(cached);

  try {
    const groups = GroupsApp.getGroups();
    const groupEmails = groups.map(g => g.getEmail());
    try { cache.put('USER_GROUPS_CACHE', JSON.stringify(groupEmails), 21600); } catch(e){}
    return groupEmails;
  } catch(e) {
    return [];
  }
}

/**
 * 指定されたグループのメンバーのメールアドレス一覧を取得する
 * 権限がない場合はエラーをスローします。
 * @param {string} groupEmail - 対象のグループメールアドレス
 * @returns {Array} メンバーアドレスの配列
 */
function getGroupMembers(groupEmail) {
  if (!getUserGroups().includes(groupEmail)) {
    throw new Error(getMsg('errUnauthorizedGroup'));
  }
  try {
    const group = GroupsApp.getGroupByEmail(groupEmail);
    if (!group) throw new Error("グループが存在しないか、アクセスできません。");
    const users = group.getUsers();
    return users.map(u => u.getEmail());
  } catch(e) {
    throw new Error("メンバー取得エラー: " + e.message);
  }
}

/**
 * 自分が所属する全グループのメンバーを重複排除して一括取得する
 * @returns {Array} ユニークなメンバーアドレスの配列
 */
function getAllMyGroupMembers() {
  try {
    const myGroups = GroupsApp.getGroups();
    const allMembers = new Set();
    allMembers.add(Session.getActiveUser().getEmail());

    myGroups.forEach(group => {
      try {
        const users = group.getUsers();
        users.forEach(u => allMembers.add(u.getEmail()));
      } catch(e) {
        // 権限がないグループはエラーを出さずにスキップ
      }
    });
    return Array.from(allMembers);
  } catch(e) {
    return [Session.getActiveUser().getEmail()];
  }
}

/**
 * 指定されたグループ宛のメーリングリスト投稿を取得する
 * 無限スクロール（ページネーション）に対応
 * @param {string} groupEmail - 対象のグループメールアドレス
 * @param {number} [offset=0] - 取得開始位置
 * @param {number} [limit=15] - 取得件数上限
 * @returns {Array|Object} 投稿スレッドの配列、またはエラーオブジェクト
 */
function getBulletinPosts(groupEmail, offset = 0, limit = 15) {
  if (!groupEmail) return { error: 'not_configured' };
  if (!getUserGroups().includes(groupEmail)) return { error: 'unauthorized' };
  try {
    const threads = GmailApp.search(`to:${groupEmail} OR list:${groupEmail}`, offset, limit);
    return threads.map(t => {
      const msgs = t.getMessages();
      const firstMsg = msgs[0];
      const lastMsg = msgs[msgs.length - 1];
      
      let bodyText = firstMsg.getPlainBody() || '';
      if (bodyText.length > 200) bodyText = bodyText.substring(0, 200) + '...';
      
      let rawFrom = firstMsg.getFrom();
      let fromName = rawFrom.includes('<') ? (rawFrom.split('<')[0].trim() || rawFrom.match(/<([^>]+)>/)[1]) : rawFrom;
      let subject = firstMsg.getSubject().trim();
      
      // メーリングリストのプレフィックスを除去
      subject = subject.replace(/^\[.*?\]\s*/, '');

      const threadMsgs = msgs.map(m => {
        let rFrom = m.getFrom();
        let fName = rFrom.includes('<') ? (rFrom.split('<')[0].trim() || rFrom.match(/<([^>]+)>/)[1]) : rFrom;
        
        // インライン画像以外の通常の添付ファイルをリスト化
        const stdAtts = m.getAttachments().map(a => ({
          name: a.getName(),
          size: a.getSize() >= 1048576 ? (a.getSize() / 1048576).toFixed(1) + ' MB' : (a.getSize() / 1024).toFixed(1) + ' KB'
        }));

        let bodyHtml = m.getBody();
        if (bodyHtml) {
          try {
            const attachments = m.getAttachments({includeInlineImages: true});
            attachments.forEach(att => {
              const mime = att.getContentType();
              if (mime.startsWith('image/')) {
                const b64 = Utilities.base64Encode(att.getBytes());
                const dataUrl = `data:${mime};base64,${b64}`;
                const name = att.getName();
                let replaced = false;
                
                // Gmail仕様におけるインライン画像のcid置換処理
                if (name) {
                  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                  const regex = new RegExp(`<img[^>]+alt=["']?${escaped}["']?[^>]*>`, 'i');
                  const match = bodyHtml.match(regex);
                  if (match) {
                    const newImgTag = match[0].replace(/src=["']?cid:[^"'\s>]+["']?/, `src="${dataUrl}"`);
                    bodyHtml = bodyHtml.replace(match[0], newImgTag);
                    replaced = true;
                  }
                }
                
                // フォールバック: altでマッチしない場合は残っているcidを置換
                if (!replaced && bodyHtml.includes("cid:")) {
                  bodyHtml = bodyHtml.replace(/src=["']?cid:[^"'\s>]+["']?/, `src="${dataUrl}"`);
                }
              }
            });
          } catch(e) {
            console.error("インライン画像の変換に失敗:", e);
          }
        } else {
          bodyHtml = m.getPlainBody().replace(/\n/g, '<br>');
        }

        return {
          id: m.getId(),
          from: fName,
          rawFrom: rFrom,
          date: m.getDate().toLocaleDateString('ja-JP', { month: 'short', day: 'numeric', hour: '2-digit', minute:'2-digit' }),
          htmlBody: bodyHtml,
          attachments: stdAtts
        };
      });

      return {
        id: lastMsg.getId(), 
        threadId: t.getId(),
        subject: subject,
        from: fromName,
        date: firstMsg.getDate().toLocaleDateString('ja-JP', { month: 'short', day: 'numeric', hour: '2-digit', minute:'2-digit' }),
        bodyPreview: bodyText,
        messages: threadMsgs,
        rawFrom: rawFrom, 
        replyCount: msgs.length - 1
      };
    });
  } catch(e) {
    console.error("掲示板取得エラー:", e);
    return { error: 'fetch_failed' };
  }
}

/**
 * 掲示板（グループ宛）へ新規投稿を送信する
 * @param {string} groupEmail - 宛先グループアドレス
 * @param {string} subject - 件名
 * @param {string} body - プレーンテキスト本文
 * @param {string} htmlBody - HTML本文
 * @param {Array} attachments - 添付ファイル配列
 * @returns {boolean} 成功時
 */
function postBulletin(groupEmail, subject, body, htmlBody = '', attachments = []) {
  if (!groupEmail) throw new Error("グループが設定されていません");
  if (!getUserGroups().includes(groupEmail)) throw new Error(getMsg('errUnauthorizedGroup'));
  try {
    let options = {};
    if (htmlBody) options.htmlBody = htmlBody;
    if (attachments.length > 0) {
      options.attachments = attachments.map(att => 
        Utilities.newBlob(Utilities.base64Decode(att.data), att.mimeType, att.name)
      );
    }
    GmailApp.sendEmail(groupEmail, subject, body, options);
    return true;
  } catch(e) {
    throw new Error("投稿に失敗しました");
  }
}

/**
 * 掲示板の特定スレッドに対して返信する
 * @param {string} messageId - 返信元のメッセージID
 * @param {string} body - プレーンテキスト本文
 * @param {string} htmlBody - HTML本文
 * @param {string} replyType - 'replyAll' または 'reply'
 * @param {string} authorEmail - 投稿者への直接返信用アドレス
 * @param {string} subject - 元の件名
 * @param {Array} attachments - 添付ファイル配列
 * @returns {boolean} 成功時
 */
function replyBulletin(messageId, body, htmlBody = '', replyType = 'replyAll', authorEmail = '', subject = '', attachments = []) {
  try {
    const msg = GmailApp.getMessageById(messageId);
    let options = {};
    if (htmlBody) options.htmlBody = htmlBody;
    if (attachments.length > 0) {
      options.attachments = attachments.map(att => 
        Utilities.newBlob(Utilities.base64Decode(att.data), att.mimeType, att.name)
      );
    }

    if (replyType === 'reply') {
      if (!authorEmail) throw new Error("宛先アドレスが不明です");
      GmailApp.sendEmail(authorEmail, "Re: " + subject, body, options);
    } else {
      msg.replyAll(body, options);
    }
    return true;
  } catch(e) {
    console.error(e);
    throw new Error("送信に失敗しました");
  }
}

/**
 * メッセージから特定の添付ファイルのバイナリ(Base64)を取得する
 * @param {string} messageId - メッセージID
 * @param {string} attachmentName - 取得したいファイル名
 * @returns {Object} ファイルのMIMEタイプとBase64データ
 */
function getAttachmentData(messageId, attachmentName) {
  try {
    const msg = GmailApp.getMessageById(messageId);
    const attachments = msg.getAttachments();
    const target = attachments.find(a => a.getName() === attachmentName);
    if (!target) throw new Error("ファイルが見つかりません");
    
    return {
      mimeType: target.getContentType(),
      data: Utilities.base64Encode(target.getBytes()),
      name: target.getName()
    };
  } catch(e) {
    throw new Error("ファイルの取得に失敗しました");
  }
}

// ==========================================
// マスタ管理関連 (スクリプトプロパティ ＆ スプレッドシート)
// ==========================================

/**
 * 実行ユーザーがシステム管理者かどうかを判定する
 * @returns {boolean} 管理者の場合はtrue
 */
function isAdminUser() {
  try {
    return DbProxy.call('isAdminUser');
  } catch (e) {
    return false;
  }
}

// ==========================================
// ワークフロー管理API
// ==========================================

/**
 * 自分が閲覧または編集可能な承認経路（ルート）の一覧を取得する
 * @returns {Array<Object>} 利用可能なルート情報の配列
 */
function getWorkflowRoutes() {
  const groups = getUserGroups();
  return DbProxy.call('getWorkflowRoutes', { groups: groups });
}

/**
 * 承認経路（マスタ）を新規作成または更新する
 * @param {Array<Object>} routes - 保存対象のルート配列
 * @param {Array<string>} [deletedIds=[]] - 削除対象のルートID配列
 * @returns {boolean} 成功時 true
 */
function saveWorkflowRoutes(routes, deletedIds = []) {
  return DbProxy.call('saveWorkflowRoutes', { routes: routes, deletedIds: deletedIds });
}

/**
 * 申請画面用に利用可能な（アクティブな）ルートのみを取得する
 * @returns {Array<Object>} アクティブなルートの配列
 */
function getAvailableRoutes() {
  const routes = getWorkflowRoutes();
  return routes.filter(r => r.isActive);
}

// ==========================================
// ワークフロー アクション・タスクAPI
// ==========================================

/**
 * 新たにワークフロー申請（チケット起票）を実行する
 * @param {Object} payload - 申請内容や経路情報を含むデータオブジェクト
 * @returns {boolean} 成功時 true
 */
function submitWorkflowApplication(payload) {
  // ユーザーのプライベートなDriveファイルが選択された場合、API側(管理者)からはアクセスできないため、
  // ユーザー権限で動作するこちら側でBase64エンコードし、添付ファイルとしてAPIに送信する
  if (payload.driveFileIds && payload.driveFileIds.length > 0) {
    if (!payload.attachments) payload.attachments = [];
    payload.driveFileIds.forEach(id => {
      const file = DriveApp.getFileById(id);
      payload.attachments.push({
        name: file.getName(),
        mimeType: file.getMimeType(),
        data: Utilities.base64Encode(file.getBlob().getBytes())
      });
    });
    delete payload.driveFileIds;
  }
  return DbProxy.call('submitWorkflowApplication', payload);
}

/**
 * 実行ユーザーが過去に起票した申請履歴（マイチケット）を取得する
 * @returns {Array<Object>} チケット情報の配列
 */
function getMyTickets() {
  return DbProxy.call('getMyTickets');
}

/**
 * 実行ユーザーが担当者となっている承認・供覧待ちタスクを取得する
 * @returns {Array<Object>} タスク情報の配列
 */
function getPendingTasks() {
  const groups = getUserGroups();
  return DbProxy.call('getPendingTasks', { groups: groups });
}

/**
 * 対象のチケットに対するアクション（承認、差戻、却下、確認）を実行する
 * @param {string} ticketId - チケットID
 * @param {number} stepOrder - アクションを実行するステップの順番
 * @param {string} actionType - アクション種別
 * @param {string} comment - コメント内容
 * @returns {boolean} 成功時 true
 */
function executeWorkflowAction(ticketId, stepOrder, actionType, comment) {
  const userGroups = getUserGroups();
  return DbProxy.call('executeWorkflowAction', {
    ticketId: ticketId,
    stepOrder: stepOrder,
    actionType: actionType,
    comment: comment,
    groups: userGroups
  });
}

// ==========================================
// BCP（防災・緊急連絡）API
// ==========================================

function getBcpData(area1, area2) {
  const result = {
    weather1: { text: "取得できませんでした", reportDatetime: "", name: "" },
    weather2: { text: "取得できませんでした", reportDatetime: "", name: "" },
    emergencyMessage: ""
  };

  // 1. プロキシから緊急メッセージを取得（キャッシュから即時返却）
  try {
    const masterData = fetchCachedPortalMasterData();
    result.emergencyMessage = masterData.emergencyMessage;
  } catch(e) {
    console.error("緊急メッセージ取得エラー:", e);
  }

  // 2. 気象庁APIから天気情報を取得 (外部APIのためProject 1で直接実行)
  const code1 = area1 || "130000";
  const code2 = area2 || "270000";

  try {
    const fetchWeather = (code) => {
      const res = UrlFetchApp.fetch(`https://www.jma.go.jp/bosai/forecast/data/overview_forecast/${code}.json`, { muteHttpExceptions: true });
      if (res.getResponseCode() === 200) {
        const data = JSON.parse(res.getContentText());
        return {
          text: data.text || "現在、気象情報はありません。",
          reportDatetime: Utilities.formatDate(new Date(data.reportDatetime), "JST", "MM/dd HH:mm 発表"),
          name: data.targetArea || ""
        };
      }
      return null;
    };

    const w1 = fetchWeather(code1);
    if (w1) result.weather1 = w1;

    if (code1 === code2) {
      result.weather2 = result.weather1;
    } else {
      const w2 = fetchWeather(code2);
      if (w2) result.weather2 = w2;
    }
  } catch (e) {
    console.error("気象情報の取得エラー:", e);
  }

  return result;
}

/**
 * 管理者が入力した全社向けの緊急アラートメッセージを保存する
 * @param {string} message - メッセージ内容
 * @returns {boolean} 成功時
 */
function saveEmergencyMessage(message) {
  if (!isAdminUser()) throw new Error("権限がありません");
  const res = DbProxy.call('saveEmergencyMessage', { message: message });
  CacheService.getScriptCache().remove('PORTAL_MASTER_DATA_CACHE');
  return res;
}

// =========================================================================
// フルスクリーンカンバン ＆ Google Tasks ハイブリッド同期 バックエンド
// =========================================================================

/**
 * すべてのカンバンタスクをAPI経由で取得し、個人のGoogle Tasksと自動同期を行う
 * 外部で作成されたGoogle Tasksはポータル側へ逆同期して一括インポートします。
 * @returns {Array<Object>} タスクデータの配列
 */
function getKanbanTasks() {
  const tasks = DbProxy.call('getKanbanTasksDB');
  const currentUser = Session.getActiveUser().getEmail();

  // APIから取得したタスクのGoogle Task IDを保持しておく
  const dbGoogleTaskIds = new Set();
  if (tasks && tasks.length > 0) {
    tasks.forEach(t => {
      if (t.googleTaskId) dbGoogleTaskIds.add(t.googleTaskId);
    });
  }

  let googleTasksMap = {};
  let importedTasks = [];
  try {
    let pageToken;
    do {
      const response = Tasks.Tasks.list('@default', {
        maxResults: 100, showCompleted: true, showHidden: true, showDeleted: true, pageToken: pageToken
      });
      if (response.items) {
        response.items.forEach(t => { 
          // 削除済みの場合は強制的に 'completed' 扱いとする
          googleTasksMap[t.id] = t.deleted ? 'completed' : t.status; 
          
          // DBに存在しない、かつ未完了・未削除のタスクを逆同期の対象とする
          if (!dbGoogleTaskIds.has(t.id) && t.status !== 'completed' && !t.deleted) {
            importedTasks.push(t);
          }
        });
      }
      pageToken = response.nextPageToken;
    } while (pageToken);
  } catch(e) {
    console.error("Google Tasks取得エラー:", e);
  }

  if (tasks && tasks.length > 0) {
    tasks.forEach(task => {
      if (task.assignee === currentUser && task.status !== 'done') {
        if (!task.googleTaskId) {
          try {
            let gTaskOpts = { title: task.title, notes: "ポータルのカンバンと同期されています。" };
            if (task.dueDate) gTaskOpts.due = task.dueDate + 'T00:00:00.000Z';
            const gTask = Tasks.Tasks.insert(gTaskOpts, '@default');
            task.googleTaskId = gTask.id;
            
            DbProxy.call('editKanbanTaskDB', { 
              taskId: task.id, title: task.title, assignee: task.assignee, 
              priority: task.priority, dueDate: task.dueDate, scope: task.scope, googleTaskId: gTask.id 
            });
          } catch (e) {}
        } else {
          try {
            const gTaskStatus = googleTasksMap[task.googleTaskId];
            // APIから返ってこない(undefined)場合も、Google Tasks側で完全に消去されたとみなし「完了」として同期する
            if (gTaskStatus === 'completed' || gTaskStatus === undefined) {
              task.status = 'done';
              DbProxy.call('updateKanbanTaskStatusDB', { taskId: task.id, newStatus: 'done' });
            }
          } catch (e) {}
        }
      }
    });
  }

  // 外部で作成されたGoogle Tasksを一括でDBへ新規追加（逆同期）
  if (importedTasks.length > 0) {
    const mappedTasks = importedTasks.map(gTask => {
      let dueDateStr = '';
      if (gTask.due) dueDateStr = Utilities.formatDate(new Date(gTask.due), "JST", "yyyy-MM-dd");
      return {
        title: gTask.title || '(無題)',
        assignee: currentUser,
        status: 'todo',
        scope: 'PERSONAL',
        priority: 'Medium',
        dueDate: dueDateStr,
        googleTaskId: gTask.id
      };
    });
    // O(1) 通信で一括インポート
    const updatedTasks = DbProxy.call('importKanbanTasksDB', { newTasks: mappedTasks });
    return updatedTasks;
  }

  return tasks || [];
}

/**
 * カンバンに新たなタスクを追加し、自分が担当者の場合はGoogle Tasksへ起票する
 * @param {string} title - タスク名
 * @param {string} assignee - 担当者のメールアドレス
 * @param {string} status - 初期ステータス
 * @param {string} scope - タスクの紐づくスコープ
 * @param {string} priority - 優先度
 * @param {string} dueDate - 期限
 * @returns {Array<Object>} 更新後のタスク一覧
 */
function addKanbanTask(title, assignee, status, scope, priority, dueDate) {
  let googleTaskId = '';
  const currentUser = Session.getActiveUser().getEmail();

  if (assignee === currentUser) {
    try {
      let gTaskOpts = { title: title, notes: "ポータルのカンバンと同期されています。" };
      if (dueDate) gTaskOpts.due = dueDate + 'T00:00:00.000Z';
      const gTask = Tasks.Tasks.insert(gTaskOpts, '@default');
      googleTaskId = gTask.id;
    } catch (e) {}
  }

  DbProxy.call('addKanbanTaskDB', {
    title: title, assignee: assignee, status: status, 
    scope: scope, priority: priority, dueDate: dueDate, googleTaskId: googleTaskId
  });

  return getKanbanTasks();
}

/**
 * カンバンタスクのステータスを更新し、自分が担当者の場合はGoogle Tasksと同期する
 * @param {string} taskId - 更新対象のタスクID
 * @param {string} newStatus - 変更後のステータス
 * @returns {Array<Object>} 更新後のタスク一覧
 */
function updateKanbanTaskStatus(taskId, newStatus) {
  const currentUser = Session.getActiveUser().getEmail();

  const updatedTasks = DbProxy.call('updateKanbanTaskStatusDB', { taskId: taskId, newStatus: newStatus });
  
  const targetTask = updatedTasks.find(t => t.id === taskId);
  if (targetTask && targetTask.assignee === currentUser && targetTask.googleTaskId) {
    try {
      const gTask = Tasks.Tasks.get('@default', targetTask.googleTaskId);
      if (newStatus === 'done') gTask.status = 'completed';
      else gTask.status = 'needsAction';
      Tasks.Tasks.patch(gTask, '@default', targetTask.googleTaskId);
    } catch (e) {}
  }

  return updatedTasks;
}

/**
 * カンバンタスクの属性を編集し、担当者変更に応じてGoogle Tasksを再設定する
 * @param {string} taskId - 編集対象のタスクID
 * @param {string} title - 新しいタスク名
 * @param {string} assignee - 新しい担当者
 * @param {string} priority - 新しい優先度
 * @param {string} dueDate - 新しい期限
 * @param {string} scope - 新しいスコープ
 * @returns {Array<Object>} 更新後のタスク一覧
 */
function editKanbanTask(taskId, title, assignee, priority, dueDate, scope) {
  const currentUser = Session.getActiveUser().getEmail();
  
  const targetTask = DbProxy.call('getKanbanTaskByIdDB', { taskId: taskId });
  if (!targetTask) return getKanbanTasks();

  const oldAssignee = targetTask.assignee;
  let googleTaskId = targetTask.googleTaskId;

  if (assignee === currentUser) {
     if (googleTaskId) {
       try {
         const gTask = Tasks.Tasks.get('@default', googleTaskId);
         gTask.title = title;
         gTask.due = dueDate ? (dueDate + 'T00:00:00.000Z') : null;
         Tasks.Tasks.patch(gTask, '@default', googleTaskId);
       } catch(e) {}
     } else {
       try {
         let gTaskOpts = { title: title, notes: "ポータルのカンバンと同期されています。" };
         if (dueDate) gTaskOpts.due = dueDate + 'T00:00:00.000Z';
         const gTask = Tasks.Tasks.insert(gTaskOpts, '@default');
         googleTaskId = gTask.id;
       } catch (e) {}
     }
  } else if (oldAssignee === currentUser && googleTaskId) {
     try { Tasks.Tasks.remove('@default', googleTaskId); } catch(e) {}
     googleTaskId = ''; 
  }

  DbProxy.call('editKanbanTaskDB', {
    taskId: taskId, title: title, assignee: assignee, 
    priority: priority, dueDate: dueDate, scope: scope, googleTaskId: googleTaskId
  });

  return getKanbanTasks();
}

/**
 * カンバンタスクを完全に削除し、関連するGoogle Tasksも破棄する
 * @param {string} taskId - 削除対象のタスクID
 * @returns {Array<Object>} 削除後のタスク一覧
 */
function deleteKanbanTask(taskId) {
  const currentUser = Session.getActiveUser().getEmail();

  const targetTask = DbProxy.call('getKanbanTaskByIdDB', { taskId: taskId });

  if (targetTask && targetTask.assignee === currentUser && targetTask.googleTaskId) {
    try { Tasks.Tasks.remove('@default', targetTask.googleTaskId); } catch (e) {}
  }

  DbProxy.call('deleteKanbanTaskDB', { taskId: taskId });
  return getKanbanTasks();
}

// ==========================================
// メンテナンスモード・アーカイブ管理
// ==========================================

/**
 * メンテナンスモードが有効かどうかを取得する
 * @returns {boolean} 有効な場合はtrue
 */
function getMaintenanceMode() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('MAINTENANCE_MODE') === 'true';
}

/**
 * メンテナンスモードの有効/無効を設定する
 * @param {boolean} isEnable - 設定値
 * @returns {boolean} 成功時
 */
function setMaintenanceMode(isEnable) {
  if (!isAdminUser()) throw new Error("権限がありません。");
  const props = PropertiesService.getScriptProperties();
  props.setProperty('MAINTENANCE_MODE', isEnable ? 'true' : 'false');
  return true;
}

/**
 * 管理画面等から手動で呼び出されるアーカイブの一括実行処理
 * @returns {string} 処理結果メッセージ
 */
function runManualArchiveBatch() {
  if (!isAdminUser()) throw new Error(getMsg('errTaskNoPermission'));
  return DbProxy.call('runManualArchiveBatch');
}

/**
 * 管理画面の設定（検索、アーカイブ、リンクマスタ）を一括で保存する
 * @param {boolean} searchEnabled - Cloud Searchの有効状態
 * @param {Object} archiveConfig - アーカイブ構成
 * @param {Array<Object>} newLinks - 新しいリンクマスタ
 * @returns {Array<Object>} 保存後のリンクマスタ配列
 */
function saveAllAdminSettings(searchEnabled, archiveConfig, newLinks) {
  if (!isAdminUser()) throw new Error(getMsg('errTaskNoPermission'));
  try {
    const links = DbProxy.call('saveAllAdminSettings', { searchEnabled: searchEnabled, archiveConfig: archiveConfig, newLinks: newLinks });
    CacheService.getScriptCache().remove('PORTAL_MASTER_DATA_CACHE');
    return links; 
  } catch (e) {
    throw new Error("バックエンド処理エラー: " + e.message);
  }
}

/**
 * 過去のアーカイブ済み申請履歴を検索して取得する
 * 無限スクロール（ページネーション）に対応
 * @param {string} query - 検索テキスト
 * @param {string} dateFrom - 検索開始日
 * @param {string} dateTo - 検索終了日
 * @param {number} [limit=50] - 取得件数上限
 * @param {number} [offset=0] - 取得開始位置
 * @returns {Array<Object>} 検索に合致したアーカイブチケットの配列
 */
function searchArchivedMyTickets(query, dateFrom, dateTo, limit = 50, offset = 0) {
  return DbProxy.call('searchArchivedMyTickets', { query: query, dateFrom: dateFrom, dateTo: dateTo, limit: limit, offset: offset });
}

// ==========================================
// グループカレンダーAPI
// ==========================================

/**
 * ドメイン内の会議室（リソース）一覧を取得する
 * Admin SDKを使用して利用可能な全リソースの情報を名前順で返します。
 * @returns {Array|Object} リソース情報の配列またはエラー
 */
function getDomainResources() {
  try {
    let allResources = [];
    let pageToken;
    do {
      const response = AdminDirectory.Resources.Calendars.list('my_customer', {
        maxResults: 100,
        pageToken: pageToken
      });
      if (response.items) {
        response.items.forEach(r => {
          allResources.push({
            email: r.resourceEmail,
            name: r.resourceName,
            description: r.resourceDescription || "",
            capacity: r.resourceCapacity || ""
          });
        });
      }
      pageToken = response.nextPageToken;
    } while (pageToken);
    
    allResources.sort((a, b) => a.name.localeCompare(b.name));
    return allResources;
  } catch (e) {
    console.error("リソース取得エラー:", e.message);
    return { error: true, message: e.message };
  }
}

/**
 * 指定グループメンバーおよび追加された会議室のカレンダー予定を取得する
 * Calendar API (Advanced Service) を使用して裏側からスケジュールを取得します。
 * @param {string} groupEmail - 対象のGoogleグループメールアドレス
 * @param {string} dateStr - 取得対象の日付 (YYYY/MM/DD形式)
 * @param {number} days - 取得する日数 (1 or 3)
 * @param {Array} addedResources - 追加された会議室配列 [{email, name}]
 * @returns {Array} 人および会議室のスケジュールデータ配列
 */
function getGroupCalendarEvents(groupEmail, dateStr, days, addedResources = []) {
  const targetDate = dateStr ? new Date(dateStr) : new Date();
  targetDate.setHours(0, 0, 0, 0);
  const nextDate = new Date(targetDate);
  const displayDays = days || 1;
  nextDate.setDate(nextDate.getDate() + displayDays); 

  let members = [];
  const myEmail = Session.getActiveUser().getEmail();
  
  if (!groupEmail || groupEmail === 'PERSONAL') {
    members = [myEmail];
  } else {
    try {
      members = getGroupMembers(groupEmail);
      const myIndex = members.indexOf(myEmail);
      if (myIndex > -1) members.splice(myIndex, 1);
      
      members.sort((a, b) => a.localeCompare(b));
      members.unshift(myEmail);
    } catch(e) {
      members = [myEmail];
    }
  }

  const colors = ["#4285F4", "#EA4335", "#FBBC05", "#34A853"];
  const results = [];

  // ユーザーの予定取得
  members.forEach((email, index) => {
    const userEvents = [];
    let displayName = email.split('@')[0];
    let hasAccess = true; 
    
    try {
      if (typeof AdminDirectory !== 'undefined') {
        const userProfile = AdminDirectory.Users.get(email);
        if (userProfile && userProfile.name && userProfile.name.fullName) {
          displayName = userProfile.name.fullName;
        }
      }
    } catch(e) {}

    try {
      if (typeof Calendar !== 'undefined') {
        const response = Calendar.Events.list(email, {
          timeMin: targetDate.toISOString(),
          timeMax: nextDate.toISOString(),
          singleEvents: true,
          orderBy: 'startTime'
        });
        
        const events = response.items || [];
        
        events.forEach(ev => {
          let startTime, endTime, isAllDay;
          
          if (ev.start.date) {
            startTime = new Date(ev.start.date.replace(/-/g, '/'));
            endTime = new Date(ev.end.date.replace(/-/g, '/'));
            isAllDay = true;
          } else {
            startTime = new Date(ev.start.dateTime);
            endTime = new Date(ev.end.dateTime);
            isAllDay = false;
          }
          
          if (startTime < targetDate) startTime = targetDate;
          if (endTime > nextDate) endTime = nextDate;

          userEvents.push({
            title: ev.summary || "予定あり", 
            startTime: startTime.getTime(),
            endTime: endTime.getTime(),
            isAllDay: isAllDay
          });
        });
      }
    } catch (e) {
      console.warn(`${email} のカレンダー取得エラー:`, e.message);
      hasAccess = false; 
    }

    const colorIndex = index % colors.length;

    results.push({
      email: email,
      name: displayName,
      color: colors[colorIndex],
      events: userEvents,
      hasAccess: hasAccess,
      isResource: false
    });
  });

  // 追加された会議室（リソース）の予定取得
  addedResources.forEach(res => {
    const resEvents = [];
    let hasAccess = true;
    try {
      if (typeof Calendar !== 'undefined') {
        const response = Calendar.Events.list(res.email, { 
          timeMin: targetDate.toISOString(), 
          timeMax: nextDate.toISOString(), 
          singleEvents: true, 
          orderBy: 'startTime' 
        });
        (response.items || []).forEach(ev => {
          let startTime, endTime, isAllDay;
          if (ev.start.date) {
            startTime = new Date(ev.start.date.replace(/-/g, '/')); 
            endTime = new Date(ev.end.date.replace(/-/g, '/')); 
            isAllDay = true;
          } else {
            startTime = new Date(ev.start.dateTime); 
            endTime = new Date(ev.end.dateTime); 
            isAllDay = false;
          }
          if (startTime < targetDate) startTime = targetDate;
          if (endTime > nextDate) endTime = nextDate;
          resEvents.push({ 
            title: ev.summary || "予定あり", 
            startTime: startTime.getTime(), 
            endTime: endTime.getTime(), 
            isAllDay: isAllDay 
          });
        });
      }
    } catch (e) { 
      console.warn(`会議室 ${res.email} 取得エラー:`, e.message);
      hasAccess = false; 
    }

    results.push({
      email: res.email, 
      name: res.name, 
      color: "#607D8B", // 会議室はブルーグレー系で固定して人と差別化
      events: resEvents, 
      hasAccess: hasAccess, 
      isResource: true
    });
  });

  return results;
}

/**
 * FreeBusy API を使用して、指定日時に「空いている」リソースのメールアドレス一覧を返す
 * @param {string[]} emails - チェック対象のリソースメールアドレス配列
 * @param {string} startIso - チェック開始日時の ISOString
 * @param {string} endIso - チェック終了日時の ISOString
 * @returns {Array} 空いているリソースのメールアドレス配列
 */
function getFreeResources(emails, startIso, endIso) {
  try {
    if (typeof Calendar === 'undefined') throw new Error("Calendar APIが有効になっていません。");
    
    let freeEmails = [];
    
    // FreeBusy API は一度に最大50件までのため、50件ずつチャンク処理（分割処理）する
    for (let i = 0; i < emails.length; i += 50) {
      const chunk = emails.slice(i, i + 50);
      const items = chunk.map(email => ({ id: email }));
      
      const request = {
        timeMin: startIso,
        timeMax: endIso,
        items: items
      };
      
      // 50件分の空き状況を一括取得
      const response = Calendar.Freebusy.query(request);
      const calendars = response.calendars || {};
      
      // busy（予定が入っている）配列が空（0件）のものを「空室」と判定
      chunk.forEach(email => {
         const calData = calendars[email];
         if (calData && calData.busy && calData.busy.length === 0) {
           freeEmails.push(email);
         }
      });
    }
    
    return freeEmails;
  } catch(e) {
    console.error("FreeBusy API エラー:", e.message);
    throw new Error(e.message);
  }
}
