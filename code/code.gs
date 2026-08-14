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
// データベース初期化処理
// ==========================================

/**
 * 本アプリで使用する全シートを一括で初期化（作成およびヘッダー設定）する
 * スプレッドシートに該当のシートがない場合のみ新規作成します。
 * AdminListはヘッダーなし。それ以外はヘッダーを設定し、固定やスタイルを適用します。
 */
function initializeDatabaseSheets() {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    
    const sheetsConfig = [
      { name: 'AdminList', headers: null },
      { name: 'StaticLinks', headers: ['Title', 'URL', 'Icon'] },
      { name: 'WF_RouteMaster', headers: ['RouteID', 'RouteName', 'Scope', 'Owner', 'IsActive', 'RetentionType', 'RetentionValue', 'CreatedAt'] },
      { name: 'WF_RouteStepMaster', headers: ['RouteID', 'StepOrder', 'ActionType', 'Assignee'] },
      { name: 'WF_Tickets', headers: ['TicketID', 'RouteID', 'Subject', 'Applicant', 'AppliedAt', 'Body', 'Attachments', 'FolderId', 'ExpireAt', 'CurrentStep', 'Status'] },
      { name: 'WF_TicketRoutes', headers: ['TicketID', 'StepOrder', 'ActionType', 'Assignee'] },
      { name: 'WF_ApprovalLogs', headers: ['TicketID', 'StepOrder', 'Actor', 'ActionAt', 'Action', 'Comment'] },
      { name: 'WF_Tickets_Archive', headers: ['TicketID', 'RouteID', 'Subject', 'Applicant', 'AppliedAt', 'Body', 'Attachments', 'FolderId', 'ExpireAt', 'CurrentStep', 'Status'] },
      { name: 'WF_TicketRoutes_Archive', headers: ['TicketID', 'StepOrder', 'ActionType', 'Assignee'] },
      { name: 'WF_ApprovalLogs_Archive', headers: ['TicketID', 'StepOrder', 'Actor', 'ActionAt', 'Action', 'Comment'] },
      { name: 'TasksDB', headers: ['Task ID', 'Title', 'Assignee', 'Status', 'Google Tasks ID', 'Last Updated', 'Scope', 'Priority', 'DueDate'] },
      { name: 'TasksDB_Archive', headers: ['Task ID', 'Title', 'Assignee', 'Status', 'Google Tasks ID', 'Last Updated', 'Scope', 'Priority', 'DueDate'] }
    ];

    // 不要シート判定のため、定義されたシート名のリストを抽出
    const requiredSheetNames = sheetsConfig.map(config => config.name);

    sheetsConfig.forEach(config => {
      let sheet = ss.getSheetByName(config.name);
      if (!sheet) {
        sheet = ss.insertSheet(config.name);
        if (config.headers) {
          sheet.appendRow(config.headers);
          sheet.getRange(1, 1, 1, config.headers.length).setFontWeight('bold').setBackground('#f3f4f6');
          sheet.setFrozenRows(1);
        }
      } else {
        // 既存シートのマイグレーション補完（TasksDB用）
        if (config.name === 'TasksDB') {
          if (sheet.getRange(1, 7).getValue() === '') sheet.getRange(1, 7).setValue('Scope');
          if (sheet.getRange(1, 8).getValue() === '') {
            sheet.getRange(1, 8, 1, 2).setValues([['Priority', 'DueDate']]);
            sheet.getRange(1, 7, 1, 3).setBackground('#f3f4f6').setFontWeight('bold');
          }
        }
      }
    });

    // 必須シートの作成後、スプレッドシート内の全シートを走査し不要なものを削除
    const allSheets = ss.getSheets();
    allSheets.forEach(sheet => {
      if (!requiredSheetNames.includes(sheet.getName())) {
        ss.deleteSheet(sheet);
      }
    });

    return true;
  } catch (e) {
    throw new Error(getMsg('errDbInit') + e.message);
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
  if (getMaintenanceMode()) {
    return { isMaintenanceMode: true };
  }

  return {
    isMaintenanceMode: false,
    userLocale: getUserLocale(),
    masterLinks: getMasterLinks(),
    userGroups: getUserGroups(),
    cloudSearchEnabled: getCloudSearchSetting(),
    userConfig: getUserConfig(),
    archiveSchedule: getArchiveSchedule()
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
 * 環境変数として設定されたスプレッドシートIDを取得する
 * プロパティが未設定の場合、コンテナバインドされているスプレッドシートのIDをフォールバックとして試行します。
 * @returns {string} スプレッドシートID
 */
function getSpreadsheetId() {
  let id = PropertiesService.getScriptProperties().getProperty('LINK_SPREADSHEET_ID');
  if (!id) {
    try {
      id = SpreadsheetApp.getActiveSpreadsheet().getId();
    } catch (e) {
      throw new Error(getMsg('errLinkSpreadsheetIdMissing'));
    }
  }
  return id;
}

/**
 * 実行ユーザーがシステム管理者かどうかを判定する
 * @returns {boolean} 管理者の場合はtrue
 */
function isAdminUser() {
  try {
    const email = Session.getActiveUser().getEmail();
    const sheet = SpreadsheetApp.openById(getSpreadsheetId()).getSheetByName('AdminList');
    if (!sheet) return false;
    const data = sheet.getDataRange().getValues();
    return data.some(row => row[0] === email); 
  } catch (e) {
    return false;
  }
}

/**
 * Cloud Search連携の有効/無効状態を取得する
 * @returns {boolean} 有効な場合はtrue
 */
function getCloudSearchSetting() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('CLOUD_SEARCH_SETTING_CACHE');
  if (cached) return cached === 'true';

  const props = PropertiesService.getScriptProperties();
  const val = props.getProperty('ENABLE_CLOUD_SEARCH') === 'true';
  try { cache.put('CLOUD_SEARCH_SETTING_CACHE', val ? 'true' : 'false', 21600); } catch(e){}
  return val;
}

/**
 * Cloud Search連携の有効/無効状態を更新する
 * @param {boolean} isEnable - 設定値
 * @returns {boolean} 成功時
 */
function updateCloudSearchSetting(isEnable) {
  if (!isAdminUser()) throw new Error("権限がありません。");
  const props = PropertiesService.getScriptProperties();
  props.setProperty('ENABLE_CLOUD_SEARCH', isEnable ? 'true' : 'false');
  
  try { CacheService.getScriptCache().put('CLOUD_SEARCH_SETTING_CACHE', isEnable ? 'true' : 'false', 21600); } catch(e){}
  return true;
}

/**
 * スプレッドシートから全社共通のリンクマスタを取得する
 * @returns {Array} リンク情報の配列
 */
function getMasterLinks() {
  try {
    const cache = CacheService.getScriptCache();
    const cachedLinks = cache.get('MASTER_LINKS_CACHE');
    if (cachedLinks) return JSON.parse(cachedLinks);

    const sheet = SpreadsheetApp.openById(getSpreadsheetId()).getSheetByName('StaticLinks');
    if (!sheet) throw new Error("シートが見つかりません");
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const links = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || !row[1]) continue; 
      
      const urlStr = row[1].toString();
      const urlHash = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, urlStr)).replace(/=/g, '').substring(0, 15);

      links.push({
        id: 'link_' + urlHash, 
        title: row[0].toString(),
        url: urlStr,
        icon: row[2] ? row[2].toString() : 'link'
      });
    }
    
    try { cache.put('MASTER_LINKS_CACHE', JSON.stringify(links), 21600); } catch(e){}
    return links;
  } catch (e) {
    console.error("リンクマスタ取得エラー:", e);
    return [];
  }
}

/**
 * 全社共通のリンクマスタをスプレッドシートに保存する
 * @param {Array} newLinks - 更新後のリンク配列
 * @returns {Array} 更新完了後のリンク配列
 */
function updateMasterLinks(newLinks) {
  if (!isAdminUser()) throw new Error("権限がありません。");
  try {
    const sheet = SpreadsheetApp.openById(getSpreadsheetId()).getSheetByName('StaticLinks');
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
    if (newLinks && newLinks.length > 0) {
      const outputData = newLinks.map(link => [link.title, link.url, link.icon || 'link']);
      sheet.getRange(2, 1, outputData.length, 3).setValues(outputData);
    }
    
    CacheService.getScriptCache().remove('MASTER_LINKS_CACHE');
    return getMasterLinks(); 
  } catch (e) {
    throw new Error("保存に失敗しました。");
  }
}

// ==========================================
// ワークフロー管理API
// ==========================================

/**
 * スプレッドシート上にワークフロー管理用の各シート群を初期生成する
 * @returns {boolean} 成功時
 */
function initWorkflowSheets() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sheets = [
    { name: 'WF_RouteMaster', headers: ['RouteID', 'RouteName', 'Scope', 'Owner', 'IsActive', 'RetentionType', 'RetentionValue', 'CreatedAt'] },
    { name: 'WF_RouteStepMaster', headers: ['RouteID', 'StepOrder', 'ActionType', 'Assignee'] },
    { name: 'WF_Tickets', headers: ['TicketID', 'RouteID', 'Subject', 'Applicant', 'AppliedAt', 'Body', 'Attachments', 'FolderId', 'ExpireAt', 'CurrentStep', 'Status'] },
    { name: 'WF_TicketRoutes', headers: ['TicketID', 'StepOrder', 'ActionType', 'Assignee'] },
    { name: 'WF_ApprovalLogs', headers: ['TicketID', 'StepOrder', 'Actor', 'ActionAt', 'Action', 'Comment'] }
  ];
  
  sheets.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.headers);
      sheet.getRange(1, 1, 1, s.headers.length).setFontWeight('bold').setBackground('#f3f4f6');
      sheet.setFrozenRows(1);
    }
  });
  return true;
}

/**
 * 自分が閲覧または編集可能な承認経路（ルート）の一覧を取得する
 * @returns {Array} 利用可能なルート情報の配列
 */
function getWorkflowRoutes() {
  const email = Session.getActiveUser().getEmail();
  const groups = getUserGroups();
  
  let routesData = [];
  let stepsData = [];
  
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('WF_MASTER_DATA_CACHE');
  
  if (cachedData) {
    const parsed = JSON.parse(cachedData);
    routesData = parsed.routes;
    stepsData = parsed.steps;
  } else {
    initWorkflowSheets();
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const routeSheet = ss.getSheetByName('WF_RouteMaster');
    const stepSheet = ss.getSheetByName('WF_RouteStepMaster');
    if (!routeSheet || !stepSheet) return []; 
    
    routesData = routeSheet.getDataRange().getValues();
    stepsData = stepSheet.getDataRange().getValues();
    
    try { cache.put('WF_MASTER_DATA_CACHE', JSON.stringify({ routes: routesData, steps: stepsData }), 21600); } catch(e){}
  }

  if (routesData.length <= 1) return [];
  
  const routes = [];
  for (let i = 1; i < routesData.length; i++) {
    const r = routesData[i];
    if (!r[0]) continue;
    const routeId = r[0], scope = r[2], owner = r[3];
    
    if (owner === email || (groups && groups.includes(scope))) {
      const steps = stepsData.filter(s => s[0] === routeId).map(s => ({
        stepOrder: Number(s[1]), actionType: s[2], assignee: s[3]
      })).sort((a, b) => a.stepOrder - b.stepOrder);
      
      routes.push({
        id: routeId, name: r[1], scope: scope, owner: owner,
        isActive: r[4] === true || r[4] === 'TRUE' || r[4] === 'true',
        retentionType: r[5] || 'PERIOD', retentionValue: r[6] || 5, steps: steps
      });
    }
  }
  return routes;
}

/**
 * 承認経路（マスタ）を新規作成または更新する
 * データ競合を防ぐため排他ロックを使用します。
 * @param {Array} routes - 保存対象のルート配列
 * @param {Array} deletedIds - 削除対象のルートID配列
 * @returns {boolean} 成功時
 */
function saveWorkflowRoutes(routes, deletedIds = []) {
  // バックエンドでの入力値検証
  if (!routes || !Array.isArray(routes)) {
    throw new Error(getMsg('errInvalidInput'));
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error(getMsg('errUpdateConflict'));
  try {
    const email = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const routeSheet = ss.getSheetByName('WF_RouteMaster');
    const stepSheet = ss.getSheetByName('WF_RouteStepMaster');
    
    const rData = routeSheet.getDataRange().getValues();
    const sData = stepSheet.getDataRange().getValues();
    const rHeaders = rData[0], sHeaders = sData[0];
    let rRows = rData.slice(1), sRows = sData.slice(1);
    
    const updateIds = routes.map(r => r.id);
    const removeIds = [...deletedIds, ...updateIds];
    const isAdmin = isAdminUser();

    rRows = rRows.filter(row => {
      if (removeIds.includes(row[0])) {
        // 既存データのオーナーが自分自身でない、かつ管理者でもない場合は操作をブロック（IDOR対策）
        if (row[3] !== email && !isAdmin) {
          throw new Error(getMsg('errUnauthorized'));
        }
        return false;
      }
      return true;
    });
    sRows = sRows.filter(row => !removeIds.includes(row[0]));
    
    const now = new Date();
    routes.forEach(route => {
      rRows.push([route.id, route.name, route.scope || 'PERSONAL', route.owner || email, route.isActive, route.retentionType, route.retentionValue, now]);
      if (route.steps) {
        route.steps.forEach((step, index) => {
          sRows.push([route.id, index + 1, step.actionType, step.assignee]);
        });
      }
    });
    
    routeSheet.clearContents();
    stepSheet.clearContents();
    if (rRows.length > 0) routeSheet.getRange(1, 1, rRows.length + 1, rHeaders.length).setValues([rHeaders, ...rRows]);
    else routeSheet.appendRow(rHeaders);
    if (sRows.length > 0) stepSheet.getRange(1, 1, sRows.length + 1, sHeaders.length).setValues([sHeaders, ...sRows]);
    else stepSheet.appendRow(sHeaders);
    
    CacheService.getScriptCache().remove('WF_MASTER_DATA_CACHE');
    
    return true;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 申請画面用に利用可能な（アクティブな）ルートのみを取得する
 * @returns {Array} アクティブなルートの配列
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
 * 添付ファイルのドライブ格納およびDBへの書き込みをトランザクション的に行います。
 * @param {Object} payload - 申請内容や経路情報を含むデータオブジェクト
 * @returns {boolean} 成功時
 */
function submitWorkflowApplication(payload) {
  // バックエンドでの厳密な入力値検証（コンソールからの直接実行等による不正アクセスを防止）
  if (!payload || !payload.subject || !payload.body || !payload.steps || payload.steps.length === 0) {
    throw new Error(getMsg('errInvalidInput'));
  }
  for (let s of payload.steps) {
    if (!s.assignee) throw new Error(getMsg('errInvalidInputAssignee'));
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(45000)) throw new Error(getMsg('errServerBusy'));
  
  try {
    const email = Session.getActiveUser().getEmail();
    const now = new Date();
    const yyyymmdd = Utilities.formatDate(now, "JST", "yyyyMMdd");
    const ticketId = 'ticket_' + Utilities.getUuid().replace(/-/g, '').substring(0, 15);
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    
    let finalRouteId = payload.routeId;
    
    // カスタム経路を新しいテンプレートとして保存する場合
    if (payload.saveTemplate && payload.templateParams) {
      finalRouteId = 'route_' + Date.now();
      const routeSheet = ss.getSheetByName('WF_RouteMaster');
      const stepSheet = ss.getSheetByName('WF_RouteStepMaster');
      
      routeSheet.appendRow([finalRouteId, payload.templateParams.name, payload.templateParams.scope, email, true, payload.retentionType, payload.retentionValue, now]);
      const sData = [];
      payload.steps.forEach((step, index) => {
        sData.push([finalRouteId, index + 1, step.actionType, step.assignee]);
      });
      if (sData.length > 0) {
        stepSheet.getRange(stepSheet.getLastRow() + 1, 1, sData.length, 4).setValues(sData);
      }
    }
    
    // アーカイブ/削除基準日の計算
    let expireAt = "";
    if (payload.retentionType === 'PERIOD') {
      const expDate = new Date(now);
      expDate.setFullYear(expDate.getFullYear() + Number(payload.retentionValue));
      expireAt = Utilities.formatDate(expDate, "JST", "yyyy/MM/dd");
    } else if (payload.retentionType === 'DATE') {
      expireAt = payload.retentionValue; 
    } else {
      expireAt = "PERMANENT"; 
    }

    // ドライブへのファイル実体保存（フォルダ生成とファイル複製）
    let folderId = "";
    let savedAttachments = [];
    if ((payload.attachments && payload.attachments.length > 0) || (payload.driveFileIds && payload.driveFileIds.length > 0)) {
      const parentFolderId = PropertiesService.getScriptProperties().getProperty('WORKFLOW_FOLDER_ID');
      if (!parentFolderId) throw new Error(getMsg('errWorkflowFolderIdMissing'));
      
      const rootFolder = DriveApp.getFolderById(parentFolderId);
      const safeSubject = payload.subject.replace(/[\\/:*?"<>|]/g, "_").substring(0, 20);
      const targetFolder = rootFolder.createFolder(`${yyyymmdd}_${safeSubject}`);
      folderId = targetFolder.getId();
      
      if (payload.attachments) {
        payload.attachments.forEach(att => {
          const blob = Utilities.newBlob(Utilities.base64Decode(att.data), att.mimeType, att.name);
          const file = targetFolder.createFile(blob);
          savedAttachments.push({ id: file.getId(), name: file.getName() });
        });
      }
      if (payload.driveFileIds) {
        payload.driveFileIds.forEach(id => {
          const originalFile = DriveApp.getFileById(id);
          const file = originalFile.makeCopy(originalFile.getName(), targetFolder);
          savedAttachments.push({ id: file.getId(), name: file.getName() });
        });
      }
    }
    
    // DB(スプレッドシート)への書き込み
    const ticketSheet = ss.getSheetByName('WF_Tickets');
    const ticketRouteSheet = ss.getSheetByName('WF_TicketRoutes');
    const logSheet = ss.getSheetByName('WF_ApprovalLogs');

    // チケットに紐づく一意の経路情報を保存
    const trData = [];
    payload.steps.forEach((step, index) => {
      trData.push([ticketId, index + 1, step.actionType, step.assignee]);
    });
    if (trData.length > 0) {
      ticketRouteSheet.getRange(ticketRouteSheet.getLastRow() + 1, 1, trData.length, 4).setValues(trData);
    }

    // 初回に誰へボールを渡すか（最初のAPPROVEステップ）を特定
    let firstApproveStep = null;
    for (let i = 0; i < payload.steps.length; i++) {
      if (payload.steps[i].actionType === 'APPROVE') {
        firstApproveStep = i + 1;
        break;
      }
    }
    const initialStep = firstApproveStep || (payload.steps.length > 0 ? payload.steps.length + 1 : 1);
    const initialStatus = firstApproveStep ? 'PENDING' : 'APPROVED';
    
    ticketSheet.appendRow([
      ticketId, finalRouteId || '', payload.subject, email, now, payload.body, 
      JSON.stringify(savedAttachments), folderId, expireAt, initialStep, initialStatus
    ]);
    logSheet.appendRow([ ticketId, 0, email, now, 'APPLIED', '新規申請' ]);
    
    return true;
  } catch(e) {
    throw new Error(e.message || String(e));
  } finally {
    lock.releaseLock();
  }
}

/**
 * 実行ユーザーが過去に起票した申請履歴（マイチケット）を取得する
 * @returns {Array} チケット情報の配列
 */
function getMyTickets() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  
  const ticketSheet = ss.getSheetByName('WF_Tickets');
  const stepSheet = ss.getSheetByName('WF_RouteStepMaster');
  if (!ticketSheet) return [];

  const tickets = ticketSheet.getDataRange().getValues().slice(1);
  const steps = stepSheet.getDataRange().getValues().slice(1);
  const myTickets = [];

  tickets.forEach(t => {
    const ticketId = t[0], routeId = t[1], subject = t[2], applicant = t[3];
    const appliedAt = t[4], body = t[5], attachments = t[6], folderId = t[7];
    const currentStep = Number(t[9]), status = t[10];

    if (applicant === email) {
      // 全体進捗バー用のため全ステップ数を算出
      const routeSteps = steps.filter(s => s[0] === routeId);
      const totalSteps = routeSteps.length > 0 ? Math.max(...routeSteps.map(s => Number(s[1]))) : 1;
      
      let files = [];
      try { 
        const parsed = JSON.parse(attachments || '[]'); 
        files = parsed.map(f => {
          if (typeof f === 'string') {
            try { return { id: f, name: DriveApp.getFileById(f).getName() }; } catch(e) { return { id: f, name: 'ファイルを開く' }; }
          }
          return f;
        });
      } catch(e){}
      
      myTickets.push({
        ticketId, routeId, subject, applicant, 
        appliedAt: Utilities.formatDate(new Date(appliedAt), "JST", "yyyy/MM/dd HH:mm"), 
        body, files, folderId, currentStep, totalSteps, status,
        actionType: 'READ_ONLY'
      });
    }
  });
  
  return myTickets.sort((a, b) => new Date(b.appliedAt) - new Date(a.appliedAt));
}

/**
 * 実行ユーザーが担当者となっている承認・供覧待ちタスクを取得する
 * @returns {Array} タスク情報の配列
 */
function getPendingTasks() {
  const email = Session.getActiveUser().getEmail();
  const groups = getUserGroups();
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  
  const ticketSheet = ss.getSheetByName('WF_Tickets');
  const masterStepSheet = ss.getSheetByName('WF_RouteStepMaster');
  const ticketRouteSheet = ss.getSheetByName('WF_TicketRoutes');
  const logSheet = ss.getSheetByName('WF_ApprovalLogs');
  if (!ticketSheet) return [];

  const tickets = ticketSheet.getDataRange().getValues().slice(1);
  const logs = logSheet.getDataRange().getValues().slice(1);
  const ticketRoutesData = ticketRouteSheet ? ticketRouteSheet.getDataRange().getValues().slice(1) : [];
  const masterStepsData = masterStepSheet.getDataRange().getValues().slice(1);
  const pendingTasks = [];

  tickets.forEach(t => {
    const ticketId = t[0], routeId = t[1], subject = t[2], applicant = t[3];
    const currentStep = Number(t[9]), status = t[10];

    if (status !== 'PENDING' && status !== 'APPROVED') return;

    // チケットに固有経路があればそちらを、なければマスタ経路を使用する
    let routeSteps = ticketRoutesData.filter(s => s[0] === ticketId);
    if (routeSteps.length === 0) routeSteps = masterStepsData.filter(s => s[0] === routeId);

    const myLogs = logs.filter(l => l[0] === ticketId && l[2] === email);
    const hasApprovedEver = myLogs.some(l => l[4] === 'APPROVE');
    let myTicketTasks = [];

    routeSteps.forEach(s => {
      const stepOrder = Number(s[1]), actionType = s[2], assignee = s[3];
      if (assignee === email || groups.includes(assignee)) {
        const hasActedThisStep = myLogs.some(l => Number(l[1]) === stepOrder);
        if (!hasActedThisStep) {
          if (actionType === 'APPROVE' && currentStep === stepOrder && status === 'PENDING') {
            myTicketTasks.push({ stepOrder, actionType, assignee });
          } else if (actionType === 'VIEW' && currentStep >= stepOrder) {
            myTicketTasks.push({ stepOrder, actionType, assignee });
          }
        }
      }
    });

    const isApproverCurrently = myTicketTasks.some(mt => mt.actionType === 'APPROVE');
    if (hasApprovedEver || isApproverCurrently) {
      myTicketTasks = myTicketTasks.filter(mt => mt.actionType === 'APPROVE');
    }

    myTicketTasks.forEach(mt => {
      let files = [];
      try { 
        const parsed = JSON.parse(t[6] || '[]'); 
        files = parsed.map(f => {
          if (typeof f === 'string') {
            try { return { id: f, name: DriveApp.getFileById(f).getName() }; } catch(e) { return { id: f, name: 'ファイルを開く' }; }
          }
          return f;
        });
      } catch(e){}
      pendingTasks.push({
        ticketId, routeId, subject, applicant, 
        appliedAt: Utilities.formatDate(new Date(t[4]), "JST", "yyyy/MM/dd HH:mm"), 
        body: t[5], files, folderId: t[7], stepOrder: mt.stepOrder, actionType: mt.actionType
      });
    });
  });
  
  return pendingTasks.sort((a, b) => new Date(b.appliedAt) - new Date(a.appliedAt));
}

/**
 * 対象のチケットに対するアクション（承認、差戻、却下、確認）を実行する
 * 指数的バックオフを用いたリトライ処理を組み込んでいます。
 * @param {string} ticketId - チケットID
 * @param {number} stepOrder - アクションを実行するステップの順番
 * @param {string} actionType - 'APPROVE', 'RETURN', 'REJECT', 'CONFIRM' 等
 * @param {string} comment - コメント内容
 * @returns {boolean} 成功時
 */
function executeWorkflowAction(ticketId, stepOrder, actionType, comment) {
  // バックエンドでの入力値検証（不正アクセス防止）
  if ((actionType === 'REJECT' || actionType === 'RETURN') && (!comment || comment.trim() === '')) {
    throw new Error(getMsg('errCommentRequired'));
  }

  const email = Session.getActiveUser().getEmail();
  const maxRetries = 3;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    const lock = LockService.getScriptLock();
    if (lock.tryLock(15000)) {
      try {
        const now = new Date();
        const ss = SpreadsheetApp.openById(getSpreadsheetId());
        const ticketSheet = ss.getSheetByName('WF_Tickets');
        const logSheet = ss.getSheetByName('WF_ApprovalLogs');
        const ticketRouteSheet = ss.getSheetByName('WF_TicketRoutes');
        const masterStepSheet = ss.getSheetByName('WF_RouteStepMaster');
        
        const tData = ticketSheet.getDataRange().getValues();
        const tIndex = tData.findIndex(t => t[0] === ticketId);
        if (tIndex === -1) throw new Error(getMsg('errTicketNotFound'));
        
        const ticket = tData[tIndex];
        let currentStep = Number(ticket[9]);
        let newStatus = ticket[10];
        
        // バックエンドでのアクセス権限（認可）検証
        let allSteps = ticketRouteSheet.getDataRange().getValues().slice(1).filter(s => s[0] === ticketId);
        if (allSteps.length === 0) allSteps = masterStepSheet.getDataRange().getValues().slice(1).filter(s => s[0] === ticket[1]);
        
        if (actionType === 'APPROVE' || actionType === 'REJECT' || actionType === 'RETURN') {
          const currentStepInfo = allSteps.find(s => Number(s[1]) === currentStep);
          if (!currentStepInfo) throw new Error(getMsg('errRouteInfoNotFound'));
          
          const assignee = currentStepInfo[3];
          const userGroups = getUserGroups();
          if (assignee !== email && !userGroups.includes(assignee)) {
            throw new Error(getMsg('errTaskNoPermission'));
          }
        }

        if (actionType === 'REJECT') newStatus = 'REJECTED';
        else if (actionType === 'RETURN') newStatus = 'RETURNED';
        else if (actionType === 'CONFIRM') {
          // 供覧確認はステータスやステップを進行させない
        } 
        else if (actionType === 'APPROVE') {
          if (newStatus !== 'PENDING') throw new Error(getMsg('errAppAlreadyCompleted'));
          
          allSteps.sort((a, b) => Number(a[1]) - Number(b[1]));
          
          let nextApproveStepOrder = null;
          for (let i = 0; i < allSteps.length; i++) {
            if (Number(allSteps[i][1]) > currentStep && allSteps[i][2] === 'APPROVE') {
              nextApproveStepOrder = Number(allSteps[i][1]);
              break;
            }
          }
          if (nextApproveStepOrder) currentStep = nextApproveStepOrder;
          else {
            currentStep = currentStep + 1;
            newStatus = 'APPROVED';
          }
        }
        
        ticketSheet.getRange(tIndex + 1, 10).setValue(currentStep);
        ticketSheet.getRange(tIndex + 1, 11).setValue(newStatus);
        logSheet.appendRow([ticketId, stepOrder, email, now, actionType, comment]);
        SpreadsheetApp.flush();
        return true;
      } catch (e) {
        if (attempt === maxRetries) throw new Error(getMsg('errProcessFailed') + e.message);
        Utilities.sleep(Math.pow(2, attempt) * 1000 + Math.random() * 1000);
      } finally {
        lock.releaseLock();
      }
    } else {
      if (attempt === maxRetries) throw new Error(getMsg('errServerBusy'));
      Utilities.sleep(Math.pow(2, attempt) * 1000 + Math.random() * 1000);
    }
  }
}

// ==========================================
// BCP（防災・緊急連絡）API
// ==========================================

/**
 * 気象庁API等から指定したエリアの防災情報を取得する
 * @param {string} area1 - エリアコード1
 * @param {string} area2 - エリアコード2
 * @returns {Object} 気象情報および管理者が設定した緊急メッセージ
 */
function getBcpData(area1, area2) {
  const result = {
    weather1: { text: "取得できませんでした", reportDatetime: "", name: "" },
    weather2: { text: "取得できませんでした", reportDatetime: "", name: "" },
    emergencyMessage: PropertiesService.getScriptProperties().getProperty('EMERGENCY_MESSAGE') || ""
  };

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
  PropertiesService.getScriptProperties().setProperty('EMERGENCY_MESSAGE', message.trim());
  return true;
}

// =========================================================================
// フルスクリーンカンバン ＆ Google Tasks ハイブリッド同期 バックエンド
// =========================================================================

/**
 * カンバンDB用のスプレッドシートオブジェクトを取得する
 * @returns {Spreadsheet} スプレッドシートのオブジェクト
 */
function getKanbanSpreadsheet() {
  try {
    const id = getSpreadsheetId();
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error("データベース用のスプレッドシートが見つかりません: " + e.message);
  }
}

/**
 * カンバンデータ格納用のシートを初期化（または構成更新）する
 * @returns {Sheet} 初期化または取得済みのシート
 */
function initKanbanSheet() {
  const ss = getKanbanSpreadsheet();
  let sheet = ss.getSheetByName('TasksDB');
  if (!sheet) {
    sheet = ss.insertSheet('TasksDB');
    sheet.getRange(1, 1, 1, 9).setValues([['Task ID', 'Title', 'Assignee', 'Status', 'Google Tasks ID', 'Last Updated', 'Scope', 'Priority', 'DueDate']]);
    sheet.getRange(1, 1, 1, 9).setBackground('#f3f3f3').setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    // 既存環境向けの列追加マイグレーション
    if (sheet.getRange(1, 7).getValue() === '') sheet.getRange(1, 7).setValue('Scope');
    if (sheet.getRange(1, 8).getValue() === '') {
      sheet.getRange(1, 8, 1, 2).setValues([['Priority', 'DueDate']]);
      sheet.getRange(1, 7, 1, 3).setBackground('#f3f3f3').setFontWeight('bold');
    }
  }
  return sheet;
}

/**
 * すべてのカンバンタスクを取得し、Google Tasksの状態と自動同期を行う
 * 外部（Google Tasks側）で作成された未登録タスクは、ポータル側へ逆同期して取り込みます。
 * @returns {Array<Object>} タスクデータの配列
 */
function getKanbanTasks() {
  const sheet = initKanbanSheet();
  const data = sheet.getDataRange().getValues();
  const currentUser = Session.getActiveUser().getEmail();

  const tasks = [];
  const dbGoogleTaskIds = new Set();
  
  if (data.length > 1) {
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let dueDateStr = '';
      if (row[8]) dueDateStr = Utilities.formatDate(new Date(row[8]), "JST", "yyyy-MM-dd");

      let task = {
        id: row[0],
        title: row[1],
        assignee: row[2],
        status: row[3],
        googleTaskId: row[4],
        scope: row[6] || 'PERSONAL',
        priority: row[7] || 'Medium',
        dueDate: dueDateStr,          
        isMyTask: (row[2] === currentUser),
        rowNum: i + 1
      };
      tasks.push(task);
      if (task.googleTaskId) dbGoogleTaskIds.add(task.googleTaskId);
    }
  }

  // Google Tasksのステータスを一括取得し、未登録のタスクを抽出する
  let googleTasksMap = {};
  let importedTasks = [];
  try {
    let pageToken;
    do {
      const response = Tasks.Tasks.list('@default', {
        maxResults: 100,
        showCompleted: true,
        showHidden: true,
        showDeleted: true, // 削除または不可視になったタスクも強制取得
        pageToken: pageToken
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
    console.error("Google Tasks一括取得エラー:", e);
  }

  let isUpdated = false;

  // DBタスクとGoogle Tasksの相互状態同期処理
  tasks.forEach(task => {
    if (task.assignee === currentUser && task.status !== 'done') {
      if (!task.googleTaskId) {
        try {
          let gTaskOpts = { title: task.title, notes: "ポータルのカンバンと同期されています。" };
          if (task.dueDate) gTaskOpts.due = task.dueDate + 'T00:00:00.000Z';
          const gTask = Tasks.Tasks.insert(gTaskOpts, '@default');
          task.googleTaskId = gTask.id;
          sheet.getRange(task.rowNum, 5).setValue(gTask.id);
          isUpdated = true;
        } catch (e) {}
      } else {
        try {
          const gTaskStatus = googleTasksMap[task.googleTaskId];
          // APIから返ってこない(undefined)場合も、Google Tasks側で完全に消去されたとみなし「完了」として同期する
          if (gTaskStatus === 'completed' || gTaskStatus === undefined) {
            task.status = 'done';
            sheet.getRange(task.rowNum, 4).setValue('done');
            sheet.getRange(task.rowNum, 6).setValue(new Date());
            isUpdated = true;
          }
        } catch (e) {}
      }
    }
  });

  // 外部で作成されたGoogle TasksをDBへ新規追加（逆同期）
  if (importedTasks.length > 0) {
    const newRows = [];
    const now = new Date();
    let nextRowNum = data.length > 0 ? data.length + 1 : 2;
    
    importedTasks.forEach(gTask => {
      const taskId = Utilities.getUuid();
      let dueDateStr = '';
      if (gTask.due) dueDateStr = Utilities.formatDate(new Date(gTask.due), "JST", "yyyy-MM-dd");
      
      const newTask = {
        id: taskId,
        title: gTask.title || '(無題)',
        assignee: currentUser,
        status: 'todo',
        googleTaskId: gTask.id,
        scope: 'PERSONAL',
        priority: 'Medium',
        dueDate: dueDateStr,
        isMyTask: true,
        rowNum: nextRowNum
      };
      tasks.push(newTask);
      newRows.push([newTask.id, newTask.title, newTask.assignee, newTask.status, newTask.googleTaskId, now, newTask.scope, newTask.priority, newTask.dueDate]);
      nextRowNum++;
    });
    
    // スプレッドシートへ一括書き込みを行い、通信回数を最適化
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    isUpdated = true;
  }

  if (isUpdated) SpreadsheetApp.flush();
  return tasks;
}

/**
 * カンバンに新たなタスクを追加する
 * 自分が担当者の場合はGoogle Tasksへの起票も行います。
 * @param {string} title - タスク名
 * @param {string} assignee - 担当者のメールアドレス
 * @param {string} status - 初期ステータス
 * @param {string} scope - タスクの紐づくスコープ（PERSONAL等）
 * @param {string} priority - 優先度（High/Medium/Low）
 * @param {string} dueDate - 期限（YYYY-MM-DD）
 * @returns {Array<Object>} 更新後のタスク一覧
 */
function addKanbanTask(title, assignee, status, scope, priority, dueDate) {
  const sheet = initKanbanSheet();
  const taskId = Utilities.getUuid();
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

  sheet.appendRow([taskId, title, assignee, status, googleTaskId, new Date(), scope || 'PERSONAL', priority || 'Medium', dueDate || '']);
  return getKanbanTasks();
}

/**
 * カンバンタスクのステータス（進行列）を移動させる
 * 衝突回避のため排他ロックと指数的バックオフによるリトライ処理を内包しています。
 * @param {string} taskId - 更新対象のタスクID
 * @param {string} newStatus - 変更後のステータス
 * @returns {Array} 更新後のタスク一覧
 */
function updateKanbanTaskStatus(taskId, newStatus) {
  const currentUser = Session.getActiveUser().getEmail();
  const maxRetries = 3;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    const lock = LockService.getScriptLock();
    if (lock.tryLock(15000)) {
      try {
        const sheet = initKanbanSheet();
        const data = sheet.getDataRange().getValues();
        let isFound = false;

        for (let i = 1; i < data.length; i++) {
          if (data[i][0] === taskId) {
            isFound = true;
            const rowNum = i + 1;
            const assignee = data[i][2];
            const googleTaskId = data[i][4];

            sheet.getRange(rowNum, 4).setValue(newStatus);
            sheet.getRange(rowNum, 6).setValue(new Date());

            if (assignee === currentUser && googleTaskId) {
              try {
                const gTask = Tasks.Tasks.get('@default', googleTaskId);
                if (newStatus === 'done') gTask.status = 'completed';
                else gTask.status = 'needsAction';
                Tasks.Tasks.patch(gTask, '@default', googleTaskId);
              } catch (e) {}
            }
            break;
          }
        }
        
        SpreadsheetApp.flush();
        return getKanbanTasks(); 

      } catch (e) {
        if (attempt === maxRetries) throw new Error("タスクの更新に失敗しました: " + e.message);
        Utilities.sleep(Math.pow(2, attempt) * 1000 + Math.random() * 1000);
      } finally {
        lock.releaseLock();
      }
    } else {
      if (attempt === maxRetries) throw new Error("サーバーが非常に混み合っています。画面を再読み込みしてください。");
      Utilities.sleep(Math.pow(2, attempt) * 1000 + Math.random() * 1000); 
    }
  }
}

/**
 * カンバンタスクの属性を編集する
 * 担当者が自分に変更された場合はGoogle Tasksと紐付け、外れた場合はGoogle Tasks側から破棄します。
 * @param {string} taskId - 編集対象のタスクID
 * @param {string} title - 新しいタスク名
 * @param {string} assignee - 新しい担当者
 * @param {string} priority - 新しい優先度
 * @param {string} dueDate - 新しい期限
 * @param {string} scope - 新しいスコープ
 * @returns {Array} 更新後のタスク一覧
 */
function editKanbanTask(taskId, title, assignee, priority, dueDate, scope) {
  const sheet = initKanbanSheet();
  const data = sheet.getDataRange().getValues();
  const currentUser = Session.getActiveUser().getEmail();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === taskId) {
      const rowNum = i + 1;
      const oldAssignee = data[i][2];
      let googleTaskId = data[i][4];

      sheet.getRange(rowNum, 2).setValue(title);
      sheet.getRange(rowNum, 3).setValue(assignee);
      sheet.getRange(rowNum, 6).setValue(new Date());
      sheet.getRange(rowNum, 7).setValue(scope || 'PERSONAL');
      sheet.getRange(rowNum, 8).setValue(priority);
      sheet.getRange(rowNum, 9).setValue(dueDate);

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
             sheet.getRange(rowNum, 5).setValue(gTask.id);
           } catch (e) {}
         }
      } else if (oldAssignee === currentUser && googleTaskId) {
         try { Tasks.Tasks.remove('@default', googleTaskId); } catch(e) {}
         sheet.getRange(rowNum, 5).setValue(''); 
      }
      break;
    }
  }
  return getKanbanTasks();
}

/**
 * カンバンタスクを完全に削除する
 * @param {string} taskId - 削除対象のタスクID
 * @returns {Array} 削除後のタスク一覧
 */
function deleteKanbanTask(taskId) {
  const currentUser = Session.getActiveUser().getEmail();
  const maxRetries = 3;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    const lock = LockService.getScriptLock();
    if (lock.tryLock(15000)) {
      try {
        const sheet = initKanbanSheet();
        const data = sheet.getDataRange().getValues();
        
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] === taskId) {
            const rowNum = i + 1;
            const assignee = data[i][2];
            const googleTaskId = data[i][4];

            if (assignee === currentUser && googleTaskId) {
              try {
                Tasks.Tasks.remove('@default', googleTaskId);
              } catch (e) {
                console.warn("Google Tasksの削除に失敗:", e);
              }
            }
            
            sheet.deleteRow(rowNum);
            break;
          }
        }
        SpreadsheetApp.flush();
        return getKanbanTasks();
      } catch (e) {
        if (attempt === maxRetries) throw new Error("削除に失敗しました: " + e.message);
        Utilities.sleep(Math.pow(2, attempt) * 1000 + Math.random() * 1000);
      } finally {
        lock.releaseLock();
      }
    } else {
      if (attempt === maxRetries) throw new Error("サーバーが混み合っています。時間をおいて再試行してください。");
      Utilities.sleep(Math.pow(2, attempt) * 1000 + Math.random() * 1000);
    }
  }
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
 * 完了したカンバンタスクをアーカイブシートへ移動する
 * @returns {string} 処理結果のメッセージ
 */
function archiveKanbanTasks() {
  const props = PropertiesService.getScriptProperties();
  
  // バッチ中はポータルへのアクセスを防ぐため一時的にメンテナンスモードへ
  props.setProperty('MAINTENANCE_MODE', 'true');
  
  try {
    const ss = getKanbanSpreadsheet();
    const sheet = ss.getSheetByName('TasksDB');
    if (!sheet) return "アーカイブするデータがありません。";

    let archiveSheet = ss.getSheetByName('TasksDB_Archive');
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet('TasksDB_Archive');
      const headers = sheet.getRange(1, 1, 1, 9).getValues();
      archiveSheet.appendRow(headers[0]);
      archiveSheet.getRange(1, 1, 1, 9).setBackground('#f3f3f3').setFontWeight('bold');
      archiveSheet.setFrozenRows(1);
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return "完了したタスクはありません。";

    const rowsToDelete = [];
    const rowsToArchive = [];

    // 行削除時のインデックスのズレを防ぐため、下からループ
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][3] === 'done') {
        rowsToArchive.push(data[i]);
        rowsToDelete.push(i + 1); 
      }
    }

    if (rowsToArchive.length > 0) {
      archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive.reverse());
      
      rowsToDelete.forEach(rowNum => {
        sheet.deleteRow(rowNum);
      });
      SpreadsheetApp.flush();
    }
    
    try { CacheService.getScriptCache().remove('KANBAN_TASKS_CACHE'); } catch(e){}
    
    return `${rowsToArchive.length}件の完了タスクをアーカイブしました。`;
  } catch(e) {
    throw new Error("アーカイブ処理中にエラー: " + e.message);
  } finally {
    props.setProperty('MAINTENANCE_MODE', 'false');
  }
}

/**
 * 自動アーカイブの実行スケジュール構成を取得する
 * @returns {Object} スケジュール構成オブジェクト
 */
function getArchiveSchedule() {
  const props = PropertiesService.getScriptProperties();
  return {
    enabled: props.getProperty('AUTO_ARCHIVE_ENABLED') === 'true',
    dow: parseInt(props.getProperty('AUTO_ARCHIVE_DOW') || '0', 10),
    hh: parseInt(props.getProperty('AUTO_ARCHIVE_HH') || '2', 10),
    mm: parseInt(props.getProperty('AUTO_ARCHIVE_MM') || '0', 10),
    wfDays: parseInt(props.getProperty('AUTO_ARCHIVE_WF_DAYS') || '90', 10)
  };
}

/**
 * 自動アーカイブスケジュールを保存し、トリガーを再設定する
 * @param {Object} config - 保存するスケジュール構成
 * @returns {boolean} 成功時
 */
function saveArchiveSchedule(config) {
  if (!isAdminUser()) throw new Error("権限がありません。");
  const props = PropertiesService.getScriptProperties();
  props.setProperty('AUTO_ARCHIVE_ENABLED', config.enabled ? 'true' : 'false');
  props.setProperty('AUTO_ARCHIVE_DOW', config.dow.toString());
  props.setProperty('AUTO_ARCHIVE_HH', config.hh.toString());
  props.setProperty('AUTO_ARCHIVE_MM', config.mm.toString());
  props.setProperty('AUTO_ARCHIVE_WF_DAYS', config.wfDays.toString()); 

  setupNextArchiveTrigger();
  return true;
}

/**
 * 次回のアーカイブ処理を実行する日時を計算する
 * @param {number} dow - 曜日 (0-6)
 * @param {number} hh - 時
 * @param {number} mm - 分
 * @returns {Date} 次回実行日時のDateオブジェクト
 */
function calculateNextArchiveDate(dow, hh, mm) {
  const now = new Date();
  let nextDate = new Date();
  nextDate.setHours(hh, mm, 0, 0);

  let daysToAdd = dow - now.getDay();
  // 計算結果が過去になる場合は次週にシフト
  if (daysToAdd < 0 || (daysToAdd === 0 && nextDate <= now)) {
    daysToAdd += 7;
  }
  nextDate.setDate(nextDate.getDate() + daysToAdd);
  return nextDate;
}

/**
 * 既存のアーカイブ用トリガーを削除し、正確な次回実行用の特定日時トリガーをセットする
 */
function setupNextArchiveTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runAutoArchiveBatch') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  const schedule = getArchiveSchedule();
  if (!schedule.enabled) return;

  const nextDate = calculateNextArchiveDate(schedule.dow, schedule.hh, schedule.mm);
  ScriptApp.newTrigger('runAutoArchiveBatch')
    .timeBased()
    .at(nextDate)
    .create();
}

/**
 * 自動起動されるアーカイブバッチの本体
 * 完了後に次回用のトリガーを自己セットアップします。
 */
function runAutoArchiveBatch() {
  const props = PropertiesService.getScriptProperties();
  try {
    archiveKanbanTasks();
    
    props.setProperty('MAINTENANCE_MODE', 'true');
    const schedule = getArchiveSchedule();
    if (schedule.wfDays > 0) {
      archiveWorkflowTasks(schedule.wfDays);
    }
  } catch(e) {
    console.error("自動アーカイブ実行エラー:", e);
  } finally {
    props.setProperty('MAINTENANCE_MODE', 'false');
    setupNextArchiveTrigger();
  }
}

/**
 * 管理画面等から手動で呼び出されるアーカイブの一括実行処理
 * @returns {string} 処理結果メッセージ
 */
function runManualArchiveBatch() {
  if (!isAdminUser()) throw new Error(getMsg('errUnauthorized'));
  const props = PropertiesService.getScriptProperties();
  try {
    const kanbanMsg = archiveKanbanTasks();
    
    props.setProperty('MAINTENANCE_MODE', 'true');
    const schedule = getArchiveSchedule();
    const wfDays = schedule.wfDays || 90;
    const wfCount = archiveWorkflowTasks(wfDays) || 0;
    
    return `${kanbanMsg}\nワークフロー: ${wfCount}件の完了チケットをアーカイブしました。`;
  } catch(e) {
    throw new Error("アーカイブ処理中にエラー: " + e.message);
  } finally {
    props.setProperty('MAINTENANCE_MODE', 'false');
  }
}

/**
 * 管理画面の設定（検索、アーカイブ、リンクマスタ）を一括で保存する
 * @param {boolean} searchEnabled - Cloud Searchの有効状態
 * @param {Object} archiveConfig - アーカイブ構成
 * @param {Array} newLinks - 新しいリンクマスタ
 * @returns {Array} 保存後のリンクマスタ配列
 */
function saveAllAdminSettings(searchEnabled, archiveConfig, newLinks) {
  if (!isAdminUser()) throw new Error("権限がありません。");
  
  try {
    const props = PropertiesService.getScriptProperties();
    
    props.setProperty('ENABLE_CLOUD_SEARCH', searchEnabled ? 'true' : 'false');
    try { CacheService.getScriptCache().put('CLOUD_SEARCH_SETTING_CACHE', searchEnabled ? 'true' : 'false', 21600); } catch(e){}

    props.setProperty('AUTO_ARCHIVE_ENABLED', archiveConfig.enabled ? 'true' : 'false');
    props.setProperty('AUTO_ARCHIVE_DOW', archiveConfig.dow.toString());
    props.setProperty('AUTO_ARCHIVE_HH', archiveConfig.hh.toString());
    props.setProperty('AUTO_ARCHIVE_MM', archiveConfig.mm.toString());
    props.setProperty('AUTO_ARCHIVE_WF_DAYS', archiveConfig.wfDays.toString()); 
    
    try {
      setupNextArchiveTrigger(); 
    } catch(triggerError) {
      throw new Error("トリガー設定エラー: " + triggerError.message);
    }

    const sheet = SpreadsheetApp.openById(getSpreadsheetId()).getSheetByName('StaticLinks');
    if(!sheet) throw new Error("「StaticLinks」シートが見つかりません。");

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
    if (newLinks && newLinks.length > 0) {
      const outputData = newLinks.map(link => [link.title, link.url, link.icon || 'link']);
      sheet.getRange(2, 1, outputData.length, 3).setValues(outputData);
    }
    CacheService.getScriptCache().remove('MASTER_LINKS_CACHE');
    
    return getMasterLinks(); 
  } catch (e) {
    throw new Error("バックエンド処理エラー[" + e.lineNumber + "行目]: " + e.message + "\n" + e.stack);
  }
}

/**
 * 規定日数を経過した完了ワークフローのデータをアーカイブシートへ退避する
 * @param {number} daysThreshold - 完了からの経過日数閾値
 * @returns {number} アーカイブしたチケットの件数
 */
function archiveWorkflowTasks(daysThreshold) {
  if (!daysThreshold || daysThreshold <= 0) return;
  
  const ss = getSpreadsheetId() ? SpreadsheetApp.openById(getSpreadsheetId()) : null;
  if (!ss) return;

  const tSheet = ss.getSheetByName('WF_Tickets');
  const trSheet = ss.getSheetByName('WF_TicketRoutes');
  const logSheet = ss.getSheetByName('WF_ApprovalLogs');
  if (!tSheet || !trSheet || !logSheet) return;

  const tData = tSheet.getDataRange().getValues();
  const trData = trSheet.getDataRange().getValues();
  const logData = logSheet.getDataRange().getValues();
  if (tData.length <= 1) return;

  const now = new Date();
  const thresholdDate = new Date(now);
  thresholdDate.setDate(thresholdDate.getDate() - daysThreshold);

  const completedStatuses = ['APPROVED', 'REJECTED', 'RETURNED'];
  const targetTicketIds = [];
  const rowsToDeleteT = [];
  
  // 各チケットの最終アクション日時をログから計算して特定する
  const lastActionMap = {};
  for (let i = 1; i < logData.length; i++) {
    const tId = logData[i][0];
    const aDate = new Date(logData[i][3]);
    if (!lastActionMap[tId] || aDate > lastActionMap[tId]) lastActionMap[tId] = aDate;
  }

  const rowsToArchiveT = [];
  
  // 下から順に処理し、インデックスズレを回避
  for (let i = tData.length - 1; i >= 1; i--) {
    const row = tData[i];
    const ticketId = row[0];
    const status = row[10];
    
    if (completedStatuses.includes(status)) {
      const lastDate = lastActionMap[ticketId] || new Date(row[4]);
      if (lastDate < thresholdDate) {
        targetTicketIds.push(ticketId);
        rowsToArchiveT.push(row);
        rowsToDeleteT.push(i + 1);
      }
    }
  }

  if (targetTicketIds.length === 0) return 0;

  const rowsToArchiveTR = [], rowsToDeleteTR = [];
  for (let i = trData.length - 1; i >= 1; i--) {
    if (targetTicketIds.includes(trData[i][0])) {
      rowsToArchiveTR.push(trData[i]);
      rowsToDeleteTR.push(i + 1);
    }
  }

  const rowsToArchiveLog = [], rowsToDeleteLog = [];
  for (let i = logData.length - 1; i >= 1; i--) {
    if (targetTicketIds.includes(logData[i][0])) {
      rowsToArchiveLog.push(logData[i]);
      rowsToDeleteLog.push(i + 1);
    }
  }

  // 内部ヘルパー関数：該当データをアーカイブ先へ追加し、元から削除
  function moveData(sheetName, originalHeaders, archiveRows, deleteRowNums) {
    if (archiveRows.length === 0) return;
    let aSheet = ss.getSheetByName(sheetName + '_Archive');
    if (!aSheet) {
      aSheet = ss.insertSheet(sheetName + '_Archive');
      aSheet.appendRow(originalHeaders);
      aSheet.getRange(1, 1, 1, originalHeaders.length).setBackground('#f3f3f3').setFontWeight('bold');
      aSheet.setFrozenRows(1);
    }
    aSheet.getRange(aSheet.getLastRow() + 1, 1, archiveRows.length, archiveRows[0].length).setValues(archiveRows.reverse());
    const originalSheet = ss.getSheetByName(sheetName);
    deleteRowNums.forEach(r => originalSheet.deleteRow(r));
  }

  moveData('WF_Tickets', tData[0], rowsToArchiveT, rowsToDeleteT);
  moveData('WF_TicketRoutes', trData[0], rowsToArchiveTR, rowsToDeleteTR);
  moveData('WF_ApprovalLogs', logData[0], rowsToArchiveLog, rowsToDeleteLog);

  SpreadsheetApp.flush();
  return targetTicketIds.length;
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
  const email = Session.getActiveUser().getEmail();

  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const ticketArchiveSheet = ss.getSheetByName('WF_Tickets_Archive');
  const stepArchiveSheet = ss.getSheetByName('WF_TicketRoutes_Archive');
  if (!ticketArchiveSheet) return []; 

  const tickets = ticketArchiveSheet.getDataRange().getValues().slice(1);
  const steps = stepArchiveSheet ? stepArchiveSheet.getDataRange().getValues().slice(1) : [];
  const masterStepSheet = ss.getSheetByName('WF_RouteStepMaster');
  const masterSteps = masterStepSheet ? masterStepSheet.getDataRange().getValues().slice(1) : [];

  const fromTime = dateFrom ? new Date(dateFrom).getTime() : 0;
  const toTime = dateTo ? new Date(dateTo).getTime() + 86400000 : Infinity; 
  const q = query ? query.toLowerCase() : "";

  const myTickets = [];

  tickets.forEach(t => {
    const ticketId = t[0], routeId = t[1], subject = t[2] || "", applicant = t[3];
    const appliedAt = new Date(t[4]);
    const appliedTime = appliedAt.getTime();

    if (applicant === email) {
      if (q && !subject.toLowerCase().includes(q)) return;
      if (appliedTime < fromTime || appliedTime > toTime) return;

      let routeSteps = steps.filter(s => s[0] === ticketId);
      if (routeSteps.length === 0) routeSteps = masterSteps.filter(s => s[0] === routeId);
      const totalSteps = routeSteps.length > 0 ? Math.max(...routeSteps.map(s => Number(s[1]))) : 1;
      
      let files = [];
      try { 
        const parsed = JSON.parse(t[6] || '[]'); 
        files = parsed.map(f => {
          if (typeof f === 'string') {
            try { return { id: f, name: DriveApp.getFileById(f).getName() }; } catch(e) { return { id: f, name: 'ファイルを開く' }; }
          }
          return f;
        });
      } catch(e){}
      
      myTickets.push({
        ticketId, routeId, subject, applicant, 
        appliedAt: Utilities.formatDate(appliedAt, "JST", "yyyy/MM/dd HH:mm"), 
        body: t[5], files, folderId: t[7], currentStep: Number(t[9]), totalSteps, status: t[10],
        actionType: 'READ_ONLY',
        isArchived: true
      });
    }
  });
  
  // 取得した全件を降順にソートし、offsetとlimitで切り取って返す（ページネーション）
  return myTickets.sort((a, b) => new Date(b.appliedAt) - new Date(a.appliedAt)).slice(offset, offset + limit);
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
