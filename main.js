function saveGmailToSheetBySenderWithDomainFilter() {
  const start = new Date();

  // 環境変数取得
  const props = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty('SPREADSHEET_ID');
  const sheetManager = SpreadsheetApp.openById(sheetId);
  const groupBaseUrl = props.getProperty('GROUP_BASE_URL');

  /**
   * ===== 初期化処理 =====
   * LAST_PROCESSED_THREAD_TIME が未設定の場合、
   * 現在時刻 − 安全マージン（10分）で初期化する
   */
  const SAFETY_MARGIN_MS = 10 * 60 * 1000; // 10分
  if (!props.getProperty('LAST_PROCESSED_THREAD_TIME')) {
    const now = Date.now();
    props.setProperty(
      'LAST_PROCESSED_THREAD_TIME',
      String(now - SAFETY_MARGIN_MS)
    );
    Logger.log('[初期化] LAST_PROCESSED_THREAD_TIME を現在-10分で初期化');
  }

  // 処理済みメッセージID履歴（短期キャッシュ）
  let messageIdHistory = JSON.parse(
    props.getProperty('MESSAGE_ID_HISTORY') || '[]'
  );
  const messageIdHistorySet = new Set(messageIdHistory);

  // 最後に処理したスレッドの最終更新時刻（ms）
  const lastProcessedMs = Number(
    props.getProperty('LAST_PROCESSED_THREAD_TIME')
  );

  /**
   * Gmail検索条件
   * - LAST_PROCESSED_THREAD_TIME − 安全マージン を基準に検索
   * - Gmail の after: は UNIX 秒指定
   */
  const searchBaseMs = Math.max(0, lastProcessedMs - SAFETY_MARGIN_MS);
  const afterSeconds = Math.floor(searchBaseMs / 1000);
  const query = `in:inbox after:${afterSeconds}`;

  Logger.log(`[差分取得] Gmail検索クエリ: ${query}`);

  /**
   * 新着メッセージを正確に扱うため、
   * 「最新スレッド」を少数（最大11件）のみ取得する
   */
  const threads = GmailApp.search(query, 0, 11);
  Logger.log(`[差分取得] 取得したスレッド数: ${threads.length}`);

  /**
   * シート分類ルール（シート分類管理シートから読み込み）
   * - key: シート名
   * - value: ドメインの配列
   */
  const categoryDomainMap = loadCategoryDomainMapFromSheet(sheetManager);

  /**
   * ドメイン → シート名の O(1) マップを事前生成
   */
  const domainToSheetMap = {};
  Object.entries(categoryDomainMap).forEach(([sheetName, domains]) => {
    domains.forEach(domain => {
      domainToSheetMap[domain.toLowerCase()] = sheetName;
    });
  });

  Logger.log(`[ドメインマップ] 生成されたドメイン数: ${Object.keys(domainToSheetMap).length}`);
  Logger.log(`[ドメインマップ] ドメイン一覧: ${Object.keys(domainToSheetMap).join(', ')}`);

  /**
   * - MESSAGE_ID_HISTORY が欠落しても重複出力しないための最終防衛
   * - シート名 → Set<msgId>
   */
  const existingMsgIdSetBySheet = {};

  // シート名を一意化して初期ロードを1回にする
  const uniqueSheetNames = [...new Set(Object.values(domainToSheetMap))];

  uniqueSheetNames.forEach(sheetName => {
    const sheet = sheetManager.getSheetByName(sheetName);
    if (!sheet) return;

    const gasInfo = getGasColumnInfo(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    const ids = sheet
      .getRange(2, gasInfo.startColumn + 6, lastRow - 1, 1)
      .getValues()
      .flat()
      .filter(Boolean);

    existingMsgIdSetBySheet[sheetName] = new Set(ids);

    Logger.log(
      `[初期ロード] シート「${sheetName}」既存msgId件数: ${ids.length}`
    );
  });

  /**
   * バッチ書き込み用バッファ
   * sheetName: rowData[][]
   */
  const sheetWriteBuffer = {};

  /**
   * 処理したすべてのメッセージIDを収集（スキップしたものも含む）
   */
  const processedIds = [];

  let savedCount = 0;
  let skippedByHistory = 0;

  // 今回処理したスレッドの最終更新時刻（最大値）
  let maxThreadLastMs = lastProcessedMs;

  /**
   * スレッド単位で処理
   */
  threads.forEach((thread, index) => {
    const threadId = thread.getId();
    const threadLastMs = thread.getLastMessageDate().getTime();
    maxThreadLastMs = Math.max(maxThreadLastMs, threadLastMs);

    if (threadLastMs <= searchBaseMs) {
      Logger.log(
        `[差分取得] スレッド[${index + 1}] スキップ（基準時刻以前）`
      );
      return;
    }

    const messages = thread.getMessages();
    Logger.log(
      `[差分取得] スレッド[${index + 1}] ID=${threadId}, メッセージ数=${messages.length}`
    );

    // 最新 → 過去へ逆順
    for (let i = messages.length - 1; i >= 0; i--) {
      const msg = messages[i];
      const msgId = msg.getId();

      processedIds.push(msgId);

      /**
       * 短期履歴による重複防止
       */
      if (messageIdHistorySet.has(msgId)) {
        skippedByHistory++;
        Logger.log(
          `[重複チェック] 履歴ヒットで打ち切り: メッセージID=${msgId}`
        );
        break;
      }

      // 本文全文を取得（転送元と日時の抽出のため）
      const fullBody = msg.getPlainBody();

      // 転送元の名前、アドレス、日時を抽出
      const forwardedName = extractForwardedName(fullBody);
      const forwardedFrom = extractForwardedFrom(fullBody);
      const forwardedDate = extractForwardedDate(fullBody);

      Logger.log(`[メッセージ処理] msgId=${msgId}, forwardedFrom=${forwardedFrom}, forwardedDate=${forwardedDate}`);

      // 転送元または日時が抽出できない場合はスキップ
      if (!forwardedFrom || !forwardedDate) {
        Logger.log(`[メッセージ処理] スキップ: 転送元または日時が抽出できませんでした`);
        continue;
      }

      const atIndex = forwardedFrom.indexOf('@');
      if (atIndex === -1) {
        Logger.log(`[メッセージ処理] スキップ: @記号が見つかりませんでした: ${forwardedFrom}`);
        continue;
      }

      const domain = forwardedFrom
        .substring(atIndex + 1)
        .toLowerCase()
        .trim();

      Logger.log(`[メッセージ処理] 抽出したドメイン: ${domain}`);

      const sheetName = domainToSheetMap[domain];
      if (!sheetName) {
        Logger.log(`[メッセージ処理] スキップ: ドメイン「${domain}」に対応するシートが見つかりませんでした`);
        Logger.log(`[メッセージ処理] 利用可能なドメイン: ${Object.keys(domainToSheetMap).join(', ')}`);
        continue;
      }

      Logger.log(`[メッセージ処理] マッチしたシート名: ${sheetName}`);

      /**
       * すでにシートに存在する msgId は絶対に再出力しない
       */
      if (existingMsgIdSetBySheet[sheetName]?.has(msgId)) {
        Logger.log(
          `[重複防止] シートに既に存在するためスキップ: msgId=${msgId}`
        );
        continue;
      }

      if (!sheetWriteBuffer[sheetName]) {
        sheetWriteBuffer[sheetName] = [];
      }

      const subject = msg.getSubject();
      const sender = forwardedName || forwardedFrom || msg.getFrom();
      const hasAttachment = msg.getAttachments().length > 0;

      const safeSubject = subject.replace(/"/g, '\\"');
      const hasSpace = subject.includes(' ');
      const groupQuery = hasSpace
        ? `"${safeSubject}"${hasAttachment ? ' has:attachment' : ''}`
        : `subject:(${safeSubject})${hasAttachment ? ' has:attachment' : ''}`;

      const groupThreadUrl =
        groupBaseUrl + encodeURIComponent(groupQuery);
      const hyperlink = hasAttachment
        ? `=HYPERLINK("${groupThreadUrl}", "あり")`
        : 'なし';

      const body = extractBodyPreview(fullBody);

      sheetWriteBuffer[sheetName].push([
        forwardedDate,
        sender,
        subject,
        body,
        hyperlink,
        threadId,
        msgId
      ]);

      savedCount++;
    }
  });

  /**
   * バッチ書き込み
   */
  Object.entries(sheetWriteBuffer).forEach(([sheetName, rows]) => {
    if (rows.length === 0) return;

    // 本文中の「日時（Date）」を基準に降順ソート
    rows.sort((a, b) => b[0].getTime() - a[0].getTime());

    let sheet = sheetManager.getSheetByName(sheetName);
    if (!sheet) {
      sheet = sheetManager.insertSheet(sheetName);
      ensureGasHeaders(sheet);
      protectAutoGeneratedColumns(sheet);
    }

    const gasInfo = getGasColumnInfo(sheet);
    sheet.insertRowsBefore(2, rows.length);
    sheet
      .getRange(2, gasInfo.startColumn, rows.length, gasInfo.columnCount)
      .setValues(rows);
  });

  /**
   * メッセージID履歴の更新（短期）
   */
  processedIds.forEach(id => messageIdHistorySet.add(id));

  const HISTORY_MAX = 5000;
  let updatedHistory = Array.from(messageIdHistorySet);
  if (updatedHistory.length > HISTORY_MAX) {
    updatedHistory = updatedHistory.slice(
      updatedHistory.length - HISTORY_MAX
    );
  }

  props.setProperty('MESSAGE_ID_HISTORY', JSON.stringify(updatedHistory));

  /**
   * 最終処理スレッド時刻の更新
   */
  props.setProperty(
    'LAST_PROCESSED_THREAD_TIME',
    String(maxThreadLastMs)
  );

  Logger.log(
    `[処理完了] 保存件数=${savedCount}, 履歴スキップ=${skippedByHistory}, 実行時間=${((new Date()) - start) / 1000}s`
  );
}

const GAS_HEADERS = ["送信日時","送信者","件名","本文","添付の有無","スレッドID","メッセージID"];
const INITIAL_HEADERS = ["対応確認","備考","対応者",...GAS_HEADERS];

/**
 * GASで自動生成するメール情報列（`GAS_HEADERS`）が
 * シート上のどの列から何列分連続して存在しているかを特定する。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象のシート
 * @returns {{ startColumn: number, columnCount: number }} 開始列番号と列数
 * @throws {Error} GAS専用ヘッダーが1行目に見つからない場合
 */
function getGasColumnInfo(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const startIndex = headers.findIndex(h => h === GAS_HEADERS[0]);
  if (startIndex === -1) throw new Error("GAS専用ヘッダーが見つかりません");
  return { startColumn: startIndex + 1, columnCount: GAS_HEADERS.length };
}

/**
 * 対象シートの1行目に、手動入力用の列＋GAS自動生成用の列ヘッダーを保証する。
 * - シートが空のときは `INITIAL_HEADERS` を丸ごとセットする。
 * - 既に列がある場合は、GAS専用ヘッダーが存在しなければ末尾に `GAS_HEADERS` を追加する。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象のシート
 */
function ensureGasHeaders(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    sheet.getRange(1, 1, 1, INITIAL_HEADERS.length).setValues([INITIAL_HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  if (headers.includes(GAS_HEADERS[0])) return;
  sheet.getRange(1, lastColumn + 1, 1, GAS_HEADERS.length).setValues([GAS_HEADERS]);
  sheet.setFrozenRows(1);
}

/**
 * GASが自動で書き込むメール情報列の範囲に保護を設定し、
 * 手動編集できないようにすることでデータ破壊を防ぐ。
 * - 対象範囲はヘッダー行を含む全行の自動生成列。
 * - 実行ユーザーだけが編集できる完全保護＋背景色を設定する。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象のシート
 */
function protectAutoGeneratedColumns(sheet) {
  const gasInfo = getGasColumnInfo(sheet);
  const range = sheet.getRange(1, gasInfo.startColumn, sheet.getMaxRows(), gasInfo.columnCount);
  const protection = range.protect().setDescription('Auto Mail Data Protected');
  const me = Session.getEffectiveUser().getEmail();
  protection.removeEditors(protection.getEditors());
  protection.addEditor(me);
  protection.setWarningOnly(false);
  range.setBackground("#f3f3f3");
}

/**
 * 転送メール本文から「転送元の名前」を抽出する。
 * - 通常テキスト形式と、太字（`**転送元の名前:**`）形式の両方に対応する。
 *
 * @param {string} body Gmailメッセージのプレーンテキスト本文
 * @returns {string|null} 転送元の名前。見つからない場合は `null`
 */
function extractForwardedName(body) {
  if (!body) return null;
  let match = body.match(/転送元の名前:\s*([^\n]+)/);
  if (match) return match[1].trim();
  match = body.match(/\*{1,2}転送元の名前:\*{1,2}\s*([^\n]+)/);
  return match ? match[1].trim() : null;
}

/**
 * 転送メール本文から「転送元アドレス:xxx@example.com」の形式で抽出する。
 *
 * @param {string} body Gmailメッセージのプレーンテキスト本文
 * @returns {string|null} 転送元メールアドレス。見つからない場合は `null`
 */
function extractForwardedFrom(body) {
  if (!body) return null;
  let match = body.match(/転送元アドレス:\s*([^\s\n]+@[^\s\n]+)/);
  if (match) {
    const addr = match[1].trim();
    if (addr) return addr;
  }  
  return null;
}

/**
 * 転送メール本文から「元メールの送信日時」を抽出して `Date` オブジェクトとして返す。
 * - 通常テキスト形式と、太字（`**日時:**`）形式の両方に対応する。
 *
 * @param {string} body Gmailメッセージのプレーンテキスト本文
 * @returns {Date|null} 解析に成功した場合は日時、失敗した場合は `null`
 */
function extractForwardedDate(body) {
  if (!body) return null;
  let match = body.match(/日時:\s*([^\n]+)/);
  if (match) {
    const d = new Date(match[1].trim());
    return isNaN(d.getTime()) ? null : d;
  }
  match = body.match(/\*{1,2}日時:\*{1,2}\s*([^\n]+)/);
  if (match) {
    const d = new Date(match[1].trim());
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

/**
 * メール本文からプレビュー用の本文（最大50文字）を抽出する。
 * `---` 区切りがある場合は、その後のテキストを優先する。
 *
 * @param {string} fullBody メール本文の全文
 * @returns {string} 最大50文字の本文プレビュー（改行は空白に変換）
 */
function extractBodyPreview(fullBody) {
  let body = fullBody;
  const borderIndex = body.indexOf('---');
  if (borderIndex !== -1) {
    let startIndex = borderIndex + 3;
    while (startIndex < body.length && /[\s\n\r]/.test(body[startIndex])) startIndex++;
    body = body.substring(startIndex, startIndex + 50);
  } else {
    body = body.substring(0, 50);
  }
  return body.replace(/\r?\n/g, ' ');
}

/**
 * 「シート分類管理」シートを初期化する。
 * - シートが存在しない場合は作成
 * - 1行目1列目に「【使い方ガイド】」を設定（フォントサイズ11、太字、ラッピングはみ出し）
 * - 2行目: 使い方説明文1（8行目以降に1行1組で入力する旨）
 * - 3行目: 使い方説明文2（1行1ドメイン、フィルター利用、重複は無視される旨）
 * - 4行目: 使い方説明文3（行削除時は次回以降そのドメインのメールはスキップされる旨）
 * - 5行目: 使い方説明文4（灰色エリアは編集禁止の旨、赤・太字）
 * - 6行目: 空行
 * - 1〜7行目: 背景色 #f3f3f3（データ入力行は背景なし）
 * - 7行目1列目に「シート名」を設定（太字、背景色薄い青）
 * - 7行目2列目に「ドメイン名」を設定（太字、背景色薄い青）
 * - 7行目を固定行として設定
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sheetManager スプレッドシートオブジェクト
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} シート分類管理シート
 */
function ensureCategoryManagementSheet(sheetManager) {
  const SHEET_NAME = 'シート分類管理';
  const HEADER_TITLE = '【使い方ガイド】';
  const USAGE_TEXT_1 = '・本シートでは、メールの送信元ドメインに対応する「出力先シート名」を設定します。8行目以降にシート名とドメイン名を1行1組で入力してください。';
  const USAGE_TEXT_2 = '・ドメイン名は1行に1つ入力します。同じシート名・ドメイン名の重複行は読み込み時に無視されます。';
  const USAGE_TEXT_3 = '・「シート名・ドメイン名」の行を削除すると、次回実行以降はそのドメインのメールはどのシートにも書かれずスキップされます。';
  const USAGE_TEXT_4 = '・背景が灰色のエリアは編集禁止です。（行、列の追加も禁止）';
  const HEADERS = ['シート名', 'ドメイン名'];
  const LIGHT_BLUE = '#BBDEFB';
  const GUIDE_GRAY = '#f3f3f3';
  const FONT_RED = '#cc0000';
  
  let sheet = sheetManager.getSheetByName(SHEET_NAME);
  
  // シートが既に存在し、ヘッダー行（7行目）も正しく設定されている場合は早期リターン
  if (sheet && sheet.getRange(7, 1).getValue() === HEADERS[0]) {
    return sheet;
  }
  
  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = sheetManager.insertSheet(SHEET_NAME);
    Logger.log(`[シート分類管理] シート「${SHEET_NAME}」を作成しました`);
  }
  
  // 1行目1列目に「【使い方ガイド】」を設定
  const titleRange = sheet.getRange(1, 1);
  if (titleRange.getValue() !== HEADER_TITLE) {
    titleRange.setValue(HEADER_TITLE);
    titleRange.setFontSize(11);
    titleRange.setFontWeight('bold');
    titleRange.setWrap(false);
    Logger.log(`[シート分類管理] 1行目のタイトルを設定しました`);
  }
  
  // 2行目1列目に使い方説明文1を設定
  const usageRange1 = sheet.getRange(2, 1);
  if (usageRange1.getValue() !== USAGE_TEXT_1) {
    usageRange1.setValue(USAGE_TEXT_1);
    usageRange1.setWrap(false);
    Logger.log(`[シート分類管理] 2行目の説明文を設定しました`);
  }
  
  // 3行目1列目に使い方説明文2を設定
  const usageRange2 = sheet.getRange(3, 1);
  if (usageRange2.getValue() !== USAGE_TEXT_2) {
    usageRange2.setValue(USAGE_TEXT_2);
    usageRange2.setWrap(false);
    Logger.log(`[シート分類管理] 3行目の説明文を設定しました`);
  }
  
  // 4行目1列目に使い方説明文3を設定
  const usageRange3 = sheet.getRange(4, 1);
  if (usageRange3.getValue() !== USAGE_TEXT_3) {
    usageRange3.setValue(USAGE_TEXT_3);
    usageRange3.setWrap(false);
    Logger.log(`[シート分類管理] 4行目の説明文を設定しました`);
  }
  
  // 5行目に使い方説明文4を設定（赤・太字）
  const usageRange4 = sheet.getRange(5, 1);
  if (usageRange4.getValue() !== USAGE_TEXT_4) {
    usageRange4.setValue(USAGE_TEXT_4);
    usageRange4.setFontColor(FONT_RED);
    usageRange4.setFontWeight('bold');
    usageRange4.setWrap(false);
    Logger.log(`[シート分類管理] 5行目の説明文を設定しました`);
  }
  
  // 6行目を空行にする
  sheet.getRange(6, 1).clearContent();
  
  // 7行目にヘッダー行を設定（既に設定されている場合はスキップ）
  const headerRange = sheet.getRange(7, 1);
  if (headerRange.getValue() !== HEADERS[0]) {
    sheet.getRange(7, 1, 1, HEADERS.length).setValues([HEADERS]);
    // ヘッダー行のスタイル設定
    const headerRow = sheet.getRange(7, 1, 1, HEADERS.length);
    headerRow.setFontWeight('bold');
    headerRow.setBackground(LIGHT_BLUE);
    Logger.log(`[シート分類管理] ヘッダー行を設定しました`);
  }
  
  // 7行目を固定（説明文とヘッダー行を固定）
  sheet.setFrozenRows(7);
  
  // 1行目〜7行目のA列〜Z列の背景色を灰色に設定（ガイド・空行・ヘッダー）
  sheet.getRange(1, 1, 7, 26).setBackground(GUIDE_GRAY);
  
  // データ入力行（8行目以降）は背景色なし
  sheet.getRange(8, 1, sheet.getMaxRows(), 2).setBackground(null);
  
  // A列とB列を保護（全行）
  const maxRows = sheet.getMaxRows();
  const range = sheet.getRange(1, 1, maxRows, 2);
  const rangeA1 = range.getA1Notation();
  
  // 既存の保護を削除（再設定のため）
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(protection => {
    const protectionRange = protection.getRange();
    // 対象範囲と一致する、または対象範囲を含む保護を削除
    if (protectionRange.getA1Notation() === rangeA1 || 
        (protectionRange.getRow() <= range.getRow() && 
         protectionRange.getLastRow() >= range.getLastRow() &&
         protectionRange.getColumn() <= range.getColumn() &&
         protectionRange.getLastColumn() >= range.getLastColumn())) {
      protection.remove();
    }
  });
  
  // 新しい保護を設定
  const protection = range.protect().setDescription('Category Management Sheet Protected');
  const me = Session.getEffectiveUser().getEmail();
  protection.removeEditors(protection.getEditors());
  protection.addEditor(me);
  protection.setWarningOnly(false);
  
  Logger.log(`[シート分類管理] A列とB列の保護を設定しました（実行ユーザー: ${me}）`);
  
  Logger.log(`[シート分類管理] シートの初期化が完了しました`);
  
  return sheet;
}

/**
 * 「シート分類管理」シートからドメイン情報を読み込み、
 * { シート名: [ドメイン1, ドメイン2, ...] } 形式のオブジェクトを返す。
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sheetManager スプレッドシートオブジェクト
 * @returns {Object<string, string[]>} シート名をキー、ドメイン配列を値とするオブジェクト
 */
function loadCategoryDomainMapFromSheet(sheetManager) {
  const SHEET_NAME = 'シート分類管理';
  
  // シートの初期化（存在しない場合は作成）
  const sheet = ensureCategoryManagementSheet(sheetManager);
  
  // データ行の確認（ヘッダー行（7行目）までしかない場合は空オブジェクトを返す）
  const lastRow = sheet.getLastRow();
  if (lastRow <= 7) {
    Logger.log(`[シート分類管理] シート「${SHEET_NAME}」にデータ行がないため空オブジェクトを返します`);
    return {};
  }
  
  // 8行目以降のデータを取得（1列目: シート名、2列目: ドメイン名・1行1ドメイン）
  const numRows = lastRow - 7;
  const dataRows = sheet.getRange(8, 1, numRows, 2).getValues();
  
  // 有効なデータ行をフィルタリング（空行を除外）
  const validRows = dataRows.filter(row => row[0] && row[1]);
  if (validRows.length === 0) {
    Logger.log(`[シート分類管理] 有効なデータ行がないため空オブジェクトを返します`);
    return {};
  }
  
  // categoryDomainMap を生成（1行 = 1シート名 + 1ドメイン）
  const categoryDomainMap = {};
  let processedCount = 0;
  
  validRows.forEach((row, index) => {
    const sheetName = String(row[0]).trim();
    const domain = String(row[1]).trim().toLowerCase();
    
    if (!sheetName || !domain) {
      Logger.log(`[シート分類管理] 行${index + 8}: シート名またはドメイン名が空のためスキップ`);
      return;
    }
    
    if (!categoryDomainMap[sheetName]) {
      categoryDomainMap[sheetName] = [];
    }
    
    if (!categoryDomainMap[sheetName].includes(domain)) {
      categoryDomainMap[sheetName].push(domain);
      processedCount++;
    } else {
      Logger.log(`[シート分類管理] スキップ（重複）: シート名=${sheetName}, ドメイン=${domain}`);
    }
  });
  
  Logger.log(`[シート分類管理] 読み込み完了: ${Object.keys(categoryDomainMap).length}シート, ${processedCount}ドメイン`);
  
  return categoryDomainMap;
}