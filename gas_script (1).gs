// ============================================================
// ラッkey chain展 GASシステム（イベントファイル用）
// 運用上の注意点・設計判断の背景は handover.md にまとめてあります。
// 特に導入時は「MASTER_FILE_ID の設定」「onInvoiceArtistChange のトリガー設定」
// 「InvoiceProductArea / InvoiceNoteCell の名前付き範囲設定」を必ず確認してください。
// ============================================================

// ===== シート名の設定 =====
const FORM_SHEET_PLAN     = 'ラッKey chain5納品予定回答';  // 納品予定フォーム回答シート
const FORM_SHEET_FINAL    = 'ラッKey chain5納品確定回答';  // 納品確定フォーム回答シート
const PRODUCT_SHEET_PLAN  = '商品一覧_予定';               // 予定の縦持ち展開先
const PRODUCT_SHEET_FINAL = '商品一覧_確定';               // 確定の縦持ち展開先
const INVOICE_SHEET_NAME  = '納品確認書';                   // 納品確認書シート
const LOG_SHEET_NAME      = '在庫変動ログ';                  // 在庫変動ログシート
const PRODUCT_CODE_SHEET  = '商品コード管理';                // 商品コード管理シート
const INVENTORY_SHEET     = '在庫管理(商品ごと)';            // 在庫管理シート（商品ごと）※旧「在庫管理」からリネーム
const INVENTORY_SHEET_BY_SKU = '在庫管理(SKUごと)';          // 在庫管理シート（SKUごと）
const SQUARE_SHEET_NAME   = 'Square売上';                  // Square売上データ貼り付け用シート
const ROYALTY_SHEET_NAME  = 'ロイヤリティレポート';           // ロイヤリティレポートシート
const ARTIST_MASTER_SHEET = '作家マスタ';                   // 作家マスタシート（第6回から別ファイルに分離）
const PERIOD_ARTIST_SHEET = '会期ごと作家情報';              // 会期管理シート（第6回から別ファイルに分離）

// ===== マスタファイルの設定（第6回から追加）=====
// 作家マスタ・会期ごと作家情報は、会期を並行運用する際の照合の手間をなくすため、
// この「イベントファイル（会期ごとにコピーして使うファイル）」から切り離し、
// 専用のマスタファイル1つに集約している。
// MASTER_FILE_IDはマスタファイルのURLに含まれるID文字列（例：
// https://docs.google.com/spreadsheets/d/【ここの部分】/edit ）を入れること。
// テンプレートファイルにこの値を一度設定しておけば、以降コピーして作る
// イベントファイルすべてに自動的に引き継がれる。
const MASTER_FILE_ID = '1U6TT97giGJtySBj5FswXkPiXsSsglIU7IXb2WSm17Xg';

// マスタファイルの「作家マスタ」シートを取得する（見つからない場合はnullを返す）
function getArtistMasterSheet() {
  const masterSs = SpreadsheetApp.openById(MASTER_FILE_ID);
  return masterSs.getSheetByName(ARTIST_MASTER_SHEET); // 見つからなければnull
}

// マスタファイルの「会期ごと作家情報」シートを取得する（見つからない場合はnullを返す）
function getPeriodArtistSheet() {
  const masterSs = SpreadsheetApp.openById(MASTER_FILE_ID);
  return masterSs.getSheetByName(PERIOD_ARTIST_SHEET); // 見つからなければnull
}


// ============================================================
// 全角数字→半角数字に変換する
// ============================================================
function toHalfWidth(val) {
  if (val === null || val === undefined) return '';
  return String(val).replace(/[０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
}


// ============================================================
// 納品予定フォームの列定義
// ============================================================
const PLAN_COL_EMAIL        = 1;
const PLAN_COL_ARTIST_NAME  = 2;
const PLAN_COL_REAL_NAME    = 3;
const PLAN_COL_ADDRESS      = 4;
const PLAN_COL_TEL          = 5;
const PLAN_PRODUCT_START    = 6;
const PLAN_COLS_PER_BLOCK   = 7;
const PLAN_OFFSET_NAME      = 0;
const PLAN_OFFSET_SAMPLE    = 1;
const PLAN_OFFSET_SAMPLE_OK = 2;
const PLAN_OFFSET_STOCK     = 3;
const PLAN_OFFSET_PRICE     = 4;
const PLAN_OFFSET_NOTE      = 5;


// ============================================================
// 納品確定フォームの列定義
// ============================================================
const FINAL_COL_ARTIST_NAME  = 2;
const FINAL_PRODUCT_START    = 3;
const FINAL_COLS_PER_BLOCK   = 7;
const FINAL_OFFSET_NAME      = 0;
const FINAL_OFFSET_SAMPLE    = 1;
const FINAL_OFFSET_SAMPLE_OK = 2;
const FINAL_OFFSET_STOCK     = 3;
const FINAL_OFFSET_PRICE     = 4;
const FINAL_OFFSET_NOTE      = 5;

const PRODUCT_BLOCKS = 20;


// ============================================================
// 共通：商品一覧シートを作成・更新する
// ============================================================
function buildProductSheet(sheetName, rows, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clearContents();
  sheet.clearFormats();

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4a86e8');
  headerRange.setFontColor('#ffffff');

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();
  sheet.autoResizeColumns(1, headers.length);
  ensureHeaderColumnWidth(sheet, headers, rows);

  // 交互背景色（縞模様）。1行ずつsetBackground()すると行数分サーバー通信が発生して
  // 重くなるため、全行分の色を2次元配列としてまとめて組み立て、setBackgrounds()で
  // 1回にまとめて書き込む。
  if (rows.length > 0) {
    const rowColors = [];
    for (let i = 0; i < rows.length; i++) {
      const bg = i % 2 === 0 ? '#ffffff' : '#e8f0fe';
      const colorRow = [];
      for (let c = 0; c < headers.length; c++) colorRow.push(bg);
      rowColors.push(colorRow);
    }
    sheet.getRange(2, 1, rows.length, headers.length).setBackgrounds(rowColors);
  }

  return sheet;
}


// ============================================================
// autoResizeColumns後、見出し・データの文字が収まらない列の幅を広げる
// 日本語1文字を約16px、半角英数字1文字を約8pxとして幅を見積もる
// rows省略時は見出しのみで判定する
// ============================================================
function ensureHeaderColumnWidth(sheet, headers, rows) {
  const MIN_PADDING = 24; // セルの左右余白分

  function estimateTextWidth(text) {
    let width = MIN_PADDING;
    const str = String(text || '');
    for (let c = 0; c < str.length; c++) {
      const code = str.charCodeAt(c);
      width += (code > 0x2E80) ? 16 : 8; // 全角相当 or 半角相当
    }
    return width;
  }

  for (let i = 0; i < headers.length; i++) {
    let estimatedWidth = estimateTextWidth(headers[i]);

    if (rows && rows.length > 0) {
      for (let r = 0; r < rows.length; r++) {
        const cellWidth = estimateTextWidth(rows[r][i]);
        if (cellWidth > estimatedWidth) estimatedWidth = cellWidth;
      }
    }

    const currentWidth = sheet.getColumnWidth(i + 1);
    if (currentWidth < estimatedWidth) {
      sheet.setColumnWidth(i + 1, estimatedWidth);
    }
  }
}


// ============================================================
// 納品予定フォーム回答を商品一覧_予定に展開する
// ============================================================
function updatePlanList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(FORM_SHEET_PLAN);

  if (!formSheet) {
    SpreadsheetApp.getUi().alert('シート「' + FORM_SHEET_PLAN + '」が見つかりません。\nシート名を確認してください。');
    return;
  }

  const formData = formSheet.getDataRange().getValues();
  if (formData.length <= 1) {
    SpreadsheetApp.getUi().alert('納品予定フォームの回答がまだありません。');
    return;
  }

  const headers = ['作家名', '本名', '住所', 'TEL', 'メールアドレス',
                   '作品名', 'サンプル数', 'サンプル販売可否',
                   '納品数', '予定税抜き価格', '税込価格（10%）', '備考'];
  const rows = [];

  for (let r = 1; r < formData.length; r++) {
    const row = formData[r];
    const artistName = row[PLAN_COL_ARTIST_NAME];
    const realName   = row[PLAN_COL_REAL_NAME];
    const address    = row[PLAN_COL_ADDRESS];
    const tel        = toHalfWidth(row[PLAN_COL_TEL]);
    const email      = row[PLAN_COL_EMAIL];

    for (let b = 0; b < PRODUCT_BLOCKS; b++) {
      const base = PLAN_PRODUCT_START + b * PLAN_COLS_PER_BLOCK;
      const productName = row[base + PLAN_OFFSET_NAME];
      if (!productName || String(productName).trim() === '') continue;

      const priceRaw = toHalfWidth(row[base + PLAN_OFFSET_PRICE]);
      const priceNum = parseFloat(priceRaw);
      const taxPrice = !isNaN(priceNum) ? Math.round(priceNum * 1.1) : '';

      rows.push([
        artistName, realName, address, tel, email,
        productName,
        toHalfWidth(row[base + PLAN_OFFSET_SAMPLE]),
        row[base + PLAN_OFFSET_SAMPLE_OK] || '',
        toHalfWidth(row[base + PLAN_OFFSET_STOCK]),
        priceRaw, taxPrice,
        row[base + PLAN_OFFSET_NOTE] || '',
      ]);
    }
  }

  if (rows.length === 0) {
    SpreadsheetApp.getUi().alert('展開できる商品データがありませんでした。');
    return;
  }

  // 作家名（0列目）で昇順ソート。フォーム回答順のままだと同じ作家の行が
  // 途中で分断され、addArtistBordersの太線が正しく引けなくなるため。
  // Array.sortは安定ソートなので、同じ作家内の作品の並び順（フォーム回答順）は保たれる。
  rows.sort(function(a, b) {
    return String(a[0]).localeCompare(String(b[0]), 'ja');
  });

  const planSheet = buildProductSheet(PRODUCT_SHEET_PLAN, rows, headers);
  addArtistBorders(planSheet, rows, headers);
  protectSheet(planSheet, '商品一覧_予定はスクリプトが自動管理します。手動編集不可。');
  SpreadsheetApp.getUi().alert('完了！\n' + rows.length + '件の予定商品データを展開しました。');
}


// ============================================================
// 納品確定フォーム回答を商品一覧_確定に展開する
// 個人情報は持たない（納品書には予定シートから直接引き継ぐ）
// ============================================================
// 商品一覧_確定を生成するコア処理。戻り値：生成したfinalSheet（データなしの場合はnull）
function buildFinalProductList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(FORM_SHEET_FINAL);

  if (!formSheet) {
    SpreadsheetApp.getUi().alert('シート「' + FORM_SHEET_FINAL + '」が見つかりません。\nシート名を確認してください。');
    return null;
  }

  const formData = formSheet.getDataRange().getValues();
  if (formData.length <= 1) {
    SpreadsheetApp.getUi().alert('納品確定フォームの回答がまだありません。');
    return null;
  }

  const headers = ['作家名', '作品名', 'サンプル数', 'サンプル販売可否',
                   '納品数', '確定税抜き価格', '税込価格（10%）', '備考'];
  const rows = [];

  for (let r = 1; r < formData.length; r++) {
    const row = formData[r];
    const artistName = String(row[FINAL_COL_ARTIST_NAME]).trim();

    for (let b = 0; b < PRODUCT_BLOCKS; b++) {
      const base = FINAL_PRODUCT_START + b * FINAL_COLS_PER_BLOCK;
      const productName = row[base + FINAL_OFFSET_NAME];
      if (!productName || String(productName).trim() === '') continue;

      const priceRaw = toHalfWidth(row[base + FINAL_OFFSET_PRICE]);
      const priceNum = parseFloat(priceRaw);
      const taxPrice = !isNaN(priceNum) ? Math.round(priceNum * 1.1) : '';

      rows.push([
        artistName,
        productName,
        toHalfWidth(row[base + FINAL_OFFSET_SAMPLE]),
        row[base + FINAL_OFFSET_SAMPLE_OK] || '',
        toHalfWidth(row[base + FINAL_OFFSET_STOCK]),
        priceRaw, taxPrice,
        row[base + FINAL_OFFSET_NOTE] || '',
      ]);
    }
  }

  if (rows.length === 0) {
    SpreadsheetApp.getUi().alert('展開できる商品データがありませんでした。');
    return null;
  }

  // 作家名（0列目）で昇順ソート。フォーム回答順のままだと同じ作家の行が
  // 途中で分断され、addArtistBordersの太線が正しく引けなくなるため。
  rows.sort(function(a, b) {
    return String(a[0]).localeCompare(String(b[0]), 'ja');
  });

  const finalSheet = buildProductSheet(PRODUCT_SHEET_FINAL, rows, headers);
  addArtistBorders(finalSheet, rows, headers);
  protectSheet(finalSheet, '商品一覧_確定はスクリプトが自動管理します。手動編集不可。');

  // 作家マスタに登録されていない作家名をチェックし、セルを薄い赤に色付け
  const missingArtists = checkArtistsAgainstMaster(finalSheet, rows);

  return { sheet: finalSheet, rowCount: rows.length, missingArtists: missingArtists };
}


// ============================================================
// 商品一覧_確定の作家名が作家マスタに存在するかチェックする
// 存在しない場合は該当セル（A列）を薄い赤に色付けし、作家名一覧を返す
// ============================================================
function checkArtistsAgainstMaster(sheet, rows) {
  const masterSheet = getArtistMasterSheet();
  const masterNames = new Set();

  if (masterSheet) {
    const masterData = masterSheet.getDataRange().getValues();
    for (let i = 1; i < masterData.length; i++) {
      const name = String(masterData[i][1] || '').trim();
      if (name) masterNames.add(normalizeToken(name));
    }
  }

  const missing = new Set();
  if (rows.length > 0) {
    // A列の現在の背景色（buildProductSheetで設定済みの縞模様）を1回で読み込み、
    // 未登録の行だけ薄い赤に差し替えてから、最後に1回でまとめて書き戻す。
    // 1行ずつsetBackground()すると行数分サーバー通信が発生して重くなるため。
    const colA = sheet.getRange(2, 1, rows.length, 1);
    const bgColors = colA.getBackgrounds();

    for (let i = 0; i < rows.length; i++) {
      const artist = String(rows[i][0] || '').trim();
      if (!artist) continue;
      if (!masterNames.has(normalizeToken(artist))) {
        missing.add(artist);
        bgColors[i][0] = '#f4cccc'; // 薄い赤
      }
    }
    colA.setBackgrounds(bgColors);
  }
  return Array.from(missing);
}

// メニュー用：商品一覧_確定の更新のみ（プルダウン・商品コード管理は更新しない）
function updateFinalListOnly() {
  const result = buildFinalProductList();
  if (!result) return;

  let message = '完了！\n' + result.rowCount + '件の確定商品データを展開しました。\n' +
    '※納品確認書プルダウン・商品コード管理は更新していません。\n' +
    '商品名のチェックが済んだら「② 納品確定を更新」を実行してください。';

  if (result.missingArtists.length > 0) {
    message += '\n\n⚠️ 作家マスタに見つからない作家名（セルを赤くしています）：\n' + result.missingArtists.join('\n');
  }

  SpreadsheetApp.getUi().alert(message);
}

// メニュー用：商品一覧_確定の更新＋プルダウン＋商品コード管理を一括更新
function updateFinalList() {
  const result = buildFinalProductList();
  if (!result) return;

  const finalSheet = result.sheet;

  // 続けて納品確認書プルダウンと商品コード管理を更新
  const dropdownCount = setupInvoiceDropdownCore(finalSheet);
  const codeResult = updateProductCodeSheetCore(finalSheet);

  let message = '完了！\n' +
    result.rowCount + '件の確定商品データを展開しました。\n' +
    '納品確認書プルダウン：' + dropdownCount + '名の作家名を設定しました。\n' +
    '商品コード管理：新規' + codeResult.newRowCount + '件を追記しました。';

  if (result.missingArtists.length > 0) {
    message += '\n\n⚠️ 商品一覧_確定で作家マスタに見つからない作家名（セルを赤くしています）：\n' + result.missingArtists.join('\n');
  }
  if (codeResult.missingArtistsInCode.length > 0) {
    message += '\n\n⚠️ 商品コード管理で作家マスタに見つからない作家名（A列を赤くしています）：\n' + codeResult.missingArtistsInCode.join('\n');
  }
  if (codeResult.duplicateList.length > 0) {
    message += '\n\n⚠️ 商品コード管理内で同じ商品（作家名＋作品名）に複数の行・複数の商品コードが振られています（F列を薄いオレンジにしています）：\n' +
      codeResult.duplicateList.join('\n') +
      '\n\nどちらか一方の行を削除・統合してください。放置すると在庫数が二重にカウントされる可能性があります。';
  }

  SpreadsheetApp.getUi().alert(message);
}


// ============================================================
// 納品確認書シートに作家選択プルダウンを設定する（確定データ基準）
// ============================================================
function setupInvoiceDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const finalSheet = ss.getSheetByName(PRODUCT_SHEET_FINAL);

  if (!finalSheet) {
    SpreadsheetApp.getUi().alert('先に「② 納品確定を更新」を実行してください。');
    return;
  }

  const count = setupInvoiceDropdownCore(finalSheet);
  if (count === -1) {
    SpreadsheetApp.getUi().alert('シート「' + INVOICE_SHEET_NAME + '」が見つかりません。');
    return;
  }
  if (count === 0) {
    SpreadsheetApp.getUi().alert('作家名が見つかりません。先に納品確定を更新してください。');
    return;
  }

  SpreadsheetApp.getUi().alert('完了！\nH3セルに' + count + '名の作家名プルダウンを設定しました。');
}

// 戻り値：設定した作家数（-1は納品確認書シートが見つからない）
function setupInvoiceDropdownCore(finalSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invoiceSheet = ss.getSheetByName(INVOICE_SHEET_NAME);
  if (!invoiceSheet) return -1;

  const data = finalSheet.getDataRange().getValues();
  const artistSet = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) artistSet.add(data[i][0]);
  }
  const artists = Array.from(artistSet).sort();

  if (artists.length === 0) return 0;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(artists, true)
    .setAllowInvalid(false)
    .build();
  invoiceSheet.getRange('H3').setDataValidation(rule);

  return artists.length;
}


// ============================================================
// 納品書を自動更新する（H3の作家名が変わったときに呼ばれる）
// ============================================================
function onEdit(e) {
  // フォーム回答シートが編集されたら黄色でマーク
  const editedSheet = e.range.getSheet().getName();
  if (editedSheet === FORM_SHEET_PLAN || editedSheet === FORM_SHEET_FINAL) {
    e.range.setBackground('#fff2cc');
    return;
  }

  // ※作家マスタの重複チェック（旧checkArtistDuplicate）は第6回からマスタファイル側の
  //   専用スクリプトに移管した。作家マスタがこのファイルに存在しなくなったため。

  // 在庫変動ログの作家名（B列）が編集されたら作品名（C列）の選択肢を更新
  if (editedSheet === LOG_SHEET_NAME) {
    updateLogProductDropdown(e);
    fillLogCheckDate(e);
  }

  // 商品コード管理の作家名（A列）・作品名（C列）が編集されたら、
  // 同じ組み合わせの行が既に存在していないか即座にチェックする
  // （商品コード管理は在庫変動ログと違ってこのファイル自身の中にあるので、
  //   簡易トリガーのままで問題なく他シートを参照できる）
  if (editedSheet === PRODUCT_CODE_SHEET) {
    checkProductCodeDuplicate(e);
  }

  // ※納品確認書H3（作家名プルダウン）の処理はここでは行わない。
  //   fillInvoice()はマスタファイルをopenByIdで開く処理を含んでおり、
  //   簡易トリガー（この関数）は「バインドされているファイル以外」への
  //   アクセスが常に禁止されているため、簡易トリガーの中で呼ぶとエラーも出さず
  //   黙って失敗する。そのため onInvoiceArtistChange という別関数に切り出し、
  //   GASエディタのトリガー画面から「インストール型トリガー」として手動設定する
  //   必要がある（onFormSubmitと同じ理由・同じ設定方法）。
  //   設定方法：GASエディタ→左サイドバーの時計アイコン（トリガー）→トリガーを追加→
  //   関数を選択「onInvoiceArtistChange」→イベントの種類「編集時」→保存
}

// ============================================================
// 商品コード管理：作家名（A列）・作品名（C列）が編集されたとき、
// 同じ組み合わせ（正規化して比較）の行が既に他にないかチェックする
// 複数行・複数セルの一括貼り付けにも対応
// ============================================================
function checkProductCodeDuplicate(e) {
  const sheet = e.range.getSheet();

  // 編集範囲がA列またはC列にかかっていない場合は対象外
  const startCol = e.range.getColumn();
  const numCols = e.range.getNumColumns();
  const touchesA = startCol <= 1 && startCol + numCols - 1 >= 1;
  const touchesC = startCol <= 3 && startCol + numCols - 1 >= 3;
  if (!touchesA && !touchesC) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // 全行のA列（作家名）・C列（作品名）をまとめて読み込む
  const allArtists = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const allProducts = sheet.getRange(2, 3, lastRow - 1, 1).getValues();

  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();
  const messages = [];

  for (let r = 0; r < numRows; r++) {
    const row = startRow + r;
    if (row <= 1) continue; // ヘッダー行は対象外

    // e.valueには頼らず実際のセル値を読み直す（複数セルへのコピペ対策）
    const artist = String(sheet.getRange(row, 1).getValue() || '').trim();
    const product = String(sheet.getRange(row, 3).getValue() || '').trim();
    if (!artist || !product) continue;

    const key = normalizeToken(artist) + '|' + normalizeToken(product);
    const duplicates = [];
    for (let i = 0; i < allArtists.length; i++) {
      const rowNum = i + 2;
      if (rowNum === row) continue; // 自分自身はスキップ
      const existingArtist = String(allArtists[i][0] || '').trim();
      const existingProduct = String(allProducts[i][0] || '').trim();
      if (!existingArtist || !existingProduct) continue;
      const existingKey = normalizeToken(existingArtist) + '|' + normalizeToken(existingProduct);
      if (existingKey === key) {
        duplicates.push(existingArtist + '／' + existingProduct + '（' + rowNum + '行目）');
      }
    }

    if (duplicates.length > 0) {
      messages.push('入力値：' + artist + '／' + product + '（' + row + '行目）\n重複している行：\n' + duplicates.join('\n'));
    }
  }

  if (messages.length > 0) {
    SpreadsheetApp.getUi().alert(
      '⚠️ 商品コード管理に重複の可能性があります\n\n' + messages.join('\n\n') +
      '\n\n同じ商品に2つの商品コードが振られてしまうと在庫管理が二重にカウントされる原因になります。\nどちらか一方の行を削除・統合してください。'
    );
  }
}

// ============================================================
// 商品コード管理シート全体を一括スキャンし、「作家名＋作品名」の組み合わせが
// 重複している行を検出してF列（商品コード）を色付けする共通関数。
// ②-1/②-2（updateProductCodeSheetCore）・③（updateInventorySheet）の
// 両方から呼ばれる（商品コード管理を更新するタイミングならどちらでも検知できるように）。
// 戻り値：重複の説明文の配列（警告メッセージ用）
// ============================================================
function detectAndHighlightProductCodeDuplicates(codeSheet) {
  if (!codeSheet) return [];
  const lastRow = codeSheet.getLastRow();
  if (lastRow < 2) return [];

  const totalRows = lastRow - 1;
  const acData = codeSheet.getRange(2, 1, totalRows, 3).getValues(); // A:C列（作家名・ラベル・作品名）

  const keyToRowNums = new Map();
  const rowInfoByKey = new Map(); // key -> {artist, product}（メッセージ表示用に1件だけ覚えておく）

  for (let i = 0; i < totalRows; i++) {
    const artist = String(acData[i][0] || '').trim();
    const product = String(acData[i][2] || '').trim();
    if (!artist || !product) continue;
    const key = normalizeToken(artist) + '|' + normalizeToken(product);
    const rowNum = i + 2;
    if (!keyToRowNums.has(key)) {
      keyToRowNums.set(key, []);
      rowInfoByKey.set(key, { artist: artist, product: product });
    }
    keyToRowNums.get(key).push(rowNum);
  }

  const duplicateRowNums = new Set();
  const duplicateList = [];
  keyToRowNums.forEach(function(rowNums, key) {
    if (rowNums.length < 2) return;
    const info = rowInfoByKey.get(key);
    rowNums.forEach(function(rn) { duplicateRowNums.add(rn); });
    duplicateList.push(info.artist + '／' + info.product + '（' + rowNums.map(function(rn) { return rn + '行目'; }).join('・') + '）');
  });

  // F列（商品コード）に色を付ける。A列（作家マスタ照合）・C列（手書き検出）とは
  // 別の列なので互いに競合しない。1行ずつではなく一括で読み書きする。
  const codeCol = codeSheet.getRange(2, 6, totalRows, 1);
  const codeColors = codeCol.getBackgrounds();
  for (let i = 0; i < totalRows; i++) {
    const rowNum = i + 2;
    if (duplicateRowNums.has(rowNum)) {
      codeColors[i][0] = '#f9cb9c'; // 薄いオレンジ（「同じ商品に複数コードが振られている」目印）
    } else if (codeColors[i][0] === '#f9cb9c') {
      codeColors[i][0] = '#ffffff'; // 解消された行は色を戻す
    }
  }
  codeCol.setBackgrounds(codeColors);

  return duplicateList;
}

// ============================================================
// 納品確認書H3（作家名プルダウン）が変更されたら納品確認書を再生成する
// 【インストール型トリガー】GASエディタのトリガー画面から手動設定が必要
// （理由：fillInvoice内でマスタファイルをopenByIdで開くため。簡易トリガーの
//   onEditからは他ファイルへアクセスできず、常に黙って失敗してしまう）
// ============================================================
function onInvoiceArtistChange(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.range.getSheet();
  const invoiceSheet = ss.getSheetByName(INVOICE_SHEET_NAME);

  if (!invoiceSheet) return;
  if (sheet.getName() !== INVOICE_SHEET_NAME) return;
  if (e.range.getA1Notation() !== 'H3') return;

  fillInvoice(e.value);
}

function fillInvoice(artistName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const finalSheet   = ss.getSheetByName(PRODUCT_SHEET_FINAL);
  const invoiceSheet = ss.getSheetByName(INVOICE_SHEET_NAME);

  if (!finalSheet || !invoiceSheet) return;

  // 確定シートから該当作家の商品を取得
  // 確定シートの列: 0:作家名 1:作品名 2:サンプル数 3:サンプル販売可否 4:納品数 5:税抜価格 6:税込価格 7:備考
  // ※完全一致ではなくnormalizeToken（全角半角・大文字小文字・空白の違いを吸収）で突き合わせる。
  //   H3のプルダウン自体は商品一覧_確定の値をそのまま使っているため通常はここでズレないはずだが、
  //   念のため他の箇所と同じ正規化方式に揃えている。
  const finalData = finalSheet.getDataRange().getValues();
  const normalizedTargetArtist = normalizeToken(artistName);
  const products = [];
  for (let i = 1; i < finalData.length; i++) {
    if (normalizeToken(String(finalData[i][0])) === normalizedTargetArtist) {
      products.push({
        name:     finalData[i][1],
        sample:   finalData[i][2],
        sampleOk: finalData[i][3],
        stock:    finalData[i][4],
        price:    finalData[i][5],
        note:     finalData[i][7],
      });
    }
  }

  if (products.length === 0) return;

  // 作家マスタから個人情報を取得（第6回からマスタファイル側の作家マスタを参照）
  // 作家マスタの列: A(0):作家番号 B(1):アーティスト名 C(2):ラベル表記名 D(3):instagram
  // E(4):ジャンル F(5):メールアドレス G(6):住所1 H(7):住所2 I(8):本名
  // J(9):電話番号 K(10):振込先情報 L(11):取引申請書送付済み M(12):取引申請書受信済み N(13):備考
  // ※こちらもnormalizeTokenで突き合わせる。ただし住所・電話番号・振込先という個人情報を
  //   人の確認なしに自動で差し込む処理のため、検索機能のような部分一致（トークンの重なり
  //   だけで拾う方式）はあえて使わない。似た名前の別の作家の個人情報を誤って引っ張って
  //   しまうリスクを避けるため、正規化後の完全一致までに留めている。
  const masterSheet = getArtistMasterSheet();
  let realName = '', address = '', tel = '', bankInfo = '';
  if (masterSheet) {
    const masterData = masterSheet.getDataRange().getValues();
    for (let i = 1; i < masterData.length; i++) {
      if (normalizeToken(String(masterData[i][1])) === normalizedTargetArtist) {
        address  = masterData[i][6] || '';  // G列: 住所1
        realName = masterData[i][8] || '';  // I列: 本名
        tel      = masterData[i][9] || '';  // J列: 電話番号
        bankInfo = masterData[i][10] || ''; // K列: 振込先情報
        break;
      }
    }
  }

  // ---- 商品欄・備考欄の位置を「名前付き範囲」から取得する ----
  // 行番号を決め打ちにせず、スプレッドシート側に「今どこにあるか」を毎回聞きに行く方式。
  // これにより、誰かが商品欄の行（17行目や27・28行目など）を手動で削除・挿入してしまっても、
  // 次にH3で作家を選び直した瞬間に自動で正しい位置・必要な行数まで復元される。
  // 【事前準備】納品確認書シートで以下の名前付き範囲を設定しておくこと（データ→名前付き範囲）：
  //   ・InvoiceProductArea：商品欄全体（例：B17:H28）
  //   ・InvoiceNoteCell　：備考を書き込むセル（例：B33）
  const BASE_ROWS = 12; // 商品欄の基本行数（これを下回っていたら自動でここまで復元する）

  let productRange = ss.getRangeByName('InvoiceProductArea');
  const noteRange = ss.getRangeByName('InvoiceNoteCell');

  if (!productRange || !noteRange) {
    SpreadsheetApp.getUi().alert(
      '⚠️ 納品確認書に必要な名前付き範囲が見つかりません。\n' +
      '「データ」→「名前付き範囲」で、商品欄に「InvoiceProductArea」、備考セルに「InvoiceNoteCell」を設定してください。\n' +
      '（削除されてしまった場合は再設定が必要です。設定方法は引き継ぎ書を参照）'
    );
    return;
  }

  const productStartRow = productRange.getRow();
  const currentRows = productRange.getNumRows();
  const desiredRows = Math.max(BASE_ROWS, products.length); // 最低12行、必要ならそれ以上

  if (currentRows < desiredRows) {
    // 行が足りない（誰かが削除した、または今回の商品数が多い）→ 不足分を復元・追加する
    const shortage = desiredRows - currentRows;
    const lastRow = productStartRow + currentRows - 1; // 現在の商品欄の最終行
    // 最終行の直前に挿入＝名前付き範囲の内側への挿入として扱われ、範囲自体も自動で広がる
    invoiceSheet.insertRowsBefore(lastRow, shortage);

    // 挿入した行に、現存する最終行（挿入によって下にずれた元の最終行）の書式・データ検証をコピーする
    const numCols = invoiceSheet.getMaxColumns();
    const templateRow = lastRow + shortage;
    for (let r = 0; r < shortage; r++) {
      invoiceSheet.getRange(templateRow, 1, 1, numCols)
        .copyTo(invoiceSheet.getRange(lastRow + r, 1, 1, numCols));
    }
  } else if (currentRows > desiredRows) {
    // 前回多くの商品数を扱った名残で行が多すぎる→基本行数まで縮めておく
    const excess = currentRows - desiredRows;
    const lastRow = productStartRow + currentRows - 1;
    invoiceSheet.deleteRows(lastRow - excess + 1, excess);
  }

  // 行の増減後、範囲を取り直す（Sheets側で自動的に追従済みのはずだが、念のため再取得する）
  productRange = ss.getRangeByName('InvoiceProductArea');
  const finalProductStartRow = productRange.getRow();
  const totalRows = productRange.getNumRows();
  const productEndRow = finalProductStartRow + totalRows - 1;
  const noteRow = ss.getRangeByName('InvoiceNoteCell').getRow();

  // ---- 合計欄の数式を毎回書き直す（自己修復）----
  // 名前付き範囲は「位置」を復元してくれるが、数式そのものが人為的に消されてしまった
  // 場合までは救えない。そのため合計欄の数式は、商品欄の直下にあるという前提のもと、
  // 実行のたびにスクリプト側で書き直す。これにより、誰かが合計欄の行（29・30行目相当）
  // を丸ごと消してしまっても、次にH3を選び直した瞬間に正しい数式へ自動修復される。
  // 元の数式（商品欄が17〜28行目・合計欄が29〜30行目だった時点のもの）：
  //   J29 = SUM(D29:E30)
  //   J30 = SUM(J17:J28)
  //   C30 = SUMPRODUCT(G17:G28, H17:H28)
  //   D30 = ROUND(C30*10%,1)
  //   C14 = J30
  const totalsRow1 = productEndRow + 1; // 元の「29行目」に相当
  const totalsRow2 = productEndRow + 2; // 元の「30行目」に相当
  const productRangeA1 = function(col) {
    return col + finalProductStartRow + ':' + col + productEndRow;
  };

  invoiceSheet.getRange('J' + totalsRow1).setFormula(
    '=SUM(D' + totalsRow1 + ':E' + totalsRow2 + ')'
  );
  invoiceSheet.getRange('J' + totalsRow2).setFormula(
    '=SUM(' + productRangeA1('J') + ')'
  );
  invoiceSheet.getRange('C' + totalsRow2).setFormula(
    '=SUMPRODUCT(' + productRangeA1('G') + ', ' + productRangeA1('H') + ')'
  );
  invoiceSheet.getRange('D' + totalsRow2).setFormula(
    '=ROUND(C' + totalsRow2 + '*10%,1)'
  );
  invoiceSheet.getRange('C14').setFormula('=J' + totalsRow2);

  // 作家情報を書き込み
  invoiceSheet.getRange('H6').setValue((realName || artistName) + ' 様'); // 本名（半角スペース＋様）
  invoiceSheet.getRange('H7').setValue(address);                // 住所
  invoiceSheet.getRange('H10').setNumberFormat('@');
  invoiceSheet.getRange('H10').setValue(tel ? 'TEL：' + String(tel) : ''); // 電話番号
  invoiceSheet.getRange('H11').setValue(bankInfo);              // 振込先情報

  // F列（サンプル販売可否）のデータ検証を解除
  invoiceSheet.getRange(finalProductStartRow, 6, totalRows, 1).clearDataValidations();

  // 商品行をクリア
  const emptyCol = Array.from({length: totalRows}, function() { return ['']; });
  invoiceSheet.getRange(finalProductStartRow, 2, totalRows, 1).setValues(emptyCol);
  invoiceSheet.getRange(finalProductStartRow, 5, totalRows, 1).setValues(emptyCol);
  invoiceSheet.getRange(finalProductStartRow, 6, totalRows, 1).setValues(emptyCol);
  invoiceSheet.getRange(finalProductStartRow, 7, totalRows, 1).setValues(emptyCol);
  invoiceSheet.getRange(finalProductStartRow, 8, totalRows, 1).setValues(emptyCol);

  // 商品データを書き込み
  const colB = [], colE = [], colF = [], colG = [], colH = [];
  for (let i = 0; i < products.length; i++) {
    const p = products[i];

    // サンプル販売可否が「可能」の場合はサンプル数を納品数に加算
    let stockNum = parseFloat(String(p.stock));
    const sampleNum = parseFloat(String(p.sample));
    const sampleOkStr = String(p.sampleOk).trim();
    if (sampleOkStr === '可能' && !isNaN(sampleNum)) {
      stockNum = (isNaN(stockNum) ? 0 : stockNum) + sampleNum;
    }
    const stockVal = isNaN(stockNum) ? (p.stock !== undefined ? p.stock : '') : stockNum;

    colB.push([p.name     !== undefined ? p.name     : '']);
    colE.push([p.sample   !== undefined ? p.sample   : '']);
    colF.push([p.sampleOk !== undefined ? p.sampleOk : '']);
    colG.push([stockVal]);
    colH.push([p.price    !== undefined ? p.price    : '']);
  }
  invoiceSheet.getRange(finalProductStartRow, 2, products.length, 1).setValues(colB);
  invoiceSheet.getRange(finalProductStartRow, 5, products.length, 1).setValues(colE);
  invoiceSheet.getRange(finalProductStartRow, 6, products.length, 1).setValues(colF);
  invoiceSheet.getRange(finalProductStartRow, 7, products.length, 1).setValues(colG);
  invoiceSheet.getRange(finalProductStartRow, 8, products.length, 1).setValues(colH);

  // 備考
  const notes = products
    .map(function(p) { return p.note ? p.name + '：' + p.note : ''; })
    .filter(function(n) { return n !== ''; })
    .join('\n');
  invoiceSheet.getRange(noteRow, 2).setWrap(true).setValue(notes);
  invoiceSheet.setRowHeight(noteRow, 21); // 一度リセット
  SpreadsheetApp.flush();
  invoiceSheet.autoResizeRows(noteRow, 1); // 内容に合わせて高さを調整
}


// ============================================================
// シートを保護する（既存の保護は一度解除してから再設定）
// ============================================================
function protectSheet(sheet, description) {
  // 既存の保護を解除
  const existing = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  existing.forEach(function(p) { p.remove(); });

  // 警告モードで再設定：編集しようとすると確認ダイアログが出るが、ブロックはしない
  sheet.protect()
    .setDescription(description)
    .setWarningOnly(true);
}


// ============================================================
// スプレッドシートを開いたときにメニューを追加
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 納品管理')
    .addItem('① 納品予定を更新', 'updatePlanList')
    .addItem('②-1 商品一覧_確定のみ更新（商品名チェック用）', 'updateFinalListOnly')
    .addItem('②-2 納品確定を更新（プルダウン・商品コードも自動更新）', 'updateFinalList')
    .addItem('③ 在庫管理(商品ごと)を更新', 'updateInventorySheet')
    .addItem('③-2 在庫管理(SKUごと)を更新', 'updateInventorySheetBySku')
    .addSeparator()
    .addItem('納品確認書を手動更新（現在の作家名で再実行）', 'manualRefreshInvoice')
    .addItem('納品確認書プルダウンのみ再設定', 'setupInvoiceDropdown')
    .addItem('会期プルダウンを更新（ロイヤリティレポート!C13・納品確認書!C13）', 'setupRoyaltyPeriodDropdown')
    .addItem('フォーム受取チェックを再実行（手動）', 'recheckFormReceipts')
    .addItem('在庫変動ログの商品名プルダウンを一斉更新', 'refreshAllLogProductDropdowns')
    .addToUi();
  // ※「🔍 作家名で検索してNo.を入力」は第6回からマスタファイル側の専用メニューに移管した
}

function manualRefreshInvoice() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invoiceSheet = ss.getSheetByName(INVOICE_SHEET_NAME);
  if (!invoiceSheet) return;
  const artistName = invoiceSheet.getRange('H3').getValue();
  if (!artistName) {
    SpreadsheetApp.getUi().alert('H3セルに作家名が選択されていません。');
    return;
  }
  fillInvoice(artistName);
}


// ============================================================
// 商品一覧シートの作家間に太めの罫線を引く
// buildProductSheet の後に呼び出す
// ============================================================
function addArtistBorders(sheet, rows, headers) {
  const colCount = headers.length;

  for (let i = 0; i < rows.length - 1; i++) {
    const currentArtist = rows[i][0];
    const nextArtist    = rows[i + 1][0];
    if (currentArtist !== nextArtist) {
      // i+2 行目の下（= i+3 行目の上）に太い罫線
      const borderRow = sheet.getRange(i + 2, 1, 1, colCount);
      borderRow.setBorder(
        null, null, true, null, null, null,
        '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
      );
    }
  }
}


// ============================================================
// 作家マスタの重複チェック（onEdit で自動実行）
// 大文字小文字・全角半角・前後空白・中間空白を無視して比較
// ============================================================
// 文字列をNFKC正規化して記号・空白を除去した文字列を返す
function normalizeToken(str) {
  if (!str) return '';
  str = str.normalize('NFKC');
  str = str.toLowerCase();
  str = str.replace(/[^a-z0-9぀-ゟ゠-ヿ一-鿿]/g, '');
  return str;
}

// 名前を「本体」と「括弧内」のトークンに分解して返す
// 例: 'キャシー(joy)' → ['キャシー', 'joy']
// 例: '𝐎𝐡𝐚𝐍𝐢(OhaNi)' → ['𝐎𝐡𝐚𝐍𝐢', 'OhaNi']
function extractTokens(str) {
  if (!str) return [];
  const tokens = [];
  // 括弧内を抽出
  const matches = str.match(/[(（]([^)）]+)[)）]/g);
  if (matches) {
    matches.forEach(function(m) {
      const inner = m.replace(/^[(（]|[)）]$/g, '');
      const t = normalizeToken(inner);
      if (t) tokens.push(t);
    });
  }
  // 括弧を除去した本体を抽出
  const body = normalizeToken(str.replace(/[(（][^)）]*[)）]/g, ''));
  if (body) tokens.push(body);
  return tokens;
}

// ※checkArtistDuplicate関数は第6回からマスタファイル側の専用スクリプトに移管した。
//   （作家マスタがこのファイルに存在しなくなり、onEditもこのファイルの編集にしか
//   反応しないため、作家マスタを直接編集するマスタファイル側に置く必要がある）


// ============================================================
// 在庫変動ログ：作家名（B列）を編集したら、同じ行の作品名（C列）に
// その作家の商品一覧_確定上の作品名リストをデータの入力規則として設定する
// ============================================================
// ============================================================
// 在庫変動ログの1行分（B列：作家名）に対して、C列（作品名）の
// プルダウンを設定する共通処理。onEdit経由の自動更新・メニューからの
// 一斉更新の両方から呼ばれる。
// ============================================================
function applyLogProductDropdown(sheet, row, finalData) {
  const artistName = String(sheet.getRange(row, 2).getValue() || '').trim();
  const productCell = sheet.getRange(row, 3); // C列：作品名

  if (!artistName || !finalData) {
    productCell.clearDataValidations();
    return;
  }

  const normalizedTarget = normalizeToken(artistName);
  const productNames = [];
  for (let i = 1; i < finalData.length; i++) {
    const rowArtist = String(finalData[i][0] || '');
    if (normalizeToken(rowArtist) === normalizedTarget) {
      const productName = finalData[i][1];
      if (productName && productNames.indexOf(productName) === -1) {
        productNames.push(productName);
      }
    }
  }

  if (productNames.length === 0) {
    // 商品一覧_確定に該当作家がいない場合（前会期の余り在庫など）は手入力を許可
    productCell.clearDataValidations();
    return;
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(productNames, true)
    .setAllowInvalid(true) // 手入力も許可（未登録作品メモ運用のため）
    .build();
  productCell.setDataValidation(rule);
}

// ============================================================
// onEdit経由：編集された範囲（B列を含む場合）だけプルダウンを更新する
// ============================================================
function updateLogProductDropdown(e) {
  const sheet = e.range.getSheet();
  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();
  const startCol = e.range.getColumn();
  const numCols = e.range.getNumColumns();

  // 編集範囲がB列（作家名）にかかっていない場合は対象外
  if (startCol > 2 || startCol + numCols - 1 < 2) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const finalSheet = ss.getSheetByName(PRODUCT_SHEET_FINAL);
  const finalData = finalSheet ? finalSheet.getDataRange().getValues() : null;

  for (let r = 0; r < numRows; r++) {
    const row = startRow + r;
    if (row <= 1) continue; // ヘッダー行は対象外
    applyLogProductDropdown(sheet, row, finalData);
  }
}

// ============================================================
// メニュー実行用：在庫変動ログの全行を対象に、C列（作品名）の
// プルダウンを一斉更新する。
// 【使いどころ】B列（作家名）が先に入力されていた行は、その時点では
// まだ商品一覧_確定が存在しない・作家名が一致しない等の理由でプルダウンが
// 未設定のまま残ることがある。そのままだと後から商品一覧_確定が更新されても、
// 既存の行のプルダウンは自動では追従しない（onEditは新たに編集された
// セルにしか反応しないため）。この関数を実行すると、シート全体を
// 再スキャンしてB列の内容をもとにC列のプルダウンをすべて設定し直す。
//
// 【パフォーマンスについて】
// 1行ずつgetValue()/setDataValidation()を呼ぶと、行数分だけサーバーとの
// 通信が発生して重くなる（1行につきB列読み取り×1・プルダウン設定×1で、
// 実質行数×2回の通信）。そのため、B列は1回のまとめ読み、プルダウンの
// 設定・解除もsetDataValidations()で範囲全体を1回にまとめて書き込む
// ことで、行数に関わらず通信回数をほぼ一定に抑えている。
// ============================================================
function refreshAllLogProductDropdowns() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOG_SHEET_NAME);
  const finalSheet = ss.getSheetByName(PRODUCT_SHEET_FINAL);

  if (!sheet) {
    ui.alert('シート「' + LOG_SHEET_NAME + '」が見つかりません。');
    return;
  }
  if (!finalSheet) {
    ui.alert('シート「' + PRODUCT_SHEET_FINAL + '」が見つかりません。\n先に「② 納品確定を更新」を実行してください。');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('在庫変動ログにデータがありません。');
    return;
  }

  const finalData = finalSheet.getDataRange().getValues();

  // 事前に「正規化した作家名 → 作品名一覧」のマップを1回だけ作っておく
  // （行ごとにfinalDataを毎回スキャンし直すのは無駄なため）
  const productsByArtist = new Map();
  for (let i = 1; i < finalData.length; i++) {
    const rowArtist = String(finalData[i][0] || '');
    const productName = finalData[i][1];
    if (!rowArtist || !productName) continue;
    const key = normalizeToken(rowArtist);
    if (!productsByArtist.has(key)) productsByArtist.set(key, []);
    const list = productsByArtist.get(key);
    if (list.indexOf(productName) === -1) list.push(productName);
  }

  // B列（作家名）をまとめて1回で読み込む
  const numRows = lastRow - 1; // ヘッダーを除いた行数
  const artistNames = sheet.getRange(2, 2, numRows, 1).getValues();

  // C列に設定するデータ検証ルールを、行ごとに配列として組み立てる
  // （nullを入れるとその行のプルダウンは解除される）
  const validations = [];
  let updatedCount = 0;

  for (let r = 0; r < numRows; r++) {
    const artistName = String(artistNames[r][0] || '').trim();
    if (!artistName) {
      validations.push([null]);
      continue;
    }

    const productNames = productsByArtist.get(normalizeToken(artistName)) || [];
    if (productNames.length === 0) {
      // 商品一覧_確定に該当作家がいない場合（前会期の余り在庫など）は手入力を許可
      validations.push([null]);
    } else {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(productNames, true)
        .setAllowInvalid(true) // 手入力も許可（未登録作品メモ運用のため）
        .build();
      validations.push([rule]);
    }
    updatedCount++;
  }

  // C列（作品名）へ一括書き込み
  sheet.getRange(2, 3, numRows, 1).setDataValidations(validations);

  ui.alert('完了！\n在庫変動ログの' + updatedCount + '行分、作品名のプルダウンを更新しました。');
}


// ============================================================
// 商品コード管理シートを更新する
// 商品一覧_確定にある「作家名＋作品名」のうち、商品コード管理に
// まだない組み合わせは新規追加する。
// 既存行の税込価格は商品一覧_確定の最新値で上書きする（商品コードは上書きしない）。
// ============================================================
function updateProductCodeSheetCore(finalSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let codeSheet = ss.getSheetByName(PRODUCT_CODE_SHEET);

  if (!codeSheet) {
    codeSheet = ss.insertSheet(PRODUCT_CODE_SHEET);
    codeSheet.getRange(1, 1, 1, 7).setValues([
      ['作家名', 'ラベル表示作家名', '作品名', '税込価格', 'ラベル表示商品名', '商品コード', '在庫残数']
    ]);
    codeSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }

  // 作家マスタから「ラベル表記名」を引くマップを作成（作家名キー、正規化して比較）
  // 第6回からマスタファイル側の作家マスタを参照
  const labelMap = new Map(); // normalizeToken(作家名) -> ラベル表記名
  const masterSheet = getArtistMasterSheet();
  if (masterSheet) {
    const masterData = masterSheet.getDataRange().getValues();
    for (let i = 1; i < masterData.length; i++) {
      const artist = String(masterData[i][1] || '').trim(); // B列：アーティスト名
      if (!artist) continue;
      labelMap.set(normalizeToken(artist), masterData[i][2] || ''); // C列：ラベル表記名
    }
  }

  // 既存行の組み合わせ→行番号のマップを作成（正規化して比較）
  // 列：A作家名 B ラベル表示作家名 C作品名 D税込価格 E ラベル表示商品名 F商品コード G在庫残数
  const existingData = codeSheet.getDataRange().getValues();
  const existingRowMap = new Map(); // key -> 行番号（1始まり、シート上の実際の行）
  for (let i = 1; i < existingData.length; i++) {
    const artist = String(existingData[i][0] || '');
    const product = String(existingData[i][2] || '');
    if (!artist && !product) continue;
    const key = normalizeToken(artist) + '|' + normalizeToken(product);
    existingRowMap.set(key, i + 1); // シート上の行番号
  }

  // D列（税込価格）に入れる数式を組み立てる
  // 商品一覧_確定のA列（作家名）・B列（作品名）を「作家名|作品名」で連結したキーで突き合わせ、
  // G列（税込価格）を取得する。商品一覧_確定が更新されるたびに自動で反映される
  // （スクリプトの再実行やボタン操作が不要になる）。
  // ※スクリプト側で行っているnormalizeTokenのようなゆるいマッチングは数式では再現できず、
  //   完全一致での突き合わせになる点に注意（表記ゆれがあると空欄になる）
  // ※XLOOKUPで範囲同士を&連結する場合、ARRAYFORMULAで明示的に包まないと
  //   配列として計算されず#N/Aになる（列全体参照(A:A)だと稀に計算が噛み合わないこともあるため、
  //   実際にデータがある行数までに絞った範囲を使う）
  const finalLastRow = Math.max(2, finalSheet.getLastRow());
  function buildTaxPriceFormula(rowNum) {
    return '=IFERROR(ARRAYFORMULA(XLOOKUP($A' + rowNum + '&"|"&$C' + rowNum + ', ' +
      PRODUCT_SHEET_FINAL + '!$A$2:$A$' + finalLastRow + '&"|"&' + PRODUCT_SHEET_FINAL + '!$B$2:$B$' + finalLastRow + ', ' +
      PRODUCT_SHEET_FINAL + '!$G$2:$G$' + finalLastRow + ')), "")';
  }

  // 商品一覧_確定から作家名・作品名を取得（税込価格は数式に任せるのでここでは読まない）
  const finalData = finalSheet.getDataRange().getValues();
  const newRows = [];
  const seenInThisRun = new Set(); // 同一実行内での重複追加防止
  let updatedCount = 0;

  // 既存行のB列（ラベル表示作家名）・D列（税込価格）は、1行ずつgetRange().setValue()すると
  // 行数分だけサーバー通信が発生して重くなるため、まずexistingData（読み込み済み）をもとに
  // 列全体を配列として組み立て、最後に1回でまとめて書き込む方式にしている。
  const totalExistingRows = existingData.length - 1; // ヘッダーを除いた行数
  const bColumnValues = []; // B列：ラベル表示作家名
  const dColumnValues = []; // D列：税込価格（数式）
  for (let i = 1; i < existingData.length; i++) {
    bColumnValues.push([existingData[i][1]]); // 現状維持がデフォルト（対象外の行はそのまま）
    dColumnValues.push([existingData[i][3]]); // 同上
  }

  for (let i = 1; i < finalData.length; i++) {
    const artist = String(finalData[i][0] || '').trim();
    const product = String(finalData[i][1] || '').trim();
    if (!artist || !product) continue;

    const key = normalizeToken(artist) + '|' + normalizeToken(product);
    const labelArtist = labelMap.has(normalizeToken(artist)) ? labelMap.get(normalizeToken(artist)) : '';

    if (existingRowMap.has(key)) {
      // 既存行：D列（税込価格）を数式に、ラベル表示作家名だけ上書き（作家マスタの最新値に追従させる）
      const rowNum = existingRowMap.get(key);
      const arrIdx = rowNum - 2; // シート上の行番号→配列インデックス（2行目が配列の0番目）
      bColumnValues[arrIdx] = [labelArtist];
      dColumnValues[arrIdx] = [buildTaxPriceFormula(rowNum)];
      updatedCount++;
      continue;
    }
    if (seenInThisRun.has(key)) continue;

    seenInThisRun.add(key);
    // [作家名, ラベル表示作家名, 作品名, 税込価格(後で数式を入れる), ラベル表示商品名, 商品コード, 在庫残数]
    // ラベル表示商品名・商品コード・在庫残数は空欄で追加、後で手入力
    newRows.push([artist, labelArtist, product, '', '', '', '']);
  }

  // B列・D列（既存分）を1回でまとめて書き込む
  if (totalExistingRows > 0) {
    codeSheet.getRange(2, 2, totalExistingRows, 1).setValues(bColumnValues);
    codeSheet.getRange(2, 4, totalExistingRows, 1).setValues(dColumnValues);
  }

  if (newRows.length > 0) {
    const startRow = codeSheet.getLastRow() + 1;
    codeSheet.getRange(startRow, 1, newRows.length, 7).setValues(newRows);
    // 新規追加した行のD列（税込価格）にも、1回でまとめて数式を設定する
    const newDColumnValues = [];
    for (let r = 0; r < newRows.length; r++) {
      newDColumnValues.push([buildTaxPriceFormula(startRow + r)]);
    }
    codeSheet.getRange(startRow, 4, newRows.length, 1).setValues(newDColumnValues);
  }

  // ---- A列（作家名）が作家マスタに存在するかチェックし、無ければ赤く色付け ----
  // labelMapのキーは既に「正規化した作家マスタの作家名」の集合なので、そのまま流用できる。
  // 1行ずつgetRange().setBackground()すると重くなるため、A列の背景色をまとめて読み込み、
  // 対象行だけ差し替えてから1回でまとめて書き戻す（他の一括処理と同じ方式）。
  // ※B列（ラベル表示作家名の上書き）とは別列なので、互いの色付けが競合しない。
  const finalLastCodeRow = codeSheet.getLastRow();
  const missingArtistsInCode = [];
  if (finalLastCodeRow >= 2) {
    const totalRows = finalLastCodeRow - 1;
    const artistCol = codeSheet.getRange(2, 1, totalRows, 1);
    const artistValues = artistCol.getValues();
    const artistColors = artistCol.getBackgrounds();

    for (let i = 0; i < totalRows; i++) {
      const artist = String(artistValues[i][0] || '').trim();
      if (!artist) continue;
      if (!labelMap.has(normalizeToken(artist))) {
        missingArtistsInCode.push(artist + '（' + (i + 2) + '行目）');
        artistColors[i][0] = '#f4cccc'; // 薄い赤
      } else if (artistColors[i][0] === '#f4cccc') {
        // 以前は未登録だったが今は解消された行は、色を元に戻す
        artistColors[i][0] = '#ffffff';
      }
    }
    artistCol.setBackgrounds(artistColors);
  }

  const duplicateList = detectAndHighlightProductCodeDuplicates(codeSheet);

  return { newRowCount: newRows.length, missingArtistsInCode: missingArtistsInCode, duplicateList: duplicateList };
}


// ============================================================
// 在庫変動ログ：B〜G列のいずれかに入力があったとき、
// 同じ行のA列（確認日）が空欄なら今日の日付を自動入力する
// 複数行・複数セルの一括貼り付けにも対応
// ============================================================
function fillLogCheckDate(e) {
  const sheet = e.range.getSheet();
  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();
  const startCol = e.range.getColumn();
  const numCols = e.range.getNumColumns();

  // ヘッダー行は対象外
  // 編集範囲がB〜G列（2〜7：作家名〜担当者）にかかっていない場合は対象外
  if (startCol + numCols - 1 < 2 || startCol > 7) return;

  const today = new Date();

  for (let r = 0; r < numRows; r++) {
    const rowNum = startRow + r;
    if (rowNum <= 1) continue;

    // この行のB〜G列に何か値があるか確認
    const rowValues = sheet.getRange(rowNum, 2, 1, 6).getValues()[0];
    const hasAnyValue = rowValues.some(function(v) { return v !== '' && v !== null; });
    if (!hasAnyValue) continue;

    const dateCell = sheet.getRange(rowNum, 1);
    if (dateCell.getValue() === '' || dateCell.getValue() === null) {
      dateCell.setValue(today);
    }
  }
}


// ============================================================
// 在庫管理シートを更新する
// 商品一覧_確定 ＋ 在庫変動ログ ＋ 商品コード管理 から組み立てる
// 行のキー：正規化した「作家名｜作品名」
// ============================================================
function updateInventorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const finalSheet = ss.getSheetByName(PRODUCT_SHEET_FINAL);
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  const codeSheet = ss.getSheetByName(PRODUCT_CODE_SHEET);

  if (!finalSheet) {
    SpreadsheetApp.getUi().alert('シート「' + PRODUCT_SHEET_FINAL + '」が見つかりません。\n先に「② 納品確定を更新」を実行してください。');
    return;
  }

  // ---- 商品コード管理から「商品コード・在庫残数（手入力の現在庫）」を引くマップを作成 ----
  // 商品コード管理の列：A作家名 B ラベル表示作家名 C作品名 D税込価格 E ラベル表示商品名 F商品コード G在庫残数
  // ※税込価格は商品コード管理経由ではなく商品一覧_確定から直接取得する（下記参照）。
  //   理由：商品コード管理D列は「②-2 納品確定を更新」実行時にしか同期されないため、
  //   「②-1 商品一覧_確定のみ更新」だけを実行した場合に価格が古いまま残ってしまうリスクがあった。
  //   商品一覧_確定を唯一の価格ソースにすることで、更新のたびに二重に手間をかける必要をなくしている。
  const codeMap = new Map(); // key -> {code, stock}
  const codeAllRows = []; // {key, artist, product, rowNum}（商品コード管理の全行。後で「在庫管理に存在しない行」の検出に使う）
  if (codeSheet) {
    const codeData = codeSheet.getDataRange().getValues();
    for (let i = 1; i < codeData.length; i++) {
      const artist = String(codeData[i][0] || '').trim();
      const product = String(codeData[i][2] || '').trim();
      if (!artist || !product) continue;
      const key = normalizeToken(artist) + '|' + normalizeToken(product);
      codeMap.set(key, {
        code: codeData[i][5],
        stock: codeData[i][6],
      });
      codeAllRows.push({ key: key, artist: artist, product: product, rowNum: i + 1 });
    }
  }

  // ---- 商品一覧_確定から基本行を作成 ----
  // key -> { artist, product, sampleQtyPlanned, deliveryQtyPlanned, sampleOk, taxPrice }
  const baseMap = new Map();
  const baseOrder = []; // 表示順を保持

  const finalData = finalSheet.getDataRange().getValues();
  for (let i = 1; i < finalData.length; i++) {
    const artist = String(finalData[i][0] || '').trim();
    const product = String(finalData[i][1] || '').trim();
    if (!artist || !product) continue;

    const sampleQty = finalData[i][2]; // サンプル数（予定）
    const sampleOk = String(finalData[i][3] || '').trim(); // サンプル販売可否
    const deliveryQty = finalData[i][4]; // 納品数（予定）
    const taxPrice = finalData[i][6]; // 税込価格

    const key = normalizeToken(artist) + '|' + normalizeToken(product);
    if (!baseMap.has(key)) {
      baseMap.set(key, {
        artist: artist,
        product: product,
        sampleQtyPlanned: sampleQty,
        deliveryQtyPlanned: deliveryQty,
        sampleOk: sampleOk,
        taxPrice: taxPrice,
      });
      baseOrder.push(key);
    }
  }

  // ---- 在庫変動ログを集計 ----
  // key -> { received, sample, reserved, reservedRelease, damaged, returned, acquired, lost, hasMemo }
  // ※「販売」種別は廃止（販売数はSquareデータ・在庫数との差分から算出するため、ここでは集計しない）
  const logMap = new Map();

  function getLogEntry(key) {
    if (!logMap.has(key)) {
      logMap.set(key, {
        artist: '', product: '',
        received: 0, sample: 0, reserved: 0, reservedRelease: 0,
        damaged: 0, returned: 0, acquired: 0, lost: 0,
        carriedIn: 0, transferredOut: 0, sampleReturned: 0,
        hasMemo: false,
      });
    }
    return logMap.get(key);
  }

  // 在庫変動ログの「種別」で認識される値の一覧
  // ※ここに無い文字列は switch 文のどのcaseにも一致せず、静かに集計から漏れる
  //   （実際に「買取在庫」と「買取済在庫」の表記違いでこの問題が起きたことがある）
  const VALID_LOG_TYPES = ['納品', 'サンプル', '取り置き', '取り置き解消', '破損', '返却', '買取済在庫', '紛失', '繰越入庫', '振替出庫', 'サンプル返却'];
  const unrecognizedTypeRows = []; // 種別が空欄、または一覧にない値になっている行

  if (logSheet) {
    const logData = logSheet.getDataRange().getValues();
    // ログの列：A確認日 B作家名 C作品名 D作品仮名 E数量 F種別 G担当者（G列は集計には未使用）
    for (let i = 1; i < logData.length; i++) {
      const artist = String(logData[i][1] || '').trim();
      let product = String(logData[i][2] || '').trim();
      const memo = String(logData[i][3] || '').trim();
      const qty = parseFloat(logData[i][4]) || 0;
      const type = String(logData[i][5] || '').trim();

      if (!artist) continue;
      if (!product && memo) product = memo; // C列が空ならD列を使う
      if (!product) continue;

      if (!type) {
        unrecognizedTypeRows.push((i + 1) + '行目：' + artist + '／' + product + '（種別が空欄）');
      } else if (VALID_LOG_TYPES.indexOf(type) === -1) {
        unrecognizedTypeRows.push((i + 1) + '行目：' + artist + '／' + product + '（種別「' + type + '」は集計対象外です）');
      }

      const key = normalizeToken(artist) + '|' + normalizeToken(product);
      const entry = getLogEntry(key);
      entry.artist = artist;
      entry.product = product;
      if (!String(logData[i][2] || '').trim() && memo) entry.hasMemo = true;

      switch (type) {
        case '納品':       entry.received += qty; break;
        case 'サンプル':    entry.sample += qty; break;
        case '取り置き':    entry.reserved += qty; break;
        case '取り置き解消': entry.reservedRelease += qty; break;
        case '破損':       entry.damaged += qty; break;
        case '返却':       entry.returned += qty; break;
        case '買取済在庫':  entry.acquired += qty; break; // 現在庫には含めるが、売上集計には含めない（ロイヤリティレポート側で対応）
        case '紛失':       entry.lost += qty; break; // 紛失分は在庫数から減算する
        case '繰越入庫':    entry.carriedIn += qty; break; // 他会期から入ってきた分。在庫数（販売可サンプル含）には加算するが、納品済み(サンプル込)には含めない
        case '振替出庫':  entry.transferredOut += qty; break; // 他会期へ出ていった分。在庫数（販売可サンプル含）から減算するが、納品済み(サンプル込)には含めない
        case 'サンプル返却': entry.sampleReturned += qty; break; // 作家へサンプルを返却した分。在庫数（販売可サンプル含）から減算する
      }

      // この行が基本行（商品一覧_確定）にない場合は新規追加対象にする
      if (!baseMap.has(key)) {
        baseMap.set(key, {
          artist: artist,
          product: product,
          sampleQtyPlanned: '', // 予定なし
          deliveryQtyPlanned: '', // 予定なし
          sampleOk: '', // 情報なし（除外しない扱い）
          taxPrice: '', // 商品一覧_確定に無い商品（買取済在庫・繰越入庫のみ等）は価格不明のため空欄
        });
        baseOrder.push(key);
      }
    }
  }

  // ---- 商品コード管理にはあるが、商品一覧_確定にもログにも存在しない行を拾い上げる ----
  // 【想定シナリオ】スタッフが商品コード管理に手書きで新しい商品を追加してしまった場合など。
  // このまま放置すると、その商品は在庫管理シートに一切出てこないため気づかれにくい。
  // 該当行は在庫管理側にも反映しつつ、商品コード管理のC列（作品名）に色を付けて目印にする。
  //
  // ここまでの時点（baseMapは商品一覧_確定＋在庫変動ログの内容のみ）で、
  // 「正規化した作家名だけ」の集合を作っておく。作家名だけなら一致する記録がある場合、
  // 「作品名の表記が微妙に違うだけ（表記ゆれ）」の可能性が高いと判定できる。
  // 逆に作家名すら一致しない場合は、本当に真新しい追加である可能性が高い。
  const knownArtistNames = new Set();
  baseMap.forEach(function(v) { knownArtistNames.add(normalizeToken(v.artist)); });

  const codeOnlyRowNums = []; // 色付け対象の行番号（商品コード管理シート上）
  const codeOnlyList = []; // 警告メッセージ用
  codeAllRows.forEach(function(row) {
    if (baseMap.has(row.key)) return; // 商品一覧_確定またはログに既に存在する
    baseMap.set(row.key, {
      artist: row.artist,
      product: row.product,
      sampleQtyPlanned: '',
      deliveryQtyPlanned: '',
      sampleOk: '',
      taxPrice: '', // 商品一覧_確定に無いため価格不明
    });
    baseOrder.push(row.key);
    codeOnlyRowNums.push(row.rowNum);

    const artistKnown = knownArtistNames.has(normalizeToken(row.artist));
    const reasonHint = artistKnown
      ? '→作家名は一致する記録あり。作品名の表記ゆれの可能性が高い'
      : '→作家名も一致する記録なし。手書きで新規追加された可能性が高い';
    codeOnlyList.push(row.artist + '／' + row.product + '（商品コード管理' + row.rowNum + '行目）' + reasonHint);
  });

  // 商品コード管理のC列（作品名）に色を付ける。既存の色付け（A列の作家マスタ照合）とは
  // 別の列なので互いに競合しない。1行ずつではなく一括で読み書きする。
  if (codeSheet) {
    const codeLastRow = codeSheet.getLastRow();
    if (codeLastRow >= 2) {
      const totalCodeRows = codeLastRow - 1;
      const productCol = codeSheet.getRange(2, 3, totalCodeRows, 1);
      const productColors = productCol.getBackgrounds();
      const codeOnlySet = new Set(codeOnlyRowNums);

      for (let i = 0; i < totalCodeRows; i++) {
        const rowNum = i + 2;
        if (codeOnlySet.has(rowNum)) {
          productColors[i][0] = '#fff2cc'; // 薄い黄色（「まだ商品一覧_確定に無い」目印）
        } else if (productColors[i][0] === '#fff2cc') {
          // 以前は該当していたが、今は商品一覧_確定に反映されて解消された行は色を戻す
          productColors[i][0] = '#ffffff';
        }
      }
      productCol.setBackgrounds(productColors);
    }
  }

  // ---- 商品コード管理内で「作家名＋作品名」の組み合わせが重複している行を検知する ----
  // 【想定シナリオ】手書きで商品コード管理に行を追加した際、既に同じ商品の行が
  // あることに気づかず、もう1行作ってしまい、同じ商品に2つの商品コードが
  // 振られてしまうケース。この場合キー自体は正規シートと一致するため、
  // 上の「codeOnly」チェック（商品一覧_確定・ログとの不一致）では検知できない。
  // ②-1/②-2実行時にも同じチェックが走るよう、共通関数化してある。
  const duplicateList = detectAndHighlightProductCodeDuplicates(codeSheet);

  // ---- 出力用の行を組み立て ----
  const headers = ['作家名', '作品名', '商品コード', '税込価格', 'サンプル数', '納品予定数',
                    '納品済み(サンプル込)', '取置き中', '紛失', '破損', '返却', '返却済サンプル', '在庫数（販売可サンプル含）', '推定販売数', '残在庫(実数)'];
  const rows = [];

  baseOrder.forEach(function(key) {
    const base = baseMap.get(key);
    const log = logMap.get(key) || {
      received: 0, sample: 0, reserved: 0, reservedRelease: 0,
      damaged: 0, returned: 0, acquired: 0, lost: 0,
      carriedIn: 0, transferredOut: 0, sampleReturned: 0,
    };
    const codeInfo = codeMap.get(key) || { code: '', stock: '' };

    // 納品予定数＝商品一覧_確定の納品数＋サンプル数（予定が空欄の場合は空欄のまま）
    let deliveryPlanned = '';
    if (base.deliveryQtyPlanned !== '' && base.deliveryQtyPlanned !== null) {
      const d = Number(base.deliveryQtyPlanned) || 0;
      const s = Number(base.sampleQtyPlanned) || 0;
      deliveryPlanned = d + s;
    }

    const delivered = log.received + log.sample; // 納品済み(サンプル込)＝納品＋サンプルの合計（実績）
    const reservedNet = log.reserved - log.reservedRelease; // 取り置き中

    // サンプル販売不可分は在庫数から除外する
    const sampleExcluded = (base.sampleOk === '不可能') ? log.sample : 0;

    // 在庫数（販売可サンプル含）＝納品済み＋買取済在庫＋繰越入庫－振替出庫－破損－返却－取り置き中－紛失
    //   －返却済サンプル－サンプル販売不可分
    // （販売は引かない＝会期末の実数カウントと比較するための理論値。繰越入庫・振替出庫・買取済在庫は
    //   「今回作家から新しく届いた点数」ではないため、納品済み(サンプル込)には含めない）
    const theoreticalStock = delivered + log.acquired + log.carriedIn - log.transferredOut
      - log.damaged - log.returned - reservedNet - log.lost - log.sampleReturned - sampleExcluded;

    // 現在庫＝商品コード管理「在庫残数」の手入力値をそのまま反映（読み取り専用のミラー）
    const manualStock = codeInfo.stock;
    const hasManualStock = manualStock !== '' && manualStock !== null && manualStock !== undefined;

    rows.push([
      base.artist,
      base.product,
      codeInfo.code,
      base.taxPrice,
      base.sampleQtyPlanned, // サンプル数：商品一覧_確定の予定値そのまま
      deliveryPlanned,
      delivered,
      reservedNet === 0 ? '' : reservedNet,
      log.lost === 0 ? '' : log.lost,
      log.damaged === 0 ? '' : log.damaged,
      log.returned === 0 ? '' : log.returned,
      log.sampleReturned === 0 ? '' : log.sampleReturned,
      theoreticalStock === 0 ? '' : theoreticalStock,
      '', // 推定販売数：後段で「在庫数（販売可サンプル含）－残在庫(実数)」の数式を差し込む（同じ行の隣接セル参照）
      hasManualStock ? manualStock : '',
    ]);
  });

  // 作家名（0列目）で昇順ソート。買取済在庫・繰越入庫のみで商品一覧_確定に
  // 存在しない商品は末尾に追加されるため、明示的にソートしないと同じ作家の行が
  // 分断され、太線が正しく引けなくなることがある。
  rows.sort(function(a, b) {
    return String(a[0]).localeCompare(String(b[0]), 'ja');
  });

  // 推定販売数（列N＝13番目、0始まりで13）に「在庫数（販売可サンプル含）－残在庫(実数)」の数式を差し込む。
  // 列M＝在庫数（販売可サンプル含）、列O＝残在庫(実数)は同じ行の隣接セルなので、行番号さえ分かれば
  // シンプルな同一シート内参照で済む（ソート後の最終的な行位置を使って組み立てる）。
  // 残在庫(実数)が空欄（会期末カウント未実施）の間は空欄のまま、という以前の挙動を
  // IF文で再現している。値ではなく数式にしたことで、残在庫(実数)が更新されると
  // スクリプトを再実行しなくても自動で再計算されるようになる。
  rows.forEach(function(row, idx) {
    const sheetRow = idx + 2; // ヘッダーが1行目なので、データはidx+2行目に書き込まれる
    row[13] = '=IF(O' + sheetRow + '="","",M' + sheetRow + '-O' + sheetRow + ')';
  });

  // ---- シートに書き込み ----
  const sheet = buildProductSheet(INVENTORY_SHEET, rows, headers);
  addArtistBorders(sheet, rows, headers);

  // 「破損」「返却」「返却済サンプル」列は普段見る必要が薄いため、デフォルトで折りたたんでおく
  // （列番号10〜12＝J・K・L列）。見たいときは列見出しの「+」をクリックすれば展開できる
  sheet.hideColumns(10, 3);

  let message = '完了！\n' + rows.length + '件の在庫データを更新しました。\n' +
    '※「残在庫(実数)」は商品コード管理の「在庫残数」を反映したものです。会期末のカウント結果はそちらに入力してください。';

  if (unrecognizedTypeRows.length > 0) {
    message += '\n\n⚠️ 在庫変動ログに種別が空欄・または認識されない行が' + unrecognizedTypeRows.length + '件あります（この行は集計に反映されていません）：\n' +
      unrecognizedTypeRows.slice(0, 15).join('\n') +
      (unrecognizedTypeRows.length > 15 ? '\n…他' + (unrecognizedTypeRows.length - 15) + '件' : '');
  }

  if (codeOnlyList.length > 0) {
    message += '\n\n⚠️ 商品コード管理にあるのに、商品一覧_確定にも在庫変動ログにも一致する記載が見つからない商品が' + codeOnlyList.length + '件あります（商品コード管理のC列を黄色くしています。この在庫管理シートには反映済みです）：\n' +
      codeOnlyList.slice(0, 15).join('\n') +
      (codeOnlyList.length > 15 ? '\n…他' + (codeOnlyList.length - 15) + '件' : '') +
      '\n\n「作品名の表記ゆれの可能性が高い」の場合→在庫変動ログ・商品一覧_確定と商品コード管理で、該当作家の作品名の表記を見比べて揃えてください\n' +
      '「手書きで新規追加された可能性が高い」の場合→在庫変動ログに「買取済在庫」等で記録がなければ追加してください（記録済みなら作家名の表記ゆれを疑ってください）';
  }

  if (duplicateList.length > 0) {
    message += '\n\n⚠️ 商品コード管理内で同じ商品（作家名＋作品名）に複数の行・複数の商品コードが振られています（商品コード管理のF列を薄いオレンジにしています）：\n' +
      duplicateList.slice(0, 15).join('\n') +
      (duplicateList.length > 15 ? '\n…他' + (duplicateList.length - 15) + '件' : '') +
      '\n\nどちらか一方の行を削除・統合してください。放置すると在庫数が二重にカウントされる可能性があります。';
  }

  SpreadsheetApp.getUi().alert(message);
}


// ============================================================
// Square売上シートを読み込み、「作家名｜作品名」キーで実質販売数（販売数－払い戻し数）を集計する
// Squareの列：A商品名(=作家名) B商品バリエーション(=作品名) C商品番号(SKU) D カテゴリ
// E販売商品数 F総売上高 G払戻済商品 H払い戻し I割引 J純売上高 K税金 L単位 M販売数 N払い戻し数
// ※SKUが空欄の行が多いため、突き合わせは作家名＋作品名（正規化）で行う
// ============================================================
function getSquareSalesMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const map = new Map(); // normalizeToken(作家名)+'|'+normalizeToken(作品名) -> 実質販売数
  const sheet = ss.getSheetByName(SQUARE_SHEET_NAME);
  if (!sheet) return map;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const artist = String(data[i][0] || '').trim();   // A列：商品名＝作家名
    const product = String(data[i][1] || '').trim();  // B列：商品バリエーション＝作品名
    if (!artist || !product) continue;

    const sold = parseFloat(data[i][12]) || 0;     // M列：販売数
    const refunded = parseFloat(data[i][13]) || 0; // N列：払い戻し数
    const net = sold - refunded;

    const key = normalizeToken(artist) + '|' + normalizeToken(product);
    map.set(key, (map.get(key) || 0) + net);
  }
  return map;
}


// ============================================================
// 在庫管理(SKUごと)シートを更新する
// 「在庫管理(商品ごと)」の行を商品コード単位でグルーピングして合算する。
// ・作家名・作品名・税込価格：同じ商品コードの最初の行の値を使う
// ・サンプル数〜現在庫：同じ商品コードの行を合計する
// ・推定販売数：商品ごとシートの「在庫数－現在庫」の合計（会期末カウントベース）
// ・Square販売数：Square売上シートを作家名＋作品名で突き合わせて合計（POSベース）
// ・両者が一致しなければ、その2セルの背景色を変える
// ・商品コードが空の行は集計対象外とし、そのまま1行ずつ出力する
// ============================================================
function updateInventorySheetBySku() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(INVENTORY_SHEET);

  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('シート「' + INVENTORY_SHEET + '」が見つかりません。\n先に「③ 在庫管理(商品ごと)を更新」を実行してください。');
    return;
  }

  const data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('在庫管理(商品ごと)にデータがありません。');
    return;
  }

  const squareMap = getSquareSalesMap();

  const headers = ['作家名', '作品名', '商品コード', '税込価格', 'サンプル数', '納品予定数',
                    '納品済み(サンプル込)', '取置き中', '紛失', '破損', '返却', '返却済サンプル', '在庫数（販売可サンプル含）', '残在庫(実数)', '推定販売数', 'Square販売数'];

  function toNum(v) {
    const n = Number(v);
    return (v === '' || v === null || v === undefined || isNaN(n)) ? 0 : n;
  }
  function hasValue(v) {
    return v !== '' && v !== null && v !== undefined;
  }

  // 在庫管理(商品ごと)の列：0作家名 1作品名 2商品コード 3税込価格 4サンプル数 5納品予定数
  // 6納品済み(サンプル込) 7取置き中 8紛失 9破損 10返却 11返却済サンプル
  // 12在庫数（販売可サンプル含） 13推定販売数 14残在庫(実数)
  const groupMap = new Map(); // 商品コード -> 集計用オブジェクト
  const groupOrder = [];
  const passthroughRows = []; // 商品コードが空の行はそのまま出力（末尾に追加）

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const code = String(row[2] || '').trim();
    const artist = row[0], product = row[1];
    const squareKey = normalizeToken(String(artist)) + '|' + normalizeToken(String(product));
    const squareNet = squareMap.has(squareKey) ? squareMap.get(squareKey) : '';

    if (!code) {
      passthroughRows.push([
        artist, product, row[2], row[3],
        row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[14],
        row[13], // 推定販売数（商品ごとシートの値そのまま、数式の計算結果が入っている）
        squareNet,
      ]);
      continue;
    }

    if (!groupMap.has(code)) {
      groupMap.set(code, {
        artist: artist, product: product, taxPrice: row[3],
        sample: 0, deliveryPlanned: 0, delivered: 0, reserved: 0, lost: 0, damaged: 0, returned: 0, sampleReturned: 0,
        stock: 0, currentStock: 0, hasCurrentStock: false,
        estimatedSold: 0, hasEstimatedSold: false,
        squareSold: 0, hasSquareSold: false,
      });
      groupOrder.push(code);
    }

    const g = groupMap.get(code);
    g.sample += toNum(row[4]);
    g.deliveryPlanned += toNum(row[5]);
    g.delivered += toNum(row[6]);
    g.reserved += toNum(row[7]);
    g.lost += toNum(row[8]);
    g.damaged += toNum(row[9]);
    g.returned += toNum(row[10]);
    g.sampleReturned += toNum(row[11]);
    g.stock += toNum(row[12]);
    if (hasValue(row[14])) { g.currentStock += toNum(row[14]); g.hasCurrentStock = true; }
    if (hasValue(row[13]))  { g.estimatedSold += toNum(row[13]); g.hasEstimatedSold = true; } // 推定販売数（数式の計算結果）
    if (squareNet !== '')  { g.squareSold += toNum(squareNet); g.hasSquareSold = true; }
  }

  // rowEntries：各行データと「不一致フラグ」をセットで持つ。ソート後もフラグが
  // 正しい行についてくるよう、rowsとmismatchRowIndexesを別々に作らずここでまとめる。
  const rowEntries = [];

  groupOrder.forEach(function(code) {
    const g = groupMap.get(code);
    const estimatedSoldVal = g.hasEstimatedSold ? g.estimatedSold : '';
    const squareSoldVal = g.hasSquareSold ? g.squareSold : '';

    const row = [
      g.artist, g.product, code, g.taxPrice,
      g.sample, g.deliveryPlanned, g.delivered, g.reserved === 0 ? '' : g.reserved,
      g.lost === 0 ? '' : g.lost,
      g.damaged === 0 ? '' : g.damaged,
      g.returned === 0 ? '' : g.returned,
      g.sampleReturned === 0 ? '' : g.sampleReturned,
      g.stock === 0 ? '' : g.stock,
      g.hasCurrentStock ? g.currentStock : '',
      estimatedSoldVal,
      squareSoldVal,
    ];
    const mismatch = g.hasEstimatedSold && g.hasSquareSold && g.estimatedSold !== g.squareSold;
    rowEntries.push({ row: row, mismatch: mismatch });
  });

  passthroughRows.forEach(function(r) {
    const est = r[14], sq = r[15];
    const mismatch = est !== '' && sq !== '' && Number(est) !== Number(sq);
    rowEntries.push({ row: r, mismatch: mismatch });
  });

  // 作家名（0列目）で昇順ソート。同じ作家の商品が分断されると太線が正しく引けなくなるため。
  // Array.sortは安定ソートなので、同じ作家内の並び順（グループ化された順）は保たれる。
  rowEntries.sort(function(a, b) {
    return String(a.row[0]).localeCompare(String(b.row[0]), 'ja');
  });

  const rows = rowEntries.map(function(entry) { return entry.row; });
  const mismatchRowIndexes = []; // rows配列上のインデックス（0始まり）
  rowEntries.forEach(function(entry, idx) {
    if (entry.mismatch) mismatchRowIndexes.push(idx);
  });

  const sheet = buildProductSheet(INVENTORY_SHEET_BY_SKU, rows, headers);
  addArtistBorders(sheet, rows, headers);

  // 「破損」「返却」「返却済サンプル」列は普段見る必要が薄いため、デフォルトで折りたたんでおく
  // （列番号10〜12＝J・K・L列）
  sheet.hideColumns(10, 3);

  // 推定販売数(15列目)とSquare販売数(16列目)が不一致の行は背景色を変える
  // 1行ずつsetBackground()すると不一致件数分サーバー通信が発生するため、
  // 該当2列分の背景色をまとめて読み込み、不一致行だけ差し替えてから1回で書き戻す。
  if (rows.length > 0) {
    const mismatchCols = sheet.getRange(2, 15, rows.length, 2);
    const mismatchColors = mismatchCols.getBackgrounds();
    mismatchRowIndexes.forEach(function(idx) {
      mismatchColors[idx][0] = '#f4cccc';
      mismatchColors[idx][1] = '#f4cccc';
    });
    mismatchCols.setBackgrounds(mismatchColors);
  }

  SpreadsheetApp.getUi().alert(
    '完了！\n' + rows.length + '件のSKU別在庫データを更新しました。\n' +
    (mismatchRowIndexes.length > 0
      ? '⚠️ 推定販売数とSquare販売数が一致しない行が' + mismatchRowIndexes.length + '件あります（セルを赤くしています）。'
      : '推定販売数とSquare販売数の不一致はありませんでした。')
  );
}


// ============================================================
// ロイヤリティレポート!C13・納品確認書!C13に、現在の会期を選ぶプルダウンを設定する
// 候補は「会期ごと作家情報」のA列（会期）からユニーク値を収集する
// ============================================================
function setupRoyaltyPeriodDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const royaltySheet = ss.getSheetByName(ROYALTY_SHEET_NAME);
  const invoiceSheet = ss.getSheetByName(INVOICE_SHEET_NAME);
  const periodSheet = getPeriodArtistSheet(); // 第6回からマスタファイル側を参照

  if (!royaltySheet) {
    SpreadsheetApp.getUi().alert('シート「' + ROYALTY_SHEET_NAME + '」が見つかりません。');
    return;
  }
  if (!periodSheet) {
    SpreadsheetApp.getUi().alert('マスタファイルに「' + PERIOD_ARTIST_SHEET + '」シートが見つかりません。MASTER_FILE_IDの設定を確認してください。');
    return;
  }

  const lastRow = periodSheet.getLastRow();
  const data = lastRow > 1 ? periodSheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];
  const periodSet = new Set();
  data.forEach(function(r) {
    const v = String(r[0] || '').trim();
    if (v) periodSet.add(v);
  });

  const periods = Array.from(periodSet);
  if (periods.length === 0) {
    SpreadsheetApp.getUi().alert('「' + PERIOD_ARTIST_SHEET + '」のA列に会期データが見つかりませんでした。');
    return;
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(periods, true)
    .setAllowInvalid(true) // 新しい会期表記が来てもブロックしない
    .build();
  royaltySheet.getRange('C13').setDataValidation(rule);

  let message = '完了！\n' + ROYALTY_SHEET_NAME + '!C13に' + periods.length + '件の会期プルダウンを設定しました。';

  if (invoiceSheet) {
    invoiceSheet.getRange('C13').setDataValidation(rule);
    message += '\n' + INVOICE_SHEET_NAME + '!C13にも同じプルダウンを設定しました。';
  } else {
    message += '\n（' + INVOICE_SHEET_NAME + 'シートが見つからなかったため、そちらは設定していません。）';
  }

  SpreadsheetApp.getUi().alert(message);
}


// ============================================================
// 「会期ごと作家情報」の中から、現在の会期かつ作家名のトークンが重なる行を探す
// トークン照合はcheckArtistDuplicateと同じ仕組み（括弧内・本体をそれぞれ比較）。
// 完全一致でなくても「saya」と「saya(Rem)」のように一部が重なれば候補に挙がる。
// 戻り値：候補の配列 [{ rowNumber, artistName }]
// ============================================================
function findPeriodArtistCandidates(targetName, currentPeriod) {
  const periodSheet = getPeriodArtistSheet(); // 第6回からマスタファイル側を参照
  if (!periodSheet) return [];

  const periodData = periodSheet.getDataRange().getValues();
  const targetTokens = extractTokens(targetName);
  const candidates = [];

  for (let i = 1; i < periodData.length; i++) {
    const rowPeriod = String(periodData[i][0] || '').trim();
    if (rowPeriod !== currentPeriod) continue;

    const rowArtist = String(periodData[i][2] || '').trim(); // C列：作家名（XLOOKUP表示）
    if (!rowArtist) continue;

    const rowTokens = extractTokens(rowArtist);
    const overlap = targetTokens.some(function(t) { return rowTokens.indexOf(t) !== -1; });
    if (overlap) {
      candidates.push({ rowNumber: i + 1, artistName: rowArtist });
    }
  }
  return candidates;
}


// ============================================================
// 作家名・対象列から「会期ごと作家情報」の該当行を探してチェックを入れる共通処理
// 現在の会期はロイヤリティレポート!C13を見る
// 名前はトークン照合でゆるく一致させる（完全一致でなくてもOK）。
// 候補が複数見つかったときは「別人かもしれない」ので自動チェックせず、要確認として返す。
// 戻り値：{ status: 'matched'|'ambiguous'|'no_match'|'no_period', candidates: string[] }
// ============================================================
function markFormReceived(artistName, targetCols) {
  const cols = Array.isArray(targetCols) ? targetCols : [targetCols];

  if (!artistName) return { status: 'no_match', candidates: [] };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const royaltySheet = ss.getSheetByName(ROYALTY_SHEET_NAME);
  const periodSheet = getPeriodArtistSheet(); // 第6回からマスタファイル側を参照
  if (!royaltySheet || !periodSheet) return { status: 'no_match', candidates: [] };

  const currentPeriod = String(royaltySheet.getRange('C13').getValue() || '').trim();
  if (!currentPeriod) return { status: 'no_period', candidates: [] }; // 現在の会期が未設定なら何もしない

  const candidates = findPeriodArtistCandidates(artistName, currentPeriod);

  if (candidates.length === 0) {
    return { status: 'no_match', candidates: [] };
  }
  if (candidates.length > 1) {
    // 別人の可能性があるため、自動チェックはせず確認待ちにする
    return { status: 'ambiguous', candidates: candidates.map(function(c) { return c.artistName; }) };
  }

  cols.forEach(function(col) {
    periodSheet.getRange(candidates[0].rowNumber, col).setValue(true);
  });
  return { status: 'matched', candidates: [candidates[0].artistName] };
}


// ============================================================
// フォーム受取チェックを手動で再実行する
// 作家マスタを後から修正した場合など、onFormSubmit実行時点と作家名が
// 変わっていてチェックが付かなかったケースをまとめて拾い直す
// ============================================================
function recheckFormReceipts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let matchedCount = 0;
  const unmatchedNames = new Set();
  const ambiguousLines = new Set();

  function processForm(sheetName, artistColIndex, targetCol) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let r = 1; r < data.length; r++) {
      const artistName = String(data[r][artistColIndex] || '').trim();
      if (!artistName) continue;
      const result = markFormReceived(artistName, targetCol);
      if (result.status === 'matched') {
        matchedCount++;
      } else if (result.status === 'ambiguous') {
        ambiguousLines.add(artistName + ' → ' + result.candidates.join(' / '));
      } else {
        unmatchedNames.add(artistName);
      }
    }
  }

  processForm(FORM_SHEET_PLAN, PLAN_COL_ARTIST_NAME, [10]);       // J列：納品予定フォーム受取
  processForm(FORM_SHEET_FINAL, FINAL_COL_ARTIST_NAME, [11, 10]); // K列：納品確定フォーム受取＋J列：納品予定フォーム受取（確定提出は予定を兼ねるとみなす）

  let message = '完了！\n' + matchedCount + '件のチェックを反映しました。';

  if (ambiguousLines.size > 0) {
    message += '\n\n⚠️ 候補が複数あり自動チェックしなかったもの（別人の可能性があるため手動で確認してください）：\n' +
      Array.from(ambiguousLines).join('\n');
  }
  if (unmatchedNames.size > 0) {
    message += '\n\n⚠️ 一致しなかった作家名（会期未設定・No.未入力・表記違いの可能性）：\n' +
      Array.from(unmatchedNames).join('\n');
  }
  SpreadsheetApp.getUi().alert(message);
}


// ============================================================
// フォーム送信時に自動実行される簡易トリガー
// 納品予定・納品確定フォームの送信を検知し、
// 「会期ごと作家情報」の該当行（現在の会期＝ロイヤリティレポート!C13、作家名のトークンが一致する行）に
// 受取済みチェックを入れる。候補が複数（別人かもしれない）場合は、
// フォーム回答シートの作家名セルにメモを付けて知らせるだけにし、自動チェックはしない。
// ============================================================
function onFormSubmit(e) {
  try {
    const submittedSheetName = e.range.getSheet().getName();
    let artistNameColIndex; // 0始まり
    let targetCols; // 会期ごと作家情報での列番号（1始まり）の配列

    if (submittedSheetName === FORM_SHEET_PLAN) {
      artistNameColIndex = PLAN_COL_ARTIST_NAME;
      targetCols = [10]; // J列：納品予定フォーム受取
    } else if (submittedSheetName === FORM_SHEET_FINAL) {
      artistNameColIndex = FINAL_COL_ARTIST_NAME;
      targetCols = [11, 10]; // K列：納品確定フォーム受取＋J列：納品予定フォーム受取（確定提出は予定を兼ねるとみなす）
    } else {
      return; // 対象外のフォームは無視
    }

    const submittedRow = e.range.getRow();
    const nameCell = e.range.getSheet().getRange(submittedRow, artistNameColIndex + 1);
    const artistName = String(nameCell.getValue() || '').trim();
    if (!artistName) return;

    const result = markFormReceived(artistName, targetCols);
    if (result.status === 'ambiguous') {
      nameCell.setNote(
        '⚠️ 会期ごと作家情報で複数候補に一致したため、フォーム受取チェックは自動で入れていません。\n' +
        '候補：' + result.candidates.join(' / ') + '\n' +
        '別人の可能性があるため、手動で確認してチェックしてください。'
      );
    }
  } catch (err) {
    // フォーム送信処理自体を失敗させたくないので、ここでは握りつぶす
    console.error('onFormSubmit エラー: ' + err);
  }
}


// ※searchArtistAndFillNumber関数は第6回からマスタファイル側の専用スクリプトに移管した。
//   （会期ごと作家情報がこのファイルに存在しなくなり、「今選んでいるセル」に依存する
//   作りのため、会期ごと作家情報を直接開いているマスタファイル側に置く必要がある）
