/***** 設定 *****/
const CONFIG = {
  SRC_SHEET_NAME: 'トリケア_CSV',      // トリケアトプスCSV貼り付け先
  DEST_SHEET_NAME: 'マネーフォワード', // マネフォ取込テンプレ
  CALC_SHEET_NAME: '計算データ',       // 途中結果（A:名前, B:①, C:②, D:③, E:④, F:⑤, G:⑥, H:⑦）
  LABELS: {
    REDUCTION: '利用減算',
    MGMT: '定期巡回総合マネジメント体制加算Ⅰ',
    INIT: '定期巡回初期加算',
    EXTRA_NAME: '定期巡回処遇改善加算Ⅱ',
  },
  EXTRA_ADD_AMOUNT: 1364, // ※今は使用しないが残しておく（固定1364案の名残）
  // 参照列（A=1）※半角英字で指定
  COLS: {
    NAME_BZ: 'BZ',   // 利用者名（キー）
    ITEM_CD: 'CD',   // サービス名/名目（サービス略称）
    UNIT_CE: 'CE',   // 単価/点数
    MULTI_FROM: 'CV',// 1カウント範囲開始
    MULTI_TO: 'DZ',  // 1カウント範囲終了
    RATE_BG: 'BG',   // 乗率（BGの数値をそのまま使用）
  },
  HAS_HEADER: true, // 1行目がヘッダー
};

/***** ユーティリティ：列記号→番号（A=1） ラベル付き *****/
function colIdx(col, keyLabel) {
  const label = keyLabel || '不明キー';
  if (col == null || String(col).trim() === '') {
    throw new Error(`colIdx: 列記号が未指定です（${label}）`);
  }
  col = String(col).trim().toUpperCase();
  if (!/^[A-Z]+$/.test(col)) {
    throw new Error(`colIdx: 列記号が不正です（${label} = "${col}"）`);
  }
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - 64);
  }
  return n;
}

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/***** ユーティリティ：BGのパース *****/
function parseBg_(value) {
  if (value === '' || value === null || typeof value === 'undefined') {
    return 1; // デフォルト1
  }
  const n = Number(value);
  if (!Number.isFinite(n)) {
    return 1; // 数値でなければ1
  }
  return n; // 0 もそのまま通す
}

/***** ユーティリティ：1円未満切り捨て（今は未使用だが残しておく） *****/
function roundDownYen_(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  return Math.floor(n);
}

/***** ユーティリティ：小数第2位まで（四捨五入） *****/
function roundTo2_(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  return Math.round(n * 100) / 100;  // 小数第2位で四捨五入
}

/***** ユーティリティ：1円単位で切り上げ *****/
function roundUpYen_(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  return Math.ceil(n);
}

/***** メニュー *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('マネフォ連携')
    .addItem('変換を実行', 'runConversion')
    .addItem('設定チェック', 'debugConfig')
    .addToUi();
}

/***** 設定チェック（どの列が怪しいかを見る用） *****/
function debugConfig() {
  const pairs = Object.entries(CONFIG.COLS).map(([k, v]) => `${k}: ${v}`);
  let idxMsg = '';
  try {
    const c = CONFIG.COLS;
    const idx = {
      NAME_BZ:    colIdx(c.NAME_BZ,    'CONFIG.COLS.NAME_BZ'),
      ITEM_CD:    colIdx(c.ITEM_CD,    'CONFIG.COLS.ITEM_CD'),
      UNIT_CE:    colIdx(c.UNIT_CE,    'CONFIG.COLS.UNIT_CE'),
      MULTI_FROM: colIdx(c.MULTI_FROM, 'CONFIG.COLS.MULTI_FROM'),
      MULTI_TO:   colIdx(c.MULTI_TO,   'CONFIG.COLS.MULTI_TO'),
      RATE_BG:    colIdx(c.RATE_BG,    'CONFIG.COLS.RATE_BG'),
    };
    idxMsg =
      '\n\n列番号\n' +
      Object.entries(idx)
        .map(([k, v]) => `${k}: ${v}`)
        .join('\n');
  } catch (e) {
    idxMsg = '\n\n列番号エラー：' + e.message;
  }
  SpreadsheetApp.getUi().alert('CONFIG.COLS\n' + pairs.join('\n') + idxMsg);
}

/***** メイン処理 *****/
function runConversion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CONFIG.SRC_SHEET_NAME);
  const dest = ss.getSheetByName(CONFIG.DEST_SHEET_NAME);
  const calcSheet = getOrCreateSheet_(CONFIG.CALC_SHEET_NAME);
  if (!src || !dest) {
    throw new Error('シート名の確認：トリケア_CSV / マネーフォワード');
  }

  // 列index（CONFIG.COLS）
  const c = CONFIG.COLS;
  const idx = {
    NAME_BZ:    colIdx(c.NAME_BZ,    'CONFIG.COLS.NAME_BZ'),
    ITEM_CD:    colIdx(c.ITEM_CD,    'CONFIG.COLS.ITEM_CD'),
    UNIT_CE:    colIdx(c.UNIT_CE,    'CONFIG.COLS.UNIT_CE'),
    MULTI_FROM: colIdx(c.MULTI_FROM, 'CONFIG.COLS.MULTI_FROM'),
    MULTI_TO:   colIdx(c.MULTI_TO,   'CONFIG.COLS.MULTI_TO'),
    RATE_BG:    colIdx(c.RATE_BG,    'CONFIG.COLS.RATE_BG'),
  };

  // マネフォ側列
  const colA  = colIdx('A',  'A');
  const colB  = colIdx('B',  'B');
  const colC  = colIdx('C',  'C');
  const colN  = colIdx('N',  'N');
  const colAD = colIdx('AD', 'AD');
  const colAF = colIdx('AF', 'AF');
  const colAG = colIdx('AG', 'AG');
  const colAL = colIdx('AL', 'AL');

  // === 係数（計算データ!I2） ===
  const coefRaw = calcSheet.getRange('I2').getValue();
  let coef;
  if (coefRaw === '' || coefRaw === null) {
    coef = 0.1; // 未設定なら 0.1
  } else {
    const n = Number(coefRaw);
    if (!Number.isFinite(n)) {
      throw new Error('係数（計算データ!I2）が数値ではありません。');
    }
    coef = n;
  }

  // === 処遇改善Ⅱ率（計算データ!J2：％として扱い、n*0.01 に変換） ===
  const extraRateRaw = calcSheet.getRange('J2').getValue();
  let extraRate = 0; // 0〜1（例：22.4% → 0.224）
  if (extraRateRaw != null && extraRateRaw !== '') {
    const s = String(extraRateRaw).replace('%', '').trim();
    const n = Number(s);
    if (Number.isFinite(n)) {
      extraRate = n * 0.01; // 入力値 × 0.01（10 → 0.1）
    }
  }

  // === CV〜DZ内で「1」を数える ===
  const isOne = (v) => {
    if (v === 1 || v === true) return true;
    if (v == null) return false;
    const s = String(v).trim();
    return s === '1' || s === '１';
  };
  const countOnes = (row) => {
    let count = 0;
    const from = idx.MULTI_FROM - 1;
    const to = Math.min(idx.MULTI_TO - 1, row.length - 1);
    for (let i = from; i <= to; i++) if (isOne(row[i])) count++;
    return count;
  };

  // === ソースデータ取得 ===
  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();
  const startRow = CONFIG.HAS_HEADER ? 2 : 1;
  if (lastRow < startRow) {
    SpreadsheetApp.getUi().alert('トリケア_CSVにデータ行がありません。');
    return;
  }
  const values = src.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();

  /** 集計用：名前ごとの①〜④＋BG 等 */
  const userAgg = {};
  const isRegularRange = (label) => {
    const s = String(label).trim();
    // 定期巡回随時Ⅱ１〜５（全角・半角数字OK）
    return /^定期巡回随時Ⅱ[1-5１-５]$/.test(s);
  };

  /** ★ 請求書明細用：名前 → Map(サービス略称 → { cd, unit, qty, bg }) */
  const invoiceAgg = {}; // { [name: string]: Map<string, { cd, unit, qty, bg }> }

  // === 行ごとに集計 ===
  for (const row of values) {
    const name = String(row[idx.NAME_BZ - 1] ?? '').trim();
    const item = String(row[idx.ITEM_CD - 1] ?? '').trim();
    if (!name && !item) continue;

    const unit = Number(row[idx.UNIT_CE - 1]) || 0;
    const bg   = parseBg_(row[idx.RATE_BG - 1]);

    // ====== ①〜④＋BG の集計（名前単位） ======
    if (name) {
      if (!userAgg[name]) {
        userAgg[name] = { one: 0, red: 0, mgmt: 0, init: 0, rate: bg };
      }
      const agg = userAgg[name];
      agg.rate = bg;

      if (isRegularRange(item)) {
        agg.one += unit;                        // ①
      } else if (item === CONFIG.LABELS.REDUCTION || /利用\s*減算/.test(item)) {
        agg.red += unit * countOnes(row);       // ②
      } else if (item === CONFIG.LABELS.MGMT) {
        agg.mgmt += unit * countOnes(row);      // ③
      } else if (item === CONFIG.LABELS.INIT) {
        agg.init += unit * countOnes(row);      // ④
      }
    }

    // ====== ★ 名前×サービス略称ごとの単位数集計（品目行用） ======
    let qty = 0;
    if (isRegularRange(item)) {
      qty = 1;
    } else if (
      item === CONFIG.LABELS.REDUCTION ||
      item === CONFIG.LABELS.MGMT ||
      item === CONFIG.LABELS.INIT ||
      /利用\s*減算/.test(item)
    ) {
      qty = countOnes(row);
    }

    if (name && item && qty > 0) {
      if (!invoiceAgg[name]) invoiceAgg[name] = new Map();
      const map = invoiceAgg[name];
      const key = item; // サービス略称（CD列）でキー
      if (!map.has(key)) {
        map.set(key, { cd: item, unit: unit, qty: 0, bg: bg });
      }
      const rec = map.get(key);
      rec.qty  += qty;
      rec.unit  = unit;
      rec.bg    = bg;   // BGも保持（AF計算用）
    }
  }

  // === ①〜⑤・⑥（理論値）を userAgg にセット（トリケア世界線） ===
  for (const [name, agg] of Object.entries(userAgg)) {
    const one   = agg.one  || 0;
    const two   = agg.red  || 0;
    const three = agg.mgmt || 0;
    const four  = agg.init || 0;
    const base1234 = one + two + three + four;

    // トリケア世界線の処遇改善Ⅱ（①〜④の合計 × J2%）→ 四捨五入して整数
    let extraTricareInt = 0;
    if (extraRate > 0 && base1234 > 0) {
      const extraTricareRaw = base1234 * extraRate;
      extraTricareInt = Math.round(extraTricareRaw); // 四捨五入（整数）
    }

    const five = base1234 + extraTricareInt;             // ⑤：①〜④＋処遇改善Ⅱ（整数）
    const rate = Number(agg.rate) || 1;
    const sixRaw = coef * five * rate;
    const six = roundUpYen_(sixRaw);                     // ⑥：係数×BG×⑤ を「切り上げ」

    agg.base1234       = base1234;
    agg.extraTricare   = extraTricareInt;
    agg.five           = five;
    agg.theoreticalG   = six;       // トリケア理論値（G列）
    agg.mfTotal        = 0;         // 後で H列用に入れる
  }

  // === マネフォ：2行目以降クリア ===
  const destLastRow = dest.getLastRow();
  const destLastCol = dest.getLastColumn();
  if (destLastRow >= 2) {
    dest.getRange(2, 1, destLastRow - 1, destLastCol).clearContent();
  }

  // === マネフォ出力用の配列を作成（高速化） ===
  const destRows = [];
  const sortedNames = Object.keys(userAgg).sort((a, b) => a.localeCompare(b, 'ja'));

  for (const name of sortedNames) {
    const agg = userAgg[name];

    // 1) 請求書行：A=40101, B=請求書, C=名前, N=様
    {
      const rowArr = new Array(colAL).fill('');
      rowArr[colA  - 1] = '40101';
      rowArr[colB  - 1] = '請求書';
      rowArr[colC  - 1] = name;
      rowArr[colN  - 1] = '様';
      destRows.push(rowArr);
    }

    // 2) 品目行：この名前に関係する品目のみ（マネフォ世界線の合計もここで集計）
    const map = invoiceAgg[name];
    let mfTotalInt = 0; // マネフォ世界線：各行ごと AF×AG を四捨五入（整数）して合計

    if (map) {
      const items = Array.from(map.values());
      items.sort((a, b) => String(a.cd).localeCompare(String(b.cd), 'ja'));

      for (const rec of items) {
        const ce  = rec.unit;      // CE（元単価）
        const qty = rec.qty;
        const bg  = rec.bg;

        // AF(単価) = CE × 係数I2 × BG → 小数第2位まで
        const afRaw = ce * coef * bg;
        const af    = roundTo2_(afRaw);

        // 行ごとの金額（AF×AG）→ 四捨五入して整数（マネフォ世界線）
        const lineAmountDec = af * qty;
        const lineAmountInt = Math.round(lineAmountDec); // 例：191.73×12=2300.76 → 2301

        mfTotalInt += lineAmountInt;

        const rowArr = new Array(colAL).fill('');
        rowArr[colA  - 1] = '40101';
        rowArr[colB  - 1] = '品目';
        rowArr[colAD - 1] = rec.cd;   // AD = サービス略称
        rowArr[colAF - 1] = af;       // AF = 単価（係数×BG込み、小数第2位）
        rowArr[colAG - 1] = qty;      // AG = 単位数
        rowArr[colAL - 1] = '非課税'; // AL = 非課税

        destRows.push(rowArr);
      }
    }

    // 3) ★ 定期巡回処遇改善加算Ⅱ 行（マネフォ世界線＋トリケア理論値に合わせた調整）
    let extraInt = 0; // マネフォ側での処遇改善Ⅱの「金額（整数）」行

    if (extraRate > 0 && mfTotalInt > 0) {
      // ベースは「通常品目の 金額(AF×AG)」合計（すでに整数）
      const extraMfRaw = mfTotalInt * extraRate;
      // まずは通常通りに計算（小数第2位丸め → 整数丸め）
      const extraMfDec = roundTo2_(extraMfRaw);
      extraInt = Math.round(extraMfDec); // 「素の」処遇改善Ⅱ（マネフォ世界線、整数）
    }

    // ---- トリケア理論値とのズレをここで吸収 ----
    const theoretical = agg.theoreticalG || 0;          // G列の理論値
    const mfTotalBeforeAdjust = mfTotalInt + extraInt;  // 調整前のマネフォ合計（整数想定）
    let diff = theoretical - mfTotalBeforeAdjust;       // diff がプラスなら不足、マイナスなら超過

    // 処遇改善Ⅱ行の金額に diff をそのまま載せて調整
    let extraIntAdjusted = extraInt + diff;
    // 万が一、0以下になった場合は0として扱う（行を出さない）
    if (extraIntAdjusted < 0) {
      extraIntAdjusted = 0;
    }

    const finalMfTotal = mfTotalInt + extraIntAdjusted; // 調整後マネフォ合計
    agg.mfTotal = finalMfTotal;                         // H列に出す値

    if (extraIntAdjusted > 0) {
      const rowArr = new Array(colAL).fill('');
      rowArr[colA  - 1] = '40101';
      rowArr[colB  - 1] = '品目';
      rowArr[colAD - 1] = CONFIG.LABELS.EXTRA_NAME; // AD = 定期巡回処遇改善加算Ⅱ
      rowArr[colAF - 1] = extraIntAdjusted;         // AF = 金額（1件分、整数）
      rowArr[colAG - 1] = 1;
      rowArr[colAL - 1] = '非課税';

      destRows.push(rowArr);
    }
  }

  // === マネフォシートへ一括書き込み ===
  if (destRows.length > 0) {
    dest.getRange(2, 1, destRows.length, colAL).setValues(destRows);
  }

  // === 計算データシート出力（A〜H列） ===
  calcSheet.clearContents();
  const header = [[
    '名前',
    '①（随時Ⅱ）',
    '②（利用減算）',
    '③（マネジメントⅠ）',
    '④（初期加算）',
    '⑤（①〜④＋処遇改善Ⅱ 四捨五入）',
    '⑥（理論値：係数×BG×⑤ 切り上げ）',
    '⑦（マネフォ世界線 合計：調整後）'
  ]];
  calcSheet.getRange(1, 1, 1, header[0].length).setValues(header);

  // 係数と処遇改善Ⅱ率（I1〜J2）を再セット
  calcSheet.getRange('I1').setValue('係数（I2：患者負担率）');
  calcSheet.getRange('I2').setValue(coefRaw);
  calcSheet.getRange('J1').setValue('処遇改善Ⅱ率（J2, %）');
  calcSheet.getRange('J2').setValue(extraRateRaw);

  const calcRows = [];
  for (const name of sortedNames) {
    const agg = userAgg[name];
    calcRows.push([
      name,
      agg.one            || 0,  // B列 ①
      agg.red            || 0,  // C列 ②
      agg.mgmt           || 0,  // D列 ③
      agg.init           || 0,  // E列 ④
      agg.five           || 0,  // F列 ⑤
      agg.theoreticalG   || 0,  // G列 ⑥（理論値：切り上げ）
      agg.mfTotal        || 0,  // H列 ⑦（マネフォ側合計：調整後）
    ]);
  }

  if (calcRows.length > 0) {
    calcSheet.getRange(2, 1, calcRows.length, header[0].length).setValues(calcRows);
  }

  SpreadsheetApp.getUi().alert(
    '変換が完了しました。\n' +
    '・計算データ!F列：①〜④の合計 × J2% を四捨五入した処遇改善Ⅱを足した値\n' +
    '・計算データ!G列：F列にBGとI2を掛けた値を1円単位で切り上げ（トリケア理論値）\n' +
    '・計算データ!H列：マネフォ世界線の合計（各行AF×AGを四捨五入）を、\n' +
    '　　定期巡回処遇改善加算Ⅱ行の金額で微調整して、理論値に一致させた値です。'
  );
}

/***** 旧互換（ボタン用） *****/
function myFunction() {
  runConversion();
}


/***** ここから Webアプリ用  *****************************************/

// スプレッドシートのURL（ボタンから開く用）
const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1xpyRU8WK92rdJ_6Eb4uv6sNhXYya-nDQa04s73ccrCc/edit';

// Webアプリの入口
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  template.SPREADSHEET_URL = SPREADSHEET_URL;
  template.MF_URL = 'https://invoice.moneyforward.com/billings';
  template.TRICARE_URL = 'https://www.tricare.jp/tricare/TZAD010.do';

  return template
    .evaluate()
    .setTitle('トリケア → マネーフォワード 変換ツール');
}

/**
 * ブラウザから送られた CSV文字列を
 * 「トリケア_CSV」に上書き → runConversion() 実行
 */
function importCsvAndConvert(csvText) {
  if (!csvText) {
    throw new Error('CSVの中身が空です。ファイルを確認してください。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = ss.getSheetByName(CONFIG.SRC_SHEET_NAME);
  if (!srcSheet) {
    throw new Error('シート「' + CONFIG.SRC_SHEET_NAME + '」が見つかりません。');
  }

  // 一度すべて消してから貼り付け
  srcSheet.clearContents();

  // CSV文字列 → 2次元配列
  // ※ブラウザ側で Shift_JIS → 文字列 に変換済みなので、ここでは通常のparseCsvでOK
  const lines = csvText.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const rows = lines
    .filter(line => line !== '')
    .map(line => Utilities.parseCsv(line)[0]);

  if (!rows || rows.length === 0) {
    throw new Error('CSVから行を読み取れませんでした。');
  }

  srcSheet
    .getRange(1, 1, rows.length, rows[0].length)
    .setValues(rows);

  // 既存の変換ロジックを実行
  runConversion();

  return {
    status: 'ok',
    rowCount: rows.length,
  };
}

/**
 * 「マネーフォワード」シートの内容を CSV文字列にして返す
 */
function getProcessedCsv() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dest = ss.getSheetByName(CONFIG.DEST_SHEET_NAME);
  if (!dest) {
    throw new Error('シート「' + CONFIG.DEST_SHEET_NAME + '」が見つかりません。');
  }

  const lastRow = dest.getLastRow();
  const lastCol = dest.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    throw new Error('マネーフォワードシートにデータがありません。');
  }

  const values = dest.getRange(1, 1, lastRow, lastCol).getValues();

  // 配列 → CSV文字列（カンマ区切り、必要な場合はダブルクォート）
  const lines = values.map(row =>
    row.map(v => {
      if (v == null) v = '';
      const s = String(v);
      if (s.includes('"') || s.includes(',') || s.includes('\n')) {
        return '"' + s.replace(/"/g, '""') + '"';
      }
      return s;
    }).join(',')
  );

  const csv = lines.join('\r\n');

  // ファイル名：yyyyMMdd_tricare_to_mf.csv
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = ('0' + (now.getMonth() + 1)).slice(-2);
  const dd = ('0' + now.getDate()).slice(-2);
  const filename = `${yyyy}${mm}${dd}_tricare_to_mf.csv`;

  return {
    filename,
    content: csv,
  };
}

/***** I2 / J2 をWebアプリから読み書きするAPI ************************/

/**
 * 計算データ!I2 / J2 の現在値を取得
 * - I2 が空 → 0.1 を返す
 * - J2 が空 → 0   を返す（％）
 */
function getRatesForWeb() {
  const sheet = getOrCreateSheet_(CONFIG.CALC_SHEET_NAME);

  // I2：係数（患者負担率）
  const coefRaw = sheet.getRange('I2').getValue();
  let coefDisplay;
  if (coefRaw === '' || coefRaw === null) {
    coefDisplay = 0.1; // 未設定なら 0.1
  } else {
    const n = Number(coefRaw);
    coefDisplay = Number.isFinite(n) ? n : 0.1;
  }

  // J2：処遇改善Ⅱ率（％）
  const extraRaw = sheet.getRange('J2').getValue();
  let extraDisplay;
  if (extraRaw === '' || extraRaw === null) {
    extraDisplay = 0; // 未設定なら 0%
  } else {
    // 「22.4」「22.4%」どちらでもOKにする
    const s = String(extraRaw).replace('%', '').trim();
    const n = Number(s);
    extraDisplay = Number.isFinite(n) ? n : 0;
  }

  return {
    coef: coefDisplay,              // I2 にそのまま入れる値
    extraRatePercent: extraDisplay, // J2 にそのまま入れる値（％）
  };
}

/**
 * Webから送られた I2 / J2 をシートに保存するだけ
 * 画面上では少数第一位が基本
 */
function updateRatesFromWeb(coef, extraRatePercent) {
  const sheet = getOrCreateSheet_(CONFIG.CALC_SHEET_NAME);

  // 文字→数値に変換（空のときはデフォルトを入れる）
  let c = coef;
  if (c === '' || c === null || typeof c === 'undefined') {
    c = 0.1;
  }
  c = Number(c);
  if (!Number.isFinite(c)) {
    throw new Error('患者負担率（I2）が数値ではありません。');
  }

  let e = extraRatePercent;
  if (e === '' || e === null || typeof e === 'undefined') {
    e = 0; // デフォルト 0%
  }
  e = Number(e);
  if (!Number.isFinite(e)) {
    throw new Error('処遇改善Ⅱ率（J2, %）が数値ではありません。');
  }

  // 少数第1位にそろえて保存
  const cRounded = Math.round(c * 10) / 10;
  const eRounded = Math.round(e * 10) / 10;

  sheet.getRange('I2').setValue(cRounded); // 「患者負担率」
  sheet.getRange('J2').setValue(eRounded); // 「処遇改善Ⅱ率（%）」

  return {
    status: 'ok',
    coef: cRounded,
    extraRatePercent: eRounded,
  };
}

/***** HTMLテンプレート読み込み用（そのままおまじないでOK） *****/
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
