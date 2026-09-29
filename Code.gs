/**
 * Webアプリケーションとしてアクセスされた際にIndex.htmlをサーブします。
 * GASのHTMLサービスを利用し、メタタグやタイトルを適切に設定します。
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('営業管理ダッシュボード EXECUTIVE VIEW')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 「ダッシュボード設定」シートからマッピング情報を取得します。
 * シートが存在しない場合は、デフォルト設定を書き込んだシートを自動生成します。
 * 
 * @return {Object} カラムキーとスプレッドシートのヘッダー名のマッピングオブジェクト
 */
function getMappingConfigOrInit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ダッシュボード設定');
  
  // デフォルトのマッピング設定定義
  var defaultConfig = [
    ['ダッシュボード用キー', 'スプレッドシート列ヘッダー名', '説明（変更不可）'],
    ['id', 'ID', '案件の一意なID。空欄でも可。'],
    ['projectId', '案件ID', '案件のコードや管理ID。'],
    ['clientCategory', '取引先カテゴリ', '取引先の規模区分（大手企業、中堅企業、スタートアップ等）'],
    ['clientName', '取引先名', '取引先の正式名称。外注サマリーの集計にも使用。'],
    ['summary', '案件概要', '案件の名称や具体的な仕事内容。'],
    ['status', 'ステータス', '商談ステータス列（フェーズ数字を含んだステータス名）。'],
    ['costCategory', '費用分類', '「売上」または「外注費用」と完全一致する文字列。'],
    ['billingMonth', '請求年月', '請求または支払いの発生月（YYYY-MM形式、または日付オブジェクト）。'],
    ['amountExpected', '金額税抜（想定）', '売上または外注費用の想定金額（円単位）。'],
    ['amountConfirmed', '金額（確定）', '売上または外注費用の確定済金額（円単位）。'],
    ['year', '売上年度', '集計対象となる年度区分（例: 26年度）。'],
    ['vendor', '発注先', '外注先（発注先）の名称（外注費用の集計に使用）。'],
    ['orderContent', '発注内容', '外注の発注内容や具体的な作業内容。'],
    ['estimateRequest', '外注費用の見積依頼', '外注費用の見積依頼ステータスや金額等の情報。']
  ];
  
  if (!sheet) {
    sheet = ss.insertSheet('ダッシュボード設定');
    sheet.getRange(1, 1, defaultConfig.length, defaultConfig[0].length).setValues(defaultConfig);
    sheet.getRange('A1:C1').setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(3, 350);
    sheet.getRange(1, 1, defaultConfig.length, 2).setHorizontalAlignment('left');
  }
  
  var data = sheet.getDataRange().getValues();
  
  // 既存のキーを特定
  var existingKeys = {};
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0] ? data[i][0].toString().trim() : '';
    if (key) {
      existingKeys[key] = true;
    }
  }
  
  // 不足しているマッピング設定を自動でシートへ追記
  var appended = false;
  for (var j = 1; j < defaultConfig.length; j++) {
    var defKey = defaultConfig[j][0];
    if (!existingKeys[defKey]) {
      sheet.appendRow(defaultConfig[j]);
      appended = true;
    }
  }
  
  // 追記した場合は再度シートデータを取得
  if (appended) {
    data = sheet.getDataRange().getValues();
  }
  
  var mapping = {};
  
  // 2行目から読み込み (ヘッダーを除く)
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0] ? data[i][0].toString().trim() : '';
    var headerName = data[i][1] ? data[i][1].toString().trim() : '';
    if (key) {
      mapping[key] = headerName;
    }
  }
  
  return mapping;
}

/**
 * フロントエンドからのカラムマッピング設定を「ダッシュボード設定」シートに上書き保存します。
 * 
 * @param {Object} mapping カラムキーと列ヘッダー名のマッピング連想配列
 * @return {boolean} 成功時にtrue
 */
function saveMappingConfig(mapping) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ダッシュボード設定');
  if (!sheet) {
    getMappingConfigOrInit(); // なければ生成
    sheet = ss.getSheetByName('ダッシュボード設定');
  }
  
  var range = sheet.getDataRange();
  var data = range.getValues();
  
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0] ? data[i][0].toString().trim() : '';
    if (key && mapping[key] !== undefined) {
      data[i][1] = mapping[key].toString().trim();
    }
  }
  
  range.setValues(data);
  return true;
}

/**
 * 「月別予算設定」シートから月別の予算額（円）を取得します。
 * シートが存在しない場合は、25〜26年度の初期設定値を自動生成します。
 * 
 * @return {Object} 年月（YYYY-MM）をキー、予算（円）を値とする連想配列
 */
function getMonthlyBudgetsOrInit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('月別予算設定');
  
  // デフォルト予算設定のデータ定義 (25〜30年度分を動的ループ生成)
  var defaultBudgets = [['売上年度', '年月', '予算金額（円）']];
  var yearsList = ['25年度', '26年度', '27年度', '28年度', '29年度', '30年度'];
  for (var yIdx = 0; yIdx < yearsList.length; yIdx++) {
    var yrStr = yearsList[yIdx];
    var yrNum = 2000 + parseInt(yrStr.replace('年度', ''), 10);
    // 各年度 4月〜翌年3月
    for (var m = 4; m <= 12; m++) {
      defaultBudgets.push([yrStr, yrNum + '-' + (m < 10 ? '0' + m : m), 25000000]);
    }
    for (var m = 1; m <= 3; m++) {
      defaultBudgets.push([yrStr, (yrNum + 1) + '-' + ('0' + m), 25000000]);
    }
  }
  
  if (!sheet) {
    sheet = ss.insertSheet('月別予算設定');
    sheet.getRange(1, 1, defaultBudgets.length, defaultBudgets[0].length).setValues(defaultBudgets);
    sheet.getRange('A1:C1').setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 180);
    sheet.getRange(1, 1, defaultBudgets.length, 3).setHorizontalAlignment('left');
  }
  
  var data = sheet.getDataRange().getValues();
  var budgets = {};
  for (var i = 1; i < data.length; i++) {
    var month = data[i][1] ? formatBillingMonth(data[i][1]) : '';
    var amount = parseFloat(data[i][2]) || 0;
    if (month) {
      budgets[month] = amount;
    }
  }
  return budgets;
}

/**
 * フロントエンドから送信された月別予算情報を「月別予算設定」シートに上書き保存します。
 * 新しい年月が含まれる場合は自動で追加し、売上年度を自動的に類推します。
 * 
 * @param {Object} budgets 年月（YYYY-MM）をキー、予算（円）を値とする連想配列
 * @return {boolean} 成功時にtrue
 */
function saveMonthlyBudgets(budgets) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('月別予算設定');
  if (!sheet) {
    getMonthlyBudgetsOrInit(); // なければ初期化
    sheet = ss.getSheetByName('月別予算設定');
  }
  
  // 1. 現在のシートデータをすべて取得
  var range = sheet.getDataRange();
  var data = range.getValues();
  
  // 重複行を検出して削除
  var monthRowIndexMap = {};
  var rowsToDelete = [];
  
  for (var i = 1; i < data.length; i++) {
    var rawMonth = data[i][1];
    var formattedMonth = rawMonth ? formatBillingMonth(rawMonth) : '';
    
    if (formattedMonth) {
      if (monthRowIndexMap[formattedMonth] === undefined) {
        monthRowIndexMap[formattedMonth] = i + 1; // 1-indexedの行番号
      } else {
        rowsToDelete.push(i + 1);
      }
    }
  }
  
  // 重複行を後ろから削除（インデックスずれ防止）
  rowsToDelete.sort(function(a, b) { return b - a; });
  rowsToDelete.forEach(function(rowIdx) {
    sheet.deleteRow(rowIdx);
  });
  
  // 2. 削除後の最新データを再ロードして更新/追加を実行
  range = sheet.getDataRange();
  data = range.getValues();
  
  monthRowIndexMap = {};
  for (var i = 1; i < data.length; i++) {
    var rawMonth = data[i][1];
    var formattedMonth = rawMonth ? formatBillingMonth(rawMonth) : '';
    if (formattedMonth) {
      monthRowIndexMap[formattedMonth] = i + 1;
    }
  }
  
  for (var m in budgets) {
    var amount = parseFloat(budgets[m]) || 0;
    var rowIdx = monthRowIndexMap[m];
    
    if (rowIdx !== undefined) {
      // 既存の年月行がある場合はC列（3列目）を上書き更新
      sheet.getRange(rowIdx, 3).setValue(amount);
    } else {
      // 存在しない年月は新規追加
      var yearStr = inferSalesYearFromMonth(m);
      sheet.appendRow([yearStr, m, amount]);
    }
  }
  
  // 3. データ行（2行目以降）を年月（B列: 2列目）で昇順ソート
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 3).sort({column: 2, ascending: true});
  }
  
  return true;
}

/**
 * 年月 (YYYY-MM) から売上年度 (例: 26年度) を推測するヘルパー関数
 */
function inferSalesYearFromMonth(monthStr) {
  if (!monthStr || !monthStr.includes('-')) return '未設定';
  var parts = monthStr.split('-');
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10);
  
  // 日本の標準的な会計年度（4月〜翌3月）
  var targetYear = y;
  if (m >= 1 && m <= 3) {
    targetYear = y - 1;
  }
  var twoDigitYear = targetYear % 100;
  return twoDigitYear + '年度';
}

/**
 * 案件マスタシートからすべての生データを抽出し、フロントエンドへ高速転送します。
 * 案件データに加え、マッピング定義、月別予算設定、マスタの全ヘッダー一覧を一括で返却します。
 * 
 * @return {Object} 案件リスト、マッピング定義、月別予算定義、マスタヘッダー配列
 */
function getRawMasterData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('案件マスタ');
  if (!sheet) {
    throw new Error('「案件マスタ」シートが見つかりません。まず「案件マスタ」という名前でシートを作成してください。');
  }
  
  // 各種設定のロード（なければ自动生成）
  var mappingConfig = getMappingConfigOrInit();
  var monthlyBudgets = getMonthlyBudgetsOrInit();
  var simulationRates = getSimulationRatesOrInit();
  
  var data = sheet.getDataRange().getValues();
  var headers = [];
  if (data.length > 0) {
    headers = data[0].map(function(h) {
      return h ? h.toString().trim() : '';
    });
  }
  
  if (data.length <= 1) {
    return {
      data: [],
      mapping: mappingConfig,
      monthlyBudgets: monthlyBudgets,
      simulationRates: simulationRates,
      masterHeaders: headers,
      changeHistory: []
    };
  }
  
  // マッピング用インデックスの動的解決
  var idxMap = {};
  for (var key in mappingConfig) {
    var headerName = mappingConfig[key];
    idxMap[key] = headers.indexOf(headerName);
  }
  
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // ステータスまたは費用分類が空の行はデータ不備としてスキップ
    var statusVal = idxMap.status !== -1 && idxMap.status !== undefined ? row[idxMap.status] : '';
    var costCategoryVal = idxMap.costCategory !== -1 && idxMap.costCategory !== undefined ? row[idxMap.costCategory] : '';
    if (!statusVal && !costCategoryVal) {
      continue;
    }
    
    rows.push({
      id: idxMap.id !== -1 && idxMap.id !== undefined ? (row[idxMap.id] || '') : '',
      projectId: idxMap.projectId !== -1 && idxMap.projectId !== undefined ? (row[idxMap.projectId] || '') : '',
      clientCategory: idxMap.clientCategory !== -1 && idxMap.clientCategory !== undefined ? (row[idxMap.clientCategory] || '') : '',
      clientName: idxMap.clientName !== -1 && idxMap.clientName !== undefined ? (row[idxMap.clientName] || '') : '',
      summary: idxMap.summary !== -1 && idxMap.summary !== undefined ? (row[idxMap.summary] || '') : '',
      status: statusVal ? statusVal.toString().trim() : '',
      costCategory: costCategoryVal ? costCategoryVal.toString().trim() : '',
      billingMonth: idxMap.billingMonth !== -1 && idxMap.billingMonth !== undefined && row[idxMap.billingMonth] ? formatBillingMonth(row[idxMap.billingMonth]) : '未設定',
      amountExpected: idxMap.amountExpected !== -1 && idxMap.amountExpected !== undefined ? (parseFloat(row[idxMap.amountExpected]) || 0) : 0,
      amountConfirmed: idxMap.amountConfirmed !== -1 && idxMap.amountConfirmed !== undefined ? (parseFloat(row[idxMap.amountConfirmed]) || 0) : 0,
      year: idxMap.year !== -1 && idxMap.year !== undefined && row[idxMap.year] ? row[idxMap.year].toString().trim() : '未設定',
      vendor: idxMap.vendor !== -1 && idxMap.vendor !== undefined ? (row[idxMap.vendor] || '') : '',
      orderContent: idxMap.orderContent !== -1 && idxMap.orderContent !== undefined ? (row[idxMap.orderContent] || '') : '',
      estimateRequest: idxMap.estimateRequest !== -1 && idxMap.estimateRequest !== undefined ? (row[idxMap.estimateRequest] || '') : ''
    });
  }
  
  return {
    data: rows,
    mapping: mappingConfig,
    monthlyBudgets: monthlyBudgets,
    simulationRates: simulationRates,
    masterHeaders: headers,
    changeHistory: getChangeHistory()
  };
}

/**
 * 日付オブジェクトまたは文字列から「YYYY-MM」形式を安全に抽出するヘルパー
 * 
 * @param {*} val 日付オブジェクトまたは文字列
 * @return {string} YYYY-MM 形式 of 文字列、あるいは '未設定'
 */
function formatBillingMonth(val) {
  if (val instanceof Date) {
    var y = val.getFullYear();
    var m = val.getMonth() + 1;
    return y + '-' + (m < 10 ? '0' + m : m);
  }
  if (typeof val === 'string') {
    val = val.trim();
    // 「YYYY/MM/DD」や「YYYY/MM」の形式に対応
    if (val.includes('/')) {
      var parts = val.split('/');
      if (parts.length >= 2) {
        var yPart = parts[0];
        var mPart = parseInt(parts[1], 10);
        return yPart + '-' + (mPart < 10 ? '0' + mPart : mPart);
      }
    }
    // 「YYYY-MM-DD」形式に対応
    if (val.includes('-')) {
      var partsDash = val.split('-');
      if (partsDash.length >= 2) {
        var yPartDash = partsDash[0];
        var mPartDash = parseInt(partsDash[1], 10);
        return yPartDash + '-' + (mPartDash < 10 ? '0' + mPartDash : mPartDash);
      }
    }
  }
  return val ? val.toString() : '未設定';
}

/**
 * 案件マスタシートから現在使用されているユニークなステータス一覧を抽出します。
 * 
 * @return {Array} ユニークなステータス名の配列
 */
function getUniqueStatusesFromMaster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName('案件マスタ');
  if (!masterSheet) return [];
  
  var mapping = getMappingConfigOrInit();
  var statusHeader = mapping['status'] || 'ステータス';
  
  var data = masterSheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var headers = data[0].map(function(h) { return h ? h.toString().trim() : ''; });
  var statusIdx = headers.indexOf(statusHeader);
  if (statusIdx === -1) return [];
  
  var statuses = [];
  for (var i = 1; i < data.length; i++) {
    var statusVal = data[i][statusIdx] ? data[i][statusIdx].toString().trim() : '';
    
    // ステータスと費用分類が両方空の行は無効データとして除外
    var costIdx = headers.indexOf(mapping['costCategory'] || '費用分類');
    var costVal = costIdx !== -1 ? data[i][costIdx] : '';
    if (!statusVal && !costVal) continue;
    
    if (statusVal && statuses.indexOf(statusVal) === -1) {
      statuses.push(statusVal);
    }
  }
  return statuses;
}

/**
 * 「受注率設定」シートからステータスごとのデフォルト受注率を取得します。
 * シートが存在しない場合は、現在の案件マスタのステータスに基づいて自動生成します。
 * 
 * @return {Object} ステータス名をキー、デフォルト受注率（%）を値とする連想配列
 */
function getSimulationRatesOrInit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('受注率設定');
  
  if (!sheet) {
    sheet = ss.insertSheet('受注率設定');
    var uniqueStatuses = getUniqueStatusesFromMaster();
    
    var defaultConfig = [['ステータス', 'デフォルト受注率（%）']];
    uniqueStatuses.forEach(function(status) {
      var defaultRate = 0;
      // 6〜9（9.請求完了を含む）は受注扱い。ただし「失注」を含むものは除く
      if (status.indexOf('失注') === -1 &&
          (status.indexOf('6') === 0 || status.indexOf('7') === 0 || status.indexOf('8') === 0 || status.indexOf('9') === 0)) {
        defaultRate = 100;
      } else if (status.indexOf('4') === 0 || status.indexOf('5') === 0) {
        defaultRate = 80;
      } else if (status.indexOf('1') === 0 || status.indexOf('2') === 0 || status.indexOf('3') === 0) {
        defaultRate = 30;
      }
      defaultConfig.push([status, defaultRate]);
    });
    
    // 万が一データがなかった場合のフォールバック
    if (defaultConfig.length === 1) {
      defaultConfig.push(['1.アプローチ中', 30]);
      defaultConfig.push(['4.提案中', 80]);
      defaultConfig.push(['6.受注確定', 100]);
      defaultConfig.push(['9.請求完了', 100]);
      defaultConfig.push(['X.失注', 0]);
    }
    
    sheet.getRange(1, 1, defaultConfig.length, defaultConfig[0].length).setValues(defaultConfig);
    sheet.getRange('A1:B1').setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 180);
    sheet.getRange(1, 1, defaultConfig.length, 2).setHorizontalAlignment('left');
  }
  
  var data = sheet.getDataRange().getValues();
  var rates = {};
  for (var i = 1; i < data.length; i++) {
    var status = data[i][0] ? data[i][0].toString().trim() : '';
    var rate = parseFloat(data[i][1]);
    if (isNaN(rate)) rate = 0;
    if (status) {
      rates[status] = rate;
    }
  }
  return rates;
}

/**
 * フロントエンドから送信されたデフォルト受注率の設定情報を「受注率設定」シートに保存します。
 * 
 * @param {Object} rates ステータス名をキー、デフォルト受注率（%）を値とする連想配列
 * @return {boolean} 成功時にtrue
 */
function saveSimulationRates(rates) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('受注率設定');
  if (!sheet) {
    sheet = ss.insertSheet('受注率設定');
  }
  
  var rows = [['ステータス', 'デフォルト受注率（%）']];
  
  // 送信されたステータスをソートして書き込む
  var sortedStatuses = Object.keys(rates).sort();
  sortedStatuses.forEach(function(status) {
    var rate = parseFloat(rates[status]);
    if (isNaN(rate)) rate = 0;
    rows.push([status, rate]);
  });
  
  sheet.clear();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange('A1:B1').setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 180);
  sheet.getRange(1, 1, rows.length, 2).setHorizontalAlignment('left');
  
  return true;
}

/**
 * スプレッドシート編集時に実行されるシンプルトリガー。
 * 「案件マスタ」のステータス列の変更を検知し、「ステータス変更履歴」シートにトランザクションログを記録します。
 */
function onEdit(e) {
  if (!e) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== '案件マスタ') return;
  
  var range = e.range;
  var startRow = range.getRow();
  var startCol = range.getColumn();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  
  // 1行目のヘッダー行の編集は無視
  if (startRow === 1) return;
  
  // カラムマッピング情報の取得
  var mapping = getMappingConfigOrInit();
  var statusHeader = mapping['status'] || 'ステータス';
  var projectIdHeader = mapping['projectId'] || '案件ID';
  var clientCategoryHeader = mapping['clientCategory'] || '取引先カテゴリ';
  var clientNameHeader = mapping['clientName'] || '取引先名';
  var summaryHeader = mapping['summary'] || '案件概要';
  var costCategoryHeader = mapping['costCategory'] || '費用分類';
  var billingMonthHeader = mapping['billingMonth'] || '請求年月';
  var amountExpectedHeader = mapping['amountExpected'] || '金額税抜（想定）';
  var yearHeader = mapping['year'] || '売上年度';
  
  // 案件マスタの現在のヘッダー位置を解決
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var statusColIdx = headers.indexOf(statusHeader) + 1; // 1-indexed
  var projectIdColIdx = headers.indexOf(projectIdHeader) + 1;
  var clientCategoryColIdx = headers.indexOf(clientCategoryHeader) + 1;
  var clientNameColIdx = headers.indexOf(clientNameHeader) + 1;
  var summaryColIdx = headers.indexOf(summaryHeader) + 1;
  var costCategoryColIdx = headers.indexOf(costCategoryHeader) + 1;
  var billingMonthColIdx = headers.indexOf(billingMonthHeader) + 1;
  var amountExpectedColIdx = headers.indexOf(amountExpectedHeader) + 1;
  var yearColIdx = headers.indexOf(yearHeader) + 1;
  
  // マッピング外のカラム「営業担当」をヘッダー名で直接解決
  var salesRepColIdx = headers.indexOf('営業担当') + 1;
  
  if (statusColIdx === 0) return; // ステータス列が特定できない場合は終了
  
  // 編集範囲に「ステータス」列が含まれているか判定
  var isStatusEdited = (startCol <= statusColIdx && statusColIdx < startCol + numCols);
  if (!isStatusEdited) return;
  
  // 「ステータス変更履歴」シートを取得、なければ新規作成
  var ss = e.source;
  var historySheet = ss.getSheetByName('ステータス変更履歴');
  if (!historySheet) {
    historySheet = ss.insertSheet('ステータス変更履歴');
    var historyHeaders = [
      '更新日', 
      '案件ID', 
      '取引先カテゴリー', 
      '取引先名', 
      '案件概要', 
      '変更前ステータス', 
      '変更後ステータス', 
      '営業担当', 
      '費用分類', 
      '請求年月', 
      '金額税抜（想定）', 
      '売上年度'
    ];
    historySheet.appendRow(historyHeaders);
    historySheet.getRange('A1:L1').setBackground('#fbbc05').setFontColor('#ffffff').setFontWeight('bold'); // Googleイエロー
    historySheet.getRange('A1:L1').setHorizontalAlignment('center');
    
    // 各列の幅を設定
    historySheet.setColumnWidth(1, 170); // 更新日
    historySheet.setColumnWidth(2, 100); // 案件ID
    historySheet.setColumnWidth(3, 130); // 取引先カテゴリー
    historySheet.setColumnWidth(4, 180); // 取引先名
    historySheet.setColumnWidth(5, 200); // 案件概要
    historySheet.setColumnWidth(6, 150); // 変更前ステータス
    historySheet.setColumnWidth(7, 150); // 変更後ステータス
    historySheet.setColumnWidth(8, 120); // 営業担当
    historySheet.setColumnWidth(9, 100); // 費用分類
    historySheet.setColumnWidth(10, 100); // 請求年月
    historySheet.setColumnWidth(11, 140); // 金額税抜（想定）
    historySheet.setColumnWidth(12, 100); // 売上年度
  }
  
  var timestamp = new Date();
  
  // 編集された各行について処理
  for (var i = 0; i < numRows; i++) {
    var currentRow = startRow + i;
    if (currentRow === 1) continue; // ヘッダー行は除外
    
    // 各項目の値を取得
    var projectId = projectIdColIdx > 0 ? sheet.getRange(currentRow, projectIdColIdx).getValue() : '';
    var clientCategory = clientCategoryColIdx > 0 ? sheet.getRange(currentRow, clientCategoryColIdx).getValue() : '';
    var clientName = clientNameColIdx > 0 ? sheet.getRange(currentRow, clientNameColIdx).getValue() : '';
    var summary = summaryColIdx > 0 ? sheet.getRange(currentRow, summaryColIdx).getValue() : '';
    var status = sheet.getRange(currentRow, statusColIdx).getValue();
    var salesRep = salesRepColIdx > 0 ? sheet.getRange(currentRow, salesRepColIdx).getValue() : '';
    var costCategory = costCategoryColIdx > 0 ? sheet.getRange(currentRow, costCategoryColIdx).getValue() : '';
    
    // 請求年月は日付型や文字列型があるので安全に取得・整形
    var rawBillingMonth = billingMonthColIdx > 0 ? sheet.getRange(currentRow, billingMonthColIdx).getValue() : '';
    var billingMonth = rawBillingMonth ? formatBillingMonth(rawBillingMonth) : '未設定';
    
    var amountExpected = amountExpectedColIdx > 0 ? sheet.getRange(currentRow, amountExpectedColIdx).getValue() : 0;
    var year = yearColIdx > 0 ? sheet.getRange(currentRow, yearColIdx).getValue() : '';
    
    // 変更前の値を取得（単一セル編集で既存値がある場合は e.oldValue。空白セルへの入力や新規行は '(新規作成)'）
    var oldValue = '(新規作成)';
    if (numRows === 1 && numCols === 1) {
      if (e.oldValue !== undefined && e.oldValue !== '') {
        oldValue = e.oldValue;
      }
    } else {
      oldValue = '(一括作成)';
    }
    
    // 実際に値が変更されている場合のみ、指定の12カラムで履歴に書き出し
    if (oldValue !== status) {
      historySheet.appendRow([
        timestamp,      // 1. 更新日
        projectId,      // 2. 案件ID
        clientCategory, // 3. 取引先カテゴリー
        clientName,     // 4. 取引先名
        summary,        // 5. 案件概要
        oldValue,       // 6. 変更前ステータス
        status,         // 7. 変更後ステータス
        salesRep,       // 8. 営業担当
        costCategory,   // 9. 費用分類
        billingMonth,   // 10. 請求年月
        amountExpected, // 11. 金額税抜（想定）
        year            // 12. 売上年度
      ]);
    }
  }
}

/**
 * 履歴データを取得します。履歴シートが存在しない、または空の場合は初期化処理を自動実行します。
 */
function getChangeHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ステータス変更履歴');
  if (!sheet || sheet.getLastRow() <= 1) {
    initializeHistoryFromMaster();
    sheet = ss.getSheetByName('ステータス変更履歴');
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
  var history = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    history.push({
      timestamp: row[0] instanceof Date ? row[0].toISOString() : row[0].toString(),
      projectId: row[1] ? row[1].toString() : '',
      clientCategory: row[2] ? row[2].toString() : '',
      clientName: row[3] ? row[3].toString() : '',
      summary: row[4] ? row[4].toString() : '',
      oldStatus: row[5] ? row[5].toString() : '',
      newStatus: row[6] ? row[6].toString() : '',
      salesRep: row[7] ? row[7].toString() : '',
      costCategory: row[8] ? row[8].toString() : '',
      billingMonth: row[9] ? row[9].toString() : '',
      amountExpected: parseFloat(row[10]) || 0,
      year: row[11] ? row[11].toString() : ''
    });
  }
  return history;
}

/**
 * 「ステータス変更履歴」シートがない、または空の時に、案件マスタの既存データから初期履歴データを作成します。
 */
function initializeHistoryFromMaster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var historySheet = ss.getSheetByName('ステータス変更履歴');
  if (historySheet && historySheet.getLastRow() > 1) return; // すでにデータがある場合はスキップ
  
  if (!historySheet) {
    historySheet = ss.insertSheet('ステータス変更履歴');
  }
  
  historySheet.clear();
  var historyHeaders = [
    '更新日', '案件ID', '取引先カテゴリー', '取引先名', '案件概要', 
    '変更前ステータス', '変更後ステータス', '営業担当', '費用分類', 
    '請求年月', '金額税抜（想定）', '売上年度'
  ];
  historySheet.appendRow(historyHeaders);
  historySheet.getRange('A1:L1').setBackground('#fbbc05').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  historySheet.setColumnWidth(1, 170);
  historySheet.setColumnWidth(2, 100);
  historySheet.setColumnWidth(3, 130);
  historySheet.setColumnWidth(4, 180);
  historySheet.setColumnWidth(5, 200);
  historySheet.setColumnWidth(6, 150);
  historySheet.setColumnWidth(7, 150);
  historySheet.setColumnWidth(8, 120);
  historySheet.setColumnWidth(9, 100);
  historySheet.setColumnWidth(10, 100);
  historySheet.setColumnWidth(11, 140);
  historySheet.setColumnWidth(12, 100);
  
  var masterSheet = ss.getSheetByName('案件マスタ');
  if (!masterSheet) return;
  
  var masterData = masterSheet.getDataRange().getValues();
  if (masterData.length <= 1) return;
  
  var mapping = getMappingConfigOrInit();
  var headers = masterData[0].map(function(h) { return h ? h.toString().trim() : ''; });
  
  var idxMap = {};
  for (var key in mapping) {
    idxMap[key] = headers.indexOf(mapping[key]);
  }
  var salesRepColIdx = headers.indexOf('営業担当');
  
  var rowsToAppend = [];
  for (var i = 1; i < masterData.length; i++) {
    var row = masterData[i];
    
    var statusVal = idxMap.status !== -1 && idxMap.status !== undefined ? row[idxMap.status] : '';
    var costCategoryVal = idxMap.costCategory !== -1 && idxMap.costCategory !== undefined ? row[idxMap.costCategory] : '';
    if (!statusVal && !costCategoryVal) continue;
    
    var projectId = idxMap.projectId !== -1 && idxMap.projectId !== undefined ? row[idxMap.projectId] : '';
    var clientCategory = idxMap.clientCategory !== -1 && idxMap.clientCategory !== undefined ? row[idxMap.clientCategory] : '';
    var clientName = idxMap.clientName !== -1 && idxMap.clientName !== undefined ? row[idxMap.clientName] : '';
    var summary = idxMap.summary !== -1 && idxMap.summary !== undefined ? row[idxMap.summary] : '';
    var salesRep = salesRepColIdx !== -1 ? row[salesRepColIdx] : '';
    
    var rawBillingMonth = idxMap.billingMonth !== -1 && idxMap.billingMonth !== undefined ? row[idxMap.billingMonth] : '';
    var billingMonth = rawBillingMonth ? formatBillingMonth(rawBillingMonth) : '未設定';
    var amountExpected = idxMap.amountExpected !== -1 && idxMap.amountExpected !== undefined ? (parseFloat(row[idxMap.amountExpected]) || 0) : 0;
    var year = idxMap.year !== -1 && idxMap.year !== undefined ? row[idxMap.year] : '';
    
    // 更新日（タイムスタンプ）の決定。過去の受注月集計と連動させるため、請求年月の月の1日をセット。
    var timestamp = new Date();
    if (billingMonth && billingMonth !== '未設定' && billingMonth.includes('-')) {
      var parts = billingMonth.split('-');
      var y = parseInt(parts[0], 10);
      var m = parseInt(parts[1], 10);
      timestamp = new Date(y, m - 1, 1, 9, 0, 0); // 請求年月の月の1日
    }
    
    rowsToAppend.push([
      timestamp,
      projectId,
      clientCategory,
      clientName,
      summary,
      '(初期登録)', // 変更前ステータス
      statusVal,    // 変更後ステータス
      salesRep,
      costCategoryVal,
      billingMonth,
      amountExpected,
      year
    ]);
  }
  
  if (rowsToAppend.length > 0) {
    historySheet.getRange(2, 1, rowsToAppend.length, 12).setValues(rowsToAppend);
  }
}
