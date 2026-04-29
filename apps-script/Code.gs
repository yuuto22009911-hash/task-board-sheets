/**
 * task-board-sheets — Google Sheets タスク管理表
 *
 * シート構成:
 *   1. タスク一覧
 *   2. 会議日程 (Google Calendar の予定一覧)
 *   3. リマインドメールタスク
 *   4. 企業リスト
 *   5. 日報
 *   6. URL
 *   7. 設定 (カレンダー連携の設定)
 *
 * 主な機能:
 *   - メニュー「タスク管理」から各操作
 *   - タスク追加で先頭行に新規行を挿入
 *   - 状況プルダウンの色分け (条件付き書式)
 *   - 会議日程の自動同期 (毎朝6時)
 *   - SMTG リマインドの自動生成 (毎朝6時、土日祝スキップ)
 *   - 日報フォーマットコピー
 *   - 改善案の投稿
 */

// ===== 定数 =====
var SHEET_TASKS    = 'タスク一覧';
var SHEET_CALENDAR = '会議日程';
var SHEET_REMIND   = 'リマインドメールタスク';
var SHEET_COMPANY  = '企業リスト';
var SHEET_REPORT   = '日報';
var SHEET_URL      = 'URL';
var SHEET_SETTINGS = '設定';
var SHEET_IDEAS    = '改善アイデア';

var STATUS_LIST   = ['未着手', '進行中', '依頼中', '完了'];
var PRIORITY_LIST = ['低', '中', '高'];
var UNIT_LIST     = ['SP', 'CM', 'AI', 'BO', '秘書', 'HR', 'DW'];

var COLOR_RED        = '#fde7e9'; // 未着手
var COLOR_YELLOW     = '#fff4cc'; // 進行中
var COLOR_GRAY       = '#e8eaed'; // 依頼中
var COLOR_GREEN      = '#e6f4ea'; // 完了
var COLOR_RED_BD     = '#d93025';
var COLOR_YELLOW_BD  = '#f29900';
var COLOR_GRAY_BD    = '#80868b';
var COLOR_GREEN_BD   = '#137333';

var HEADER_BG    = '#f1f3f4';
var HEADER_COLOR = '#202124';
var BORDER_COLOR = '#dadce0';

// ===== メニュー =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('タスク管理')
    .addItem('初期セットアップ (最初の1回)', 'setupAll')
    .addItem('スキーマ更新 (データを保持して書式のみ再適用)', 'updateSchema')
    .addSeparator()
    .addItem('タスク追加 (上の行に挿入)',     'addTaskRow')
    .addItem('会議日程を今すぐ更新',           'syncCalendarEvents')
    .addItem('SMTG リマインドを今すぐ作成',    'scanSmtgReminders')
    .addItem('日報フォーマットを表示 (コピー用)', 'showDailyReportTemplate')
    .addSeparator()
    .addItem('利用可能なカレンダー一覧を確認', 'listAvailableCalendars')
    .addItem('カレンダー設定を開く',           'openSettings')
    .addItem('改善アイデアを送る',             'showIdeasDialog')
    .addItem('使い方を表示',                   'showHelp')
    .addToUi();
}

// ===== 初期セットアップ (新規スプレッドシート用) =====
function setupAll() {
  _setup_(true);
}

// ===== スキーマ更新 (既存データは消さず、ヘッダー/書式/データ検証だけ再適用) =====
function updateSchema() {
  _setup_(false);
}

function _setup_(freshInstall) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupTasksSheet_(ss, freshInstall);
  setupCalendarSheet_(ss, freshInstall);
  setupRemindSheet_(ss, freshInstall);
  setupCompanySheet_(ss, freshInstall);
  setupReportSheet_(ss, freshInstall);
  setupUrlSheet_(ss, freshInstall);
  setupSettingsSheet_(ss, freshInstall);
  setupIdeasSheet_(ss, freshInstall);
  // 初期表示は タスク一覧 シート
  ss.setActiveSheet(ss.getSheetByName(SHEET_TASKS));
  // デフォルトの「シート1」が残っていれば削除
  var blank = ss.getSheetByName('シート1');
  if (blank) ss.deleteSheet(blank);
  // 自動トリガーをセット
  installDailyTriggers_();
  var msg = freshInstall
    ? '7つのシートを新規作成しました。\n\n次のステップ:\n1. 「設定」シートでカレンダーIDを確認 (複数指定可、"all"で全カレンダー)\n2. メニュー「会議日程を今すぐ更新」で予定一覧を取り込み\n3. メニュー「SMTG リマインドを今すぐ作成」で動作確認\n\n毎朝6時の自動同期トリガーも仕掛け済みです。'
    : '既存データは保持したまま、書式・データ検証・自動トリガーを最新状態に更新しました。';
  SpreadsheetApp.getUi().alert(freshInstall ? '初期セットアップ完了' : 'スキーマ更新完了', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function setupTasksSheet_(ss, fresh) {
  var existed = !!ss.getSheetByName(SHEET_TASKS);
  var sheet = getOrCreateSheet_(ss, SHEET_TASKS);
  if (fresh || !existed) sheet.clear();
  sheet.setHiddenGridlines(false);

  var headers = ['タスク追加日', '企業名', 'タスク内容', '期日', '優先度', '詳細', '状況'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet.getRange(1, 1, 1, headers.length));

  setColumnWidths_(sheet, [110, 160, 320, 120, 90, 240, 100]);

  setDateValidation_(sheet, 4);
  setListValidation_(sheet, 5, PRIORITY_LIST);
  setListValidation_(sheet, 7, STATUS_LIST);
  applyStatusConditionalFormat_(sheet, 7);

  sheet.setFrozenRows(1);
  ensureMinRows_(sheet, 50);
}

function setupCalendarSheet_(ss, fresh) {
  var existed = !!ss.getSheetByName(SHEET_CALENDAR);
  var sheet = getOrCreateSheet_(ss, SHEET_CALENDAR);
  if (fresh || !existed) {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  }

  var headers = ['開始日時', '終了日時', '予定タイトル', '場所', '説明', '終日', 'カレンダー名', 'イベントID'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet.getRange(1, 1, 1, headers.length));
  setColumnWidths_(sheet, [150, 150, 320, 180, 280, 60, 180, 280]);

  sheet.getRange(2, 5, sheet.getMaxRows() - 1, 1).setWrap(true).setVerticalAlignment('top');

  // 既存の今日ハイライトルールを除去してから再適用 (重複防止)
  removeRulesForRange_(sheet, sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()));
  applyTodayHighlight_(sheet, 1);

  sheet.setFrozenRows(1);
  ensureMinRows_(sheet, 50);
}

function setupRemindSheet_(ss, fresh) {
  var existed = !!ss.getSheetByName(SHEET_REMIND);
  var sheet = getOrCreateSheet_(ss, SHEET_REMIND);
  if (fresh || !existed) sheet.clear();
  var headers = ['期日', 'SMTG 開催日時', '対象予定タイトル', '状況', 'カレンダーイベントID'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet.getRange(1, 1, 1, headers.length));
  setColumnWidths_(sheet, [120, 180, 360, 100, 280]);

  setListValidation_(sheet, 4, ['未着手', '完了']);
  applyStatusConditionalFormat_(sheet, 4);

  sheet.setFrozenRows(1);
  ensureMinRows_(sheet, 30);
}

function setupCompanySheet_(ss, fresh) {
  var existed = !!ss.getSheetByName(SHEET_COMPANY);
  var sheet = getOrCreateSheet_(ss, SHEET_COMPANY);
  if (fresh || !existed) sheet.clear();
  var headers = ['企業名', 'ユニット', 'アサイン日', '支援期間', '議事録URL', '支援責任者', '支援担当者'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet.getRange(1, 1, 1, headers.length));
  setColumnWidths_(sheet, [200, 100, 130, 130, 280, 140, 140]);

  setListValidation_(sheet, 2, UNIT_LIST);
  setDateValidation_(sheet, 3);

  sheet.setFrozenRows(1);
  ensureMinRows_(sheet, 30);
}

function setupReportSheet_(ss, fresh) {
  var existed = !!ss.getSheetByName(SHEET_REPORT);
  var sheet = getOrCreateSheet_(ss, SHEET_REPORT);
  if (fresh || !existed) sheet.clear();
  var headers = ['日付', '内容'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet.getRange(1, 1, 1, headers.length));
  setColumnWidths_(sheet, [140, 800]);
  setDateValidation_(sheet, 1);

  sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1).setWrap(true).setVerticalAlignment('top');

  sheet.setFrozenRows(1);
  ensureMinRows_(sheet, 30);
}

function setupUrlSheet_(ss, fresh) {
  var existed = !!ss.getSheetByName(SHEET_URL);
  var sheet = getOrCreateSheet_(ss, SHEET_URL);
  if (fresh || !existed) sheet.clear();
  var headers = ['タイトル', 'URL', 'カテゴリ', 'メモ'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet.getRange(1, 1, 1, headers.length));
  setColumnWidths_(sheet, [220, 360, 140, 280]);
  sheet.setFrozenRows(1);
  ensureMinRows_(sheet, 30);
}

function setupSettingsSheet_(ss, fresh) {
  var existed = !!ss.getSheetByName(SHEET_SETTINGS);
  var sheet = getOrCreateSheet_(ss, SHEET_SETTINGS);

  // 既存の値を保持しながら説明列・スタイルだけ更新する
  var existingValues = {};
  if (existed && sheet.getLastRow() >= 2) {
    var prev = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    prev.forEach(function(row) {
      if (row[0]) existingValues[row[0]] = row[1];
    });
  }

  if (fresh || !existed) sheet.clear();
  setColumnWidths_(sheet, [220, 380, 320]);

  var defaults = {
    'カレンダーID': 'primary',
    'SMTG キーワード': 'SMTG',
    '先読み日数': 14,
    '過去日数': 3,
    '日報テンプレート': defaultDailyReportTemplate_(),
  };
  var rows = [
    ['設定項目', '値', '説明'],
    ['カレンダーID',
      pickValue_(existingValues, 'カレンダーID', defaults['カレンダーID']),
      'メイン1個だけなら "primary" でOK。\n複数のカレンダーから取得したい時は改行(Alt+Enter)で複数IDを入力。\n"all" と書くとあなたがアクセスできる全カレンダーをスキャン。\nIDは Google Calendar 「設定と共有 → カレンダーの統合 → カレンダーID」で取得 (...@group.calendar.google.com 等)。'],
    ['SMTG キーワード',
      pickValue_(existingValues, 'SMTG キーワード', defaults['SMTG キーワード']),
      '予定タイトルにこのキーワード(部分一致、大小無視)を含む予定をリマインド対象にします。'],
    ['先読み日数',
      pickValue_(existingValues, '先読み日数', defaults['先読み日数']),
      '今日から何日先までのカレンダーをスキャンするか(整数)。会議日程シートとSMTGリマインドの両方に使われます。'],
    ['過去日数',
      pickValue_(existingValues, '過去日数', defaults['過去日数']),
      '会議日程シートに何日前までの予定を残すか(整数)。0 にすると今日以降のみ。'],
    ['日報テンプレート',
      pickValue_(existingValues, '日報テンプレート', defaults['日報テンプレート']),
      '日報の雛形。{DATE} が今日の日付に置き換わる。'],
  ];
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  styleHeader_(sheet.getRange(1, 1, 1, 3));

  sheet.getRange(2, 2, rows.length - 1, 1).setBackground('#fffbf0').setVerticalAlignment('top').setWrap(true);
  sheet.getRange(1, 1, rows.length, 3).setBorder(true, true, true, true, true, true, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setFrozenRows(1);
  sheet.getRange(2, 1, rows.length - 1, 3).setVerticalAlignment('top').setWrap(true);
}

function pickValue_(map, key, fallback) {
  return map.hasOwnProperty(key) && map[key] !== '' && map[key] != null ? map[key] : fallback;
}

function setupIdeasSheet_(ss, fresh) {
  var existed = !!ss.getSheetByName(SHEET_IDEAS);
  var sheet = getOrCreateSheet_(ss, SHEET_IDEAS);
  if (fresh || !existed) sheet.clear();
  var headers = ['投稿日時', 'カテゴリ', '内容'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet.getRange(1, 1, 1, headers.length));
  setColumnWidths_(sheet, [150, 120, 600]);
  sheet.setFrozenRows(1);
  if (fresh || !existed) sheet.hideSheet();
}

// ===== タスク追加 (上の行に挿入) =====
function addTaskRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_TASKS);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('先に「初期セットアップ」を実行してください');
    return;
  }
  // 2行目(ヘッダーの直下)に新規行を挿入
  sheet.insertRowBefore(2);
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  sheet.getRange(2, 1).setValue(today);
  sheet.getRange(2, 7).setValue('未着手');

  // データ検証/書式は insertRowBefore でも維持されるはずだが、念のため再適用
  setDateValidation_(sheet, 4);
  setListValidation_(sheet, 5, PRIORITY_LIST);
  setListValidation_(sheet, 7, STATUS_LIST);

  ss.setActiveSheet(sheet);
  // 企業名セルにフォーカス
  sheet.setActiveRange(sheet.getRange(2, 2));
}

// ===== カレンダー解決 (複数指定 / "all" 対応) =====
function getCalendars_(settings) {
  var ids = settings.calendarIds || ['primary'];
  // "all" または "*" で全カレンダー
  for (var i = 0; i < ids.length; i++) {
    if (ids[i] === 'all' || ids[i] === '*') {
      return CalendarApp.getAllCalendars();
    }
  }
  var result = [];
  for (var j = 0; j < ids.length; j++) {
    var id = ids[j];
    var c = (id === 'primary')
      ? CalendarApp.getDefaultCalendar()
      : CalendarApp.getCalendarById(id);
    if (c) result.push(c);
  }
  return result;
}

// ===== 会議日程の同期 (カレンダー全予定をシートに反映) =====
function syncCalendarEvents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CALENDAR);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('先に「初期セットアップ」を実行してください');
    return;
  }

  var settings = readSettings_();
  var calendars = getCalendars_(settings);

  if (!calendars || calendars.length === 0) {
    SpreadsheetApp.getUi().alert(
      'カレンダーが見つかりません',
      '「設定」シートのカレンダーIDを確認してください。\n"primary" / "all" / 正しいカレンダーID のいずれか。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // 範囲: (今日 - 過去日数) 〜 (今日 + 先読み日数)
  var now = new Date();
  var pastDays = settings.pastDays || 0;
  var start = new Date(now.getTime() - pastDays * 24 * 60 * 60 * 1000);
  start.setHours(0, 0, 0, 0);
  var end = new Date(now.getTime() + settings.lookaheadDays * 24 * 60 * 60 * 1000);
  end.setHours(23, 59, 59, 999);

  // 全カレンダーから予定を集めてマージ
  var allEvents = [];
  for (var i = 0; i < calendars.length; i++) {
    var cal = calendars[i];
    try {
      var evs = cal.getEvents(start, end);
      for (var j = 0; j < evs.length; j++) {
        allEvents.push({ ev: evs[j], calName: cal.getName() });
      }
    } catch (e) {
      // アクセス権が無いカレンダーはスキップ
      console.warn('Skip calendar ' + cal.getId() + ': ' + e);
    }
  }

  // 開始時刻昇順
  allEvents.sort(function(a, b) {
    return a.ev.getStartTime().getTime() - b.ev.getStartTime().getTime();
  });

  // 既存データをクリア (ヘッダー以外)
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 8).clearContent();
  }

  if (allEvents.length === 0) {
    SpreadsheetApp.getUi().alert('対象期間に予定はありません', 'スキャン対象カレンダー数: ' + calendars.length, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var rows = allEvents.map(function(item) {
    var ev = item.ev;
    var startTime = ev.getStartTime();
    var endTime   = ev.getEndTime();
    var allDay    = ev.isAllDayEvent();
    return [
      Utilities.formatDate(startTime, Session.getScriptTimeZone(), allDay ? 'yyyy/MM/dd (E)' : 'yyyy/MM/dd (E) HH:mm'),
      Utilities.formatDate(endTime,   Session.getScriptTimeZone(), allDay ? 'yyyy/MM/dd (E)' : 'yyyy/MM/dd (E) HH:mm'),
      ev.getTitle() || '',
      ev.getLocation() || '',
      truncate_(ev.getDescription() || '', 500),
      allDay ? '○' : '',
      item.calName,
      ev.getId()
    ];
  });

  sheet.getRange(2, 1, rows.length, 8).setValues(rows);
  sheet.getRange(2, 5, rows.length, 1).setWrap(true).setVerticalAlignment('top');

  SpreadsheetApp.getUi().alert(
    '会議日程の同期完了',
    '対象カレンダー: ' + calendars.length + ' 個\n' +
    '取得期間: ' + Utilities.formatDate(start, Session.getScriptTimeZone(), 'M/d') +
    ' 〜 ' + Utilities.formatDate(end, Session.getScriptTimeZone(), 'M/d') + '\n' +
    '反映件数: ' + rows.length + ' 件',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function truncate_(s, max) {
  if (!s) return '';
  s = String(s);
  return s.length > max ? s.slice(0, max) + '...' : s;
}

// ===== SMTG リマインド自動生成 =====
function scanSmtgReminders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingSheet = ss.getSheetByName(SHEET_SETTINGS);
  var remindSheet  = ss.getSheetByName(SHEET_REMIND);
  if (!settingSheet || !remindSheet) {
    SpreadsheetApp.getUi().alert('先に「初期セットアップ」を実行してください');
    return;
  }

  var settings = readSettings_();
  var calendars = getCalendars_(settings);
  if (!calendars || calendars.length === 0) {
    SpreadsheetApp.getUi().alert('カレンダーが見つかりません', '「設定」シートのカレンダーIDを確認してください。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var now    = new Date();
  var future = new Date(now.getTime() + settings.lookaheadDays * 24 * 60 * 60 * 1000);
  var keyword = settings.smtgKeyword.toLowerCase();
  var existing = readExistingEventIds_(remindSheet);

  var created = 0, skipped = 0, totalMatched = 0;
  for (var c = 0; c < calendars.length; c++) {
    var cal = calendars[c];
    try {
      var events = cal.getEvents(now, future);
      for (var i = 0; i < events.length; i++) {
        var ev = events[i];
        var title = ev.getTitle() || '';
        if (title.toLowerCase().indexOf(keyword) === -1) continue;
        totalMatched++;
        var evId = ev.getId();
        if (existing[evId]) { skipped++; continue; }
        var startDate = ev.getStartTime();
        var dueDate = previousBusinessDay_(startDate);
        appendRemindRow_(remindSheet, dueDate, startDate, title, evId);
        created++;
      }
    } catch (e) {
      console.warn('Skip calendar ' + cal.getId() + ': ' + e);
    }
  }

  sortRemindSheet_(remindSheet);

  SpreadsheetApp.getUi().alert(
    'SMTG リマインドのスキャン結果',
    'スキャン対象カレンダー: ' + calendars.length + ' 個\n' +
    'マッチした予定: ' + totalMatched + ' 件\n' +
    '新規作成: ' + created + ' 件\n' +
    '既存で重複(スキップ): ' + skipped + ' 件',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function appendRemindRow_(sheet, dueDate, startDate, title, evId) {
  var lastRow = sheet.getLastRow();
  var nextRow = Math.max(2, lastRow + 1);
  sheet.getRange(nextRow, 1, 1, 5).setValues([[
    Utilities.formatDate(dueDate,   Session.getScriptTimeZone(), 'yyyy/MM/dd (E)'),
    Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy/MM/dd (E) HH:mm'),
    title,
    '未着手',
    evId
  ]]);
}

function readExistingEventIds_(sheet) {
  var lastRow = sheet.getLastRow();
  var map = {};
  if (lastRow < 2) return map;
  var values = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0]) map[values[i][0]] = true;
  }
  return map;
}

function sortRemindSheet_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 2) return;
  sheet.getRange(2, 1, lastRow - 1, 5).sort([{column: 1, ascending: true}]);
}

// ===== 営業日計算 (土日 + 日本祝日スキップ) =====
function previousBusinessDay_(date) {
  var d = new Date(date.getTime());
  d.setDate(d.getDate() - 1);
  while (!isBusinessDay_(d)) {
    d.setDate(d.getDate() - 1);
  }
  return d;
}

function isBusinessDay_(date) {
  var day = date.getDay();
  if (day === 0 || day === 6) return false;
  return !isJapanHoliday_(date);
}

function isJapanHoliday_(date) {
  // Google が公開している日本の祝日カレンダーを参照
  var holidayCal = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
  if (!holidayCal) return false;
  var start = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0);
  var end   = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59);
  var events = holidayCal.getEvents(start, end);
  return events.length > 0;
}

// ===== 日報テンプレート表示 =====
function showDailyReportTemplate() {
  var settings = readSettings_();
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd (E)');
  var template = (settings.reportTemplate || defaultDailyReportTemplate_()).replace(/\{DATE\}/g, today);

  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:-apple-system,Segoe UI,sans-serif;margin:0;padding:16px}textarea{width:100%;height:280px;font-family:monospace;font-size:13px;line-height:1.55;padding:10px;box-sizing:border-box;border:1px solid #dadce0;border-radius:6px}p{font-size:12px;color:#5f6368;margin:8px 0}button{background:#1a73e8;color:#fff;border:0;padding:8px 18px;border-radius:6px;cursor:pointer;font-size:13px}</style>' +
    '<p>下のテンプレートをコピーして使ってください。「日報」シートに直接貼り付けて記録できます。</p>' +
    '<textarea id="t" readonly>' + escapeHtml_(template) + '</textarea>' +
    '<p><button onclick="navigator.clipboard.writeText(document.getElementById(\'t\').value); google.script.host.close()">クリップボードにコピーして閉じる</button></p>'
  ).setWidth(560).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, '日報フォーマット (' + today + ')');
}

function defaultDailyReportTemplate_() {
  return '日付: {DATE}\n\n本日の作業:\n- \n\n進捗:\n- \n\n課題:\n- \n\n明日の予定:\n- ';
}

// ===== 設定シート読み取り =====
function readSettings_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var map = {};
  for (var i = 0; i < values.length; i++) {
    map[values[i][0]] = values[i][1];
  }
  var rawId = String(map['カレンダーID'] || 'primary');
  // 改行 / カンマ / セミコロン区切りで分解、空白除去
  var ids = rawId.split(/[\n,;]+/).map(function(s){ return s.trim(); }).filter(function(s){ return s; });
  if (ids.length === 0) ids = ['primary'];
  return {
    calendarIds:    ids,
    calendarId:     ids[0],  // 互換用 (旧コード参照向け)
    smtgKeyword:    map['SMTG キーワード']  || 'SMTG',
    lookaheadDays:  parseInt(map['先読み日数'] || '14', 10),
    pastDays:       parseInt(map['過去日数']   || '3', 10),
    reportTemplate: map['日報テンプレート']  || defaultDailyReportTemplate_(),
  };
}

// ===== 利用可能なカレンダー一覧を確認 (デバッグ用) =====
function listAvailableCalendars() {
  var calendars = CalendarApp.getAllCalendars();
  var rows = calendars.map(function(c) {
    var id = c.getId();
    var name = c.getName();
    var isPrimary = (c.getId() === CalendarApp.getDefaultCalendar().getId());
    return { id: id, name: name, isPrimary: isPrimary };
  });
  rows.sort(function(a, b) {
    if (a.isPrimary && !b.isPrimary) return -1;
    if (b.isPrimary && !a.isPrimary) return 1;
    return a.name.localeCompare(b.name);
  });

  var html = '<style>body{font-family:-apple-system,Segoe UI,sans-serif;margin:0;padding:18px;font-size:13px}h2{font-size:15px;margin:0 0 6px}p{color:#5f6368;margin:4px 0 14px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #dadce0;padding:7px 9px;text-align:left;font-size:12.5px;vertical-align:top}th{background:#f1f3f4}td.id{font-family:monospace;font-size:11.5px;word-break:break-all}.tag{display:inline-block;background:#1a73e8;color:#fff;font-size:10.5px;padding:1px 6px;border-radius:4px;margin-left:6px}</style>';
  html += '<h2>利用可能なカレンダー一覧</h2>';
  html += '<p>' + rows.length + '個のカレンダーが利用可能です。下の表のカレンダーIDをコピーして「設定」シートに貼り付けると、そのカレンダーから予定を取得できます。<br>複数指定する場合は改行(Alt+Enter)で区切ります。すべて取得したい場合は <code>all</code> と入力するだけで OK。</p>';
  html += '<table><tr><th>カレンダー名</th><th>カレンダーID</th></tr>';
  rows.forEach(function(r) {
    html += '<tr><td>' + escapeHtml_(r.name) + (r.isPrimary ? '<span class="tag">primary</span>' : '') + '</td><td class="id">' + escapeHtml_(r.id) + '</td></tr>';
  });
  html += '</table>';

  var output = HtmlService.createHtmlOutput(html).setWidth(700).setHeight(540);
  SpreadsheetApp.getUi().showModalDialog(output, '利用可能なカレンダー一覧');
}

// ===== 設定シートを開く =====
function openSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('先に「初期セットアップ」を実行してください');
    return;
  }
  ss.setActiveSheet(sheet);
}

// ===== 改善案ダイアログ =====
function showIdeasDialog() {
  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:-apple-system,Segoe UI,sans-serif;margin:0;padding:16px}select,textarea,button{font-size:13px;font-family:inherit}select,textarea{width:100%;padding:8px;box-sizing:border-box;border:1px solid #dadce0;border-radius:6px;margin-bottom:10px}textarea{height:160px;resize:vertical}button{background:#1a73e8;color:#fff;border:0;padding:8px 18px;border-radius:6px;cursor:pointer}label{font-size:12px;color:#5f6368;display:block;margin-bottom:4px}</style>' +
    '<label>カテゴリ</label>' +
    '<select id="c"><option>機能追加</option><option>UI改善</option><option>バグ報告</option><option>その他</option></select>' +
    '<label>こうしたら良いと思うこと</label>' +
    '<textarea id="t" placeholder="自由に記入"></textarea>' +
    '<button onclick="google.script.run.withSuccessHandler(()=>{alert(\'保存しました\');google.script.host.close()}).submitIdea(document.getElementById(\'c\').value, document.getElementById(\'t\').value)">送信</button>'
  ).setWidth(440).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, '改善アイデアを送る');
}

function submitIdea(category, content) {
  if (!content || !content.trim()) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_IDEAS);
  if (!sheet) sheet = setupIdeasSheet_(ss);
  // 非表示なら一時的に表示してから書き込み (省略可)
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
  sheet.appendRow([ts, category, content]);
}

// ===== ヘルプ表示 =====
function showHelp() {
  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:-apple-system,Segoe UI,sans-serif;margin:0;padding:18px;line-height:1.6;font-size:13px}h2{font-size:15px;margin:0 0 8px}h3{font-size:13px;margin:14px 0 6px}ul{padding-left:20px}code{background:#f1f3f4;padding:1px 6px;border-radius:4px}</style>' +
    '<h2>task-board-sheets 使い方</h2>' +
    '<h3>1. タスク一覧</h3><ul>' +
    '<li>メニュー「タスク管理」>「タスク追加」で先頭行に新規行が挿入されます</li>' +
    '<li>状況列はプルダウンで未着手/進行中/依頼中/完了を選ぶと自動で色分けされます</li>' +
    '<li>期日セルをダブルクリックでカレンダー(日付ピッカー)が開きます</li>' +
    '</ul>' +
    '<h3>2. 会議日程</h3><ul>' +
    '<li>設定したGoogleカレンダーの予定を一覧表示するシートです</li>' +
    '<li>毎朝6時に自動更新。手動で更新したい時はメニュー「会議日程を今すぐ更新」</li>' +
    '<li>取得期間: (今日 - 過去日数) 〜 (今日 + 先読み日数)。設定シートで調整可能</li>' +
    '<li>今日以降の予定は薄い青背景+太字でハイライト</li>' +
    '</ul>' +
    '<h3>3. リマインドメールタスク</h3><ul>' +
    '<li>「設定」シートのカレンダーIDを確認(デフォルトは <code>primary</code>=メインカレンダー)</li>' +
    '<li>カレンダーに「SMTG」を含む予定があると、その前営業日に自動でリマインドが入ります</li>' +
    '<li>毎朝6時に自動スキャン。手動で実行したい時はメニューから「SMTG リマインドを今すぐ作成」</li>' +
    '<li>同じ予定に対して既存リマインドがあれば重複生成されません</li>' +
    '</ul>' +
    '<h3>4. 企業リスト</h3><ul><li>ユニットはプルダウン(SP/CM/AI/BO/秘書/HR/DW)、議事録URLはリンクで開きます</li></ul>' +
    '<h3>5. 日報</h3><ul><li>「日報フォーマットを表示」でテンプレートをコピーできます。日付は自動挿入されます</li></ul>' +
    '<h3>6. URL</h3><ul><li>よく使うURLを集めるシートです。タイトル/URL/カテゴリ/メモを記入</li></ul>' +
    '<h3>7. 設定</h3><ul>' +
    '<li>カレンダーID: <code>primary</code> ならメインカレンダー。別カレンダーは Google Calendar の「設定と共有 → カレンダーの統合」で取得</li>' +
    '<li>SMTGキーワード: 大小無視・部分一致でマッチ</li>' +
    '<li>先読み日数: 何日先まで予定をスキャンするか</li>' +
    '<li>過去日数: 会議日程シートに何日前の予定まで残すか</li>' +
    '<li>日報テンプレート: 日報の雛形。{DATE} が今日の日付に置き換わる</li>' +
    '</ul>'
  ).setWidth(560).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'task-board-sheets の使い方');
}

// ===== 自動トリガー設定 =====
function installDailyTriggers_() {
  // 既存のトリガーを削除して重複を防ぐ
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'scanSmtgReminders' || fn === 'syncCalendarEvents') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // 06:00 JST: 会議日程の同期
  ScriptApp.newTrigger('syncCalendarEvents')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
  // 06:00 JST: SMTG リマインド作成 (Google が同時刻トリガーを微調整して順次実行)
  ScriptApp.newTrigger('scanSmtgReminders')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
}

// ===== 共通ユーティリティ =====
function getOrCreateSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function styleHeader_(range) {
  range.setBackground(HEADER_BG)
       .setFontColor(HEADER_COLOR)
       .setFontWeight('bold')
       .setHorizontalAlignment('left')
       .setVerticalAlignment('middle')
       .setBorder(false, false, true, false, false, false, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  range.getSheet().setRowHeight(1, 32);
}

function setColumnWidths_(sheet, widths) {
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
}

function setListValidation_(sheet, col, list) {
  var range = sheet.getRange(2, col, sheet.getMaxRows() - 1, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

function setDateValidation_(sheet, col) {
  var range = sheet.getRange(2, col, sheet.getMaxRows() - 1, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText('セルをダブルクリックすると日付ピッカーが開きます')
    .build();
  range.setDataValidation(rule);
  range.setNumberFormat('yyyy/mm/dd');
}

function applyStatusConditionalFormat_(sheet, col) {
  var rangeA1 = sheet.getRange(2, col, sheet.getMaxRows() - 1, 1).getA1Notation();
  var rules = sheet.getConditionalFormatRules();
  // 既存の同範囲ルールを除外
  rules = rules.filter(function(r) {
    var rs = r.getRanges();
    for (var i = 0; i < rs.length; i++) {
      if (rs[i].getA1Notation() === rangeA1) return false;
    }
    return true;
  });
  var range = sheet.getRange(rangeA1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('未着手').setBackground(COLOR_RED).setFontColor(COLOR_RED_BD).setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('進行中').setBackground(COLOR_YELLOW).setFontColor(COLOR_YELLOW_BD).setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('依頼中').setBackground(COLOR_GRAY).setFontColor(COLOR_GRAY_BD).setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('完了').setBackground(COLOR_GREEN).setFontColor(COLOR_GREEN_BD).setRanges([range]).build());
  sheet.setConditionalFormatRules(rules);
}

function removeRulesForRange_(sheet, range) {
  var rangeA1 = range.getA1Notation();
  var rules = sheet.getConditionalFormatRules();
  rules = rules.filter(function(r) {
    var rs = r.getRanges();
    for (var i = 0; i < rs.length; i++) {
      if (rs[i].getA1Notation() === rangeA1) return false;
    }
    return true;
  });
  sheet.setConditionalFormatRules(rules);
}

function applyTodayHighlight_(sheet, dateCol) {
  // 開始日時が今日以降の行を太字+薄い青背景でハイライト
  var range = sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns());
  var col = dateCol; // 開始日時列 (1-indexed)
  // 「セルの値の最初の10文字を yyyy/MM/dd 形式の今日と比較」する custom formula
  var formula = '=AND($A2<>"", LEFT($A2,10) >= TEXT(TODAY(),"YYYY/MM/DD"))';
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground('#e8f0fe')
    .setBold(true)
    .setRanges([range])
    .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

function ensureMinRows_(sheet, min) {
  var max = sheet.getMaxRows();
  if (max < min) sheet.insertRowsAfter(max, min - max);
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
                  .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
