/**
 * 事務事業評価 事業一覧ツール（評価画面）専用 GAS+Sheets バックエンド
 *
 * 既存の事務事業優先度評価ツール（v3・10自治体）のGASとは別物。
 * 評価は 👍good／👎bad の2択で、端末ごとのランダムな識別番号（voter_id）を使い、
 * 同じ端末から同じ事業に送り直した場合は「最新の評価だけ」を集計する（評価の変更・取り消しに対応）。
 *
 * 【シート】（setupSheets() を1回実行すると自動で作られる）
 *   ratings          ： id | tool_id | voter_id | event_no | rating | submitted_at
 *                       rating は good / bad / none（none＝評価の取り消し）。行は追記のみで、集計時に最新を採用する
 *   comments         ： id | tool_id | event_no | comment | submitted_at
 *   comment_reports  ： id | comment_id | tool_id | reported_at
 *
 * 【注意】書き込みもdoGetで行う。GAS Web Appは匿名アクセス時に302リダイレクトするため、
 * ブラウザのfetch()でPOSTするとGETへ変換されボディが失われる（既存GASと同じ理由）。
 */

var RATINGS_SHEET = 'ratings';
var COMMENTS_SHEET = 'comments';
var REPORTS_SHEET = 'comment_reports';
var RATING_VALUES = { good: true, bad: true, none: true };
var MAX_ROWS_PER_REQUEST = 200;
var REPORT_HIDE_THRESHOLD = 3; // 同一コメントへの通報がこの件数に達したら一覧から除外する
var COMMENT_MAX_LENGTH = 500;

// クライアント側と同じ4パターン。直接APIを呼ばれた場合に備えてサーバー側でもチェックする
var PII_PATTERNS = [
  /0\d{1,4}-\d{1,4}-\d{3,4}/,
  /[\w.+-]+@[\w-]+\.[a-zA-Z]{2,}/,
  /https?:\/\/\S+/,
  /(私|僕|自分)の名前は|本名は/
];
function containsPii_(text) {
  for (var i = 0; i < PII_PATTERNS.length; i++) {
    if (PII_PATTERNS[i].test(text)) return true;
  }
  return false;
}

var HEADERS = {};
HEADERS[RATINGS_SHEET] = ['id', 'tool_id', 'voter_id', 'event_no', 'rating', 'submitted_at'];
HEADERS[COMMENTS_SHEET] = ['id', 'tool_id', 'event_no', 'comment', 'submitted_at'];
HEADERS[REPORTS_SHEET] = ['id', 'comment_id', 'tool_id', 'reported_at'];

// 初回に1回だけ手動で実行する（3枚のシートとヘッダー行を作る。すでにあれば何もしない）
function setupSheets() {
  Object.keys(HEADERS).forEach(function (name) { getSheet_(name); });
}

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function readRows_(name) {
  var values = getSheet_(name).getDataRange().getValues();
  return values.slice(1).filter(function (row) { return row[0] !== '' && row[0] !== null; });
}

function appendRows_(name, makeRows) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_(name);
    var lastRow = sheet.getLastRow(); // ヘッダー行を含む
    var nextId = Math.max(lastRow - 1, 0);
    var now = new Date().toISOString();
    var rows = makeRows(function () { return ++nextId; }, now);
    if (rows.length) sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
    return rows;
  } finally {
    lock.releaseLock();
  }
}

function validToolId_(s) { return typeof s === 'string' && /^[0-9a-z_]{5,60}$/.test(s); }
function validVoterId_(s) { return typeof s === 'string' && /^[0-9a-z]{8,40}$/.test(s); }

function doPost(e) {
  return jsonOut_({ ok: false, error: 'not supported: use GET' });
}

function doGet(e) {
  try {
    var p = e.parameter;
    var action = p.action;
    var toolId = p.tool_id;
    if (!validToolId_(toolId)) return jsonOut_({ ok: false, error: 'invalid tool_id' });

    // ── 評価の送信：rows = [[event_no, "good"|"bad"|"none"], ...] ──────────
    if (action === 'rate') {
      var voterId = p.voter_id;
      if (!validVoterId_(voterId)) return jsonOut_({ ok: false, error: 'invalid voter_id' });
      var pairs = JSON.parse(p.rows || '[]');
      if (!Array.isArray(pairs) || !pairs.length) return jsonOut_({ ok: false, error: 'no rows' });
      if (pairs.length > MAX_ROWS_PER_REQUEST) return jsonOut_({ ok: false, error: 'too many rows' });
      for (var i = 0; i < pairs.length; i++) {
        var no = Number(pairs[i][0]);
        if (!(no >= 1 && no === Math.floor(no)) || !RATING_VALUES[pairs[i][1]]) return jsonOut_({ ok: false, error: 'invalid row' });
      }
      var inserted = appendRows_(RATINGS_SHEET, function (nextId, now) {
        return pairs.map(function (pr) { return [nextId(), toolId, voterId, Number(pr[0]), pr[1], now]; });
      });
      return jsonOut_({ ok: true, inserted: inserted.length });
    }

    // ── 集計：端末×事業ごとに最新の評価を採用して good/bad を数える ──────────
    if (action === 'agg') {
      var latest = {}; // "voter|event_no" -> rating（行は追記順なので後勝ち）
      readRows_(RATINGS_SHEET).forEach(function (row) {
        if (String(row[1]) !== toolId) return;
        latest[row[2] + '|' + row[3]] = row[4];
      });
      var agg = {};
      Object.keys(latest).forEach(function (key) {
        var rating = latest[key];
        if (rating !== 'good' && rating !== 'bad') return;
        var no = key.split('|')[1];
        if (!agg[no]) agg[no] = { good: 0, bad: 0 };
        agg[no][rating]++;
      });
      return jsonOut_({ ok: true, agg: agg });
    }

    // ── コメント（追加のみ。修正・削除はできない） ─────────────────────
    if (action === 'comment') {
      var eventNo = Number(p.event_no);
      var comment = String(p.comment || '');
      if (!(eventNo >= 1)) return jsonOut_({ ok: false, error: 'missing event_no' });
      if (!comment.trim()) return jsonOut_({ ok: false, error: 'empty comment' });
      if (comment.length > COMMENT_MAX_LENGTH) return jsonOut_({ ok: false, error: 'comment too long' });
      if (containsPii_(comment)) return jsonOut_({ ok: false, error: 'pii detected' });
      var crow = appendRows_(COMMENTS_SHEET, function (nextId, now) {
        return [[nextId(), toolId, eventNo, comment, now]];
      });
      return jsonOut_({ ok: true, id: crow[0][0] });
    }

    if (action === 'comment_report') {
      var commentId = Number(p.comment_id);
      if (!(commentId >= 1)) return jsonOut_({ ok: false, error: 'missing comment_id' });
      appendRows_(REPORTS_SHEET, function (nextId, now) {
        return [[nextId(), commentId, toolId, now]];
      });
      return jsonOut_({ ok: true });
    }

    if (action === 'comment_list') {
      var reportCounts = {};
      readRows_(REPORTS_SHEET).forEach(function (row) {
        var cid = String(row[1]);
        reportCounts[cid] = (reportCounts[cid] || 0) + 1;
      });
      var comments = readRows_(COMMENTS_SHEET)
        .filter(function (row) { return String(row[1]) === toolId && (reportCounts[String(row[0])] || 0) < REPORT_HIDE_THRESHOLD; })
        .map(function (row) {
          return {
            id: row[0],
            event_no: String(row[2]),
            comment: row[3],
            submitted_at: row[4] instanceof Date ? row[4].toISOString() : row[4]
          };
        });
      comments.sort(function (a, b) { return String(a.submitted_at).localeCompare(String(b.submitted_at)); });
      return jsonOut_({ ok: true, comments: comments });
    }

    return jsonOut_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
