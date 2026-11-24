/**
 * 몽글 Pop Pop! 랭킹 저장소 (Google Apps Script)
 *
 * 이 코드를 구글 시트의 [확장 프로그램 → Apps Script]에 붙여 넣고
 * [배포 → 새 배포 → 웹 앱]으로 배포하면, 게임이 이 시트에 점수를 저장하고
 * TOP 10을 불러와요. 시트 탭 이름이 게임 버전(v1, v2)이에요.
 * 장난 기록은 시트에서 그 행을 지우면 순위표에서도 사라져요.
 */

const HEADERS = ['저장 시각', '닉네임', '점수', '라운드', '모드', '플레이어'];
const VERSIONS = ['v1', 'v2'];
const TOP_N = 10;

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'top') {
    return json({ ok: true, rows: topRows(cleanVer(p.ver), TOP_N) });
  }
  return json({ ok: true, message: '몽글 Pop Pop! 랭킹 서버가 동작 중이에요.' });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ ok: false, error: 'bad_json' });
  }
  const ver = cleanVer(body.ver);
  const name = cleanName(body.name);
  const score = Math.floor(Number(body.score));
  const round = Math.floor(Number(body.round));
  const mode = Number(body.mode) === 2 ? 2 : 1;
  const player = Number(body.player) === 2 ? 2 : 1;
  if (!isFinite(score) || score <= 0 || score > 9999990) return json({ ok: false, error: 'bad_score' });
  if (!isFinite(round) || round < 1 || round > 999) return json({ ok: false, error: 'bad_round' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheetFor(ver).appendRow([new Date(), name, score, round, mode, player]);
  } finally {
    lock.releaseLock();
  }
  return json({ ok: true });
}

function topRows(ver, n) {
  const sh = sheetFor(ver);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues()
    .filter(function (r) { return typeof r[2] === 'number' && r[1] !== ''; })
    .sort(function (a, b) { return b[2] - a[2]; })
    .slice(0, n)
    .map(function (r) {
      return {
        at: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
        name: String(r[1]).replace(/^'/, ''),
        score: r[2],
        round: Number(r[3]) || 1,
        mode: Number(r[4]) === 2 ? 2 : 1,
        player: Number(r[5]) === 2 ? 2 : 1,
      };
    });
}

function sheetFor(ver) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ver);
  if (!sh) {
    sh = ss.insertSheet(ver);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function cleanVer(v) {
  v = String(v || 'v1');
  return VERSIONS.indexOf(v) >= 0 ? v : 'v1';
}

// 닉네임: 10자까지, 수식으로 해석되지 않게 막아요.
function cleanName(s) {
  s = String(s || '').replace(/[\u0000-\u001f]/g, '').replace(/\s+/g, ' ').trim().slice(0, 10) || '익명';
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
