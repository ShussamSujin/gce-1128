/**
 * GCE 1128 합격 자랑판 저장소
 * - 독립형 Apps Script 프로젝트에서 아래 시트 ID의 시트에 기록해요.
 * - "자랑판" 탭의 1행: 시각 | 이름 | 학교 | 레벨 | 링크 | 한마디
 */
const SPREADSHEET_ID = "1tPBbnue-O7vx1saG8uuc_Rmzgejy5hormROPDsb98Bw";
const SHEET_NAME = "자랑판";

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["시각", "이름", "학교", "레벨", "링크", "한마디"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 목록 읽기: GET ?action=list
function doGet(e) {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return json_({ ok: true, items: [] });
  const rows = sh.getRange(2, 1, last - 1, 6).getValues();
  const items = rows
    .filter(r => r[1])
    .map(r => ({
      time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      name: String(r[1]), school: String(r[2]), level: String(r[3]),
      link: String(r[4]), note: String(r[5]),
    }))
    .reverse();
  return json_({ ok: true, items });
}

// 등록: POST(본문 JSON) { name, school, level, link, note }
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || "{}");
    const name = String(data.name || "").trim().slice(0, 20);
    const school = String(data.school || "").trim().slice(0, 30);
    const level = String(data.level || "").trim();
    const link = String(data.link || "").trim().slice(0, 300);
    const note = String(data.note || "").trim().slice(0, 60);

    if (!name || !school) return json_({ ok: false, error: "이름과 학교를 적어 주세요." });
    if (!["Level 1", "Level 2"].includes(level)) return json_({ ok: false, error: "레벨을 골라 주세요." });
    if (link && !/^https?:\/\//.test(link)) return json_({ ok: false, error: "링크는 http로 시작해야 해요." });

    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      const sh = getSheet_();
      // 같은 사람·같은 레벨이 또 올리면 새 줄 대신 기존 줄을 갱신해요.
      const last = sh.getLastRow();
      if (last >= 2) {
        const rows = sh.getRange(2, 1, last - 1, 6).getValues();
        for (let i = 0; i < rows.length; i++) {
          if (rows[i][1] === name && rows[i][2] === school && rows[i][3] === level) {
            sh.getRange(i + 2, 1, 1, 6).setValues([[new Date(), name, school, level, link, note]]);
            return json_({ ok: true, updated: true });
          }
        }
      }
      sh.appendRow([new Date(), name, school, level, link, note]);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
