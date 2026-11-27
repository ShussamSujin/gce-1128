/**
 * GCE 1128 연수 사이트 저장소 (Apps Script 웹 앱)
 * 시트 탭: 자랑판 | 실습계정 | 실습링크
 *
 *  GET  ?action=list            → 합격 자랑판 목록
 *  GET  ?action=links           → 실습 링크 목록 (탭·제목·주소·메모·실습)
 *  GET  ?action=account&name=…  → 이름이 정확히 일치하는 실습 계정
 *  POST { name, school, level, link, note } → 자랑판 등록
 */
const SPREADSHEET_ID = "1tPBbnue-O7vx1saG8uuc_Rmzgejy5hormROPDsb98Bw";
const SHEET_BOARD = "자랑판";
const SHEET_ACCOUNTS = "실습계정";
const SHEET_LINKS = "실습링크";

function ss_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

function sheet_(name, header) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function rows_(sh, cols) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cols).getValues().filter(r => String(r[0] || r[1] || "").trim());
}

// ---------- GET ----------
function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = String(p.action || "list");


  if (action === "links") {
    const sh = sheet_(SHEET_LINKS, ["탭", "제목", "주소", "메모", "실습"]);
    const items = rows_(sh, 5).map(r => ({
      session: String(r[0]), title: String(r[1]), url: String(r[2]), note: String(r[3]), lab: String(r[4]),
    })).filter(i => i.title && i.url);
    return json_({ ok: true, items });
  }

  if (action === "account") {
    const name = String(p.name || "").replace(/\s+/g, "");
    if (!name) return json_({ ok: false, error: "이름을 입력해 주세요." });
    const sh = sheet_(SHEET_ACCOUNTS, ["지역", "학교", "이름", "아이디", "비밀번호", "구분"]);
    const hits = rows_(sh, 6).filter(r => String(r[2]).replace(/\s+/g, "") === name);
    if (!hits.length) return json_({ ok: false, error: "명단에 없는 이름이에요. 강사에게 알려 주세요." });
    const items = hits.map(r => ({ region: String(r[0]), school: String(r[1]), name: String(r[2]), id: String(r[3]), pw: String(r[4]), group: String(r[5]) }));
    return json_({ ok: true, items });
  }

  const sh = sheet_(SHEET_BOARD, ["시각", "이름", "학교", "레벨", "링크", "한마디"]);
  const items = rows_(sh, 6).filter(r => r[1]).map(r => ({
    time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
    name: String(r[1]), school: String(r[2]), level: String(r[3]), link: String(r[4]), note: String(r[5]),
  })).reverse();
  return json_({ ok: true, items });
}

// ---------- POST: 자랑판 등록 ----------
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
      const sh = sheet_(SHEET_BOARD, ["시각", "이름", "학교", "레벨", "링크", "한마디"]);
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

// ---------- 참고 ----------
// 실습계정·실습링크 탭의 내용은 시트에서 직접 관리해요(공개 저장소에는 명단을 두지 않아요).
// 사람이 늘면 "실습계정" 탭 맨 아래 행에 지역 | 학교 | 이름 | 아이디 | 비밀번호 | 구분 순으로 추가하면 돼요.
function setup() {
  // 시트 탭과 머리글만 만들어요. 데이터는 시트에 직접 입력해요.
  sheet_(SHEET_ACCOUNTS, ["지역", "학교", "이름", "아이디", "비밀번호", "구분"]);
  sheet_(SHEET_LINKS, ["탭", "제목", "주소", "메모", "실습"]);
  return "tabs ready";
}
