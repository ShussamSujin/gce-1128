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

  // 관리용: 실습링크 탭을 아래 seedLinks() 목록으로 다시 써요 (?action=seedlinks&key=gce1128-setup)
  if (action === "seedlinks") {
    if (String(p.key) !== "gce1128-setup") return json_({ ok: false, error: "key" });
    return json_({ ok: true, result: seedLinks() });
  }


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
    // 7번째 열 "짝"에 짝꿍 아이디를 적으면 그대로, 비어 있으면 번호 순서대로 둘씩(홀수면 마지막 셋) 묶어요.
    const sh = sheet_(SHEET_ACCOUNTS, ["지역", "학교", "이름", "아이디", "비밀번호", "구분", "짝"]);
    const all = rows_(sh, 7).filter(r => String(r[3]).trim());
    const toObj = r => ({ region: String(r[0]), school: String(r[1]), name: String(r[2]), id: String(r[3]).trim(), pw: String(r[4]), group: String(r[5]) });
    const partnersOf = (idx) => {
      const manual = String(all[idx][6] || "").split(/[,\s]+/).filter(Boolean);
      if (manual.length) return all.filter(r => manual.includes(String(r[3]).trim())).map(toObj);
      const n = all.length, pairIdx = Math.floor(idx / 2);
      const lastTrio = (n % 2 === 1) && idx >= n - 3;
      const mates = lastTrio ? [n - 3, n - 2, n - 1] : [pairIdx * 2, pairIdx * 2 + 1];
      return mates.filter(i => i !== idx && i < n).map(i => toObj(all[i]));
    };
    const hits = [];
    all.forEach((r, i) => { if (String(r[2]).replace(/\s+/g, "") === name) hits.push(Object.assign(toObj(r), { partners: partnersOf(i).map(x => ({ name: x.name, school: x.school, id: x.id })) })); });
    if (!hits.length) return json_({ ok: false, error: "명단에 없는 이름이에요. 강사에게 알려 주세요." });
    return json_({ ok: true, items: hits });
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
// 사람이 늘면 "실습계정" 탭 맨 아래 행에 지역 | 학교 | 이름 | 아이디 | 비밀번호 | 구분 | 짝 순으로 추가하면 돼요.
// "짝" 열은 비워 두면 번호 순서대로 둘씩 자동 배정, 아이디를 적으면 그 사람과 짝이에요.
function setup() {
  // 시트 탭과 머리글만 만들어요. 데이터는 시트에 직접 입력해요.
  sheet_(SHEET_ACCOUNTS, ["지역", "학교", "이름", "아이디", "비밀번호", "구분", "짝"]);
  sheet_(SHEET_LINKS, ["탭", "제목", "주소", "메모", "실습"]);
  return "tabs ready";
}

// 실습링크 탭 기본 목록 (슬라이드 순서). 시트에서 직접 고쳐도 되고, seedlinks로 이 목록으로 되돌릴 수도 있어요.
function seedLinks() {
  const lk = sheet_(SHEET_LINKS, ["탭", "제목", "주소", "메모", "실습"]);
  const L = [
    ["바로가기","새 문서","https://docs.new/","",""],
    ["바로가기","새 슬라이드","https://slides.new/","",""],
    ["바로가기","새 스프레드시트","https://sheets.new/","",""],
    ["바로가기","새 설문지","https://forms.new/","",""],
    ["바로가기","새 사이트","https://sites.new/","",""],
    ["바로가기","구글 드라이브","https://drive.google.com/drive/my-drive","",""],
    ["바로가기","구글 클래스룸","https://classroom.google.com/h","",""],
    ["바로가기","Gemini","https://gemini.google.com/","",""],
    ["바로가기","GCE 시험 등록·응시","https://educertifications.google/","개인 Gmail로 로그인!!",""],
    ["PART 1 · 레벨 1","구글 드라이브 열기","https://drive.google.com/drive/my-drive","폴더 만들기 · 색상 · 공유 권한 3가지","실습 1"],
    ["PART 1 · 레벨 1","전교 시간표 (참고 · 예시 시트)","https://docs.google.com/spreadsheets/d/1zXA9R7yyQSOcnM3cE4b3O4vOHXLzeXoh4GclD3XfJQc/edit?usp=sharing","구글로 바뀐 학교 예시",""],
    ["PART 1 · 레벨 1","부장회의록 (참고 · 예시 시트)","https://docs.google.com/spreadsheets/d/1N_C7yvQWdRfVQKCVD8u2tgVYDqYmO09Jm-SjXjLlsco/edit?usp=sharing","구글로 바뀐 학교 예시",""],
    ["PART 1 · 레벨 1","구글 캘린더","https://calendar.google.com/calendar/u/0/r/month","일정 만들기 · 공유 · Meet 붙이기",""],
    ["PART 1 · 레벨 1","전화로 Meet 참여 (GCE 팁)","https://support.google.com/meet/answer/9683440?hl=ko","고객센터",""],
    ["PART 1 · 레벨 1","구글 클래스룸","https://classroom.google.com/h","과제 배부 · 루브릭",""],
    ["PART 1 · 레벨 1","클래스룸 연습 세트 (GCE 팁)","https://www.youtube.com/watch?v=J11wz-3UJGc","YouTube",""],
    ["PART 1 · 레벨 1","새 설문지로 자동 채점 퀴즈","https://forms.new/","퀴즈로 만들기 · 정답·점수",""],
    ["PART 1 · 레벨 1","새 문서로 역할 지침서 쓰기","https://docs.new/","제목 스타일 · 목차 · 제안 · 버전 기록","실습 2"],
    ["PART 1 · 레벨 1","모둠 협업 슬라이드 (따라하기)","https://docs.google.com/presentation/d/1RwVZRRYHC6bmp5O2L3_h6U2lWgiXiEA37e39ntgEveM/copy","열면 사본이 만들어져요",""],
    ["PART 1 · 레벨 1","조건부 서식 활용 예시 (참고)","https://docs.google.com/spreadsheets/d/15rY0ZfqXmSFwiXrdOCEMFAPW45DAdurSQcpSNNSFiyk/edit?usp=sharing","시트 · 차트와 조건부 서식",""],
    ["PART 1 · 레벨 1","유튜브 재생목록 완성 예시","https://www.youtube.com/playlist?list=PLSGLpj8V9phQ","주소 확인 필요",""],
    ["PART 1 · 레벨 1","완성 미리 보기 · 다문화 여행사 사이트","https://sites.google.com/view/2026-1128","구글 사이트 조립과 게시",""],
    ["PART 1 · 레벨 1","구글 어스 · 80일간의 세계일주","https://earth.google.com/earth/d/1VjbeP5qdvEMz3tKSWpEUSwnMfe23BL0-?usp=sharing","사이트에 넣을 자료",""],
    ["PART 1 · 레벨 1","구글 어스 · 어서 와! 우리 마을 동작","https://earth.google.com/earth/d/1QwyOkMrIDFbRLw7N7wMqv9GdOUXbUUX4?usp=sharing","사이트에 넣을 자료",""],
    ["PART 1 · 레벨 1","구글 어스 · 다문화 여행사 TOP 10 코스","https://earth.google.com/earth/d/1X0XY18eAzrgZHqMMtgOCpObjAfz4VH7l?usp=sharing","완성 예시",""],
    ["PART 2 · 레벨 1 응시","GCE 시험 등록·응시","https://educertifications.google/","개인 Gmail로 로그인 → Level 1",""],
    ["PART 2 · 레벨 1 응시","Google Certified Educator 소개","https://support.google.com/googleforeducation/answer/16542372","고객센터",""],
    ["PART 3 · 레벨 2","구글 캘린더 · 약속 일정","https://calendar.google.com/calendar/u/0/r/month","만들기 → 약속 일정 · Meet 붙이기","실습 3"],
    ["PART 3 · 레벨 2","구글 클래스룸 · 과제·성적 가져오기·원본성","https://classroom.google.com/h","짝꿍을 학생으로 초대","실습 4"],
    ["PART 3 · 레벨 2","새 설문지 퀴즈 (과제에 붙이기)","https://forms.new/","동영상 1개 + 확인 질문 2개","실습 4"],
    ["PART 3 · 레벨 2","새 설문지 (응답 모으기)","https://forms.new/","색 · 진도(%) 질문","실습 5"],
    ["PART 3 · 레벨 2","열 통계 (GCE 팁)","https://support.google.com/docs/answer/10098582?hl=ko","고객센터","실습 5"],
    ["PART 3 · 레벨 2","피봇 테이블 (GCE 팁)","https://support.google.com/docs/answer/1272900?hl=ko","고객센터","실습 5"],
    ["PART 3 · 레벨 2","찾기 및 바꾸기 옵션 (GCE 팁)","https://support.google.com/docs/answer/62754?hl=ko","고객센터",""],
    ["PART 3 · 레벨 2","이미지 자리표시자 (GCE 팁)","https://drive.google.com/file/d/1KdVS4Gpne0Ht4v3uoCHtl4czi6kr6gyl/view","드라이브 영상",""],
    ["PART 3 · 레벨 2","댓글로 슬라이드 할당 (GCE 팁)","https://drive.google.com/file/d/18JbuxYj0JC3iD8aZPoTxjCc1jF629nvx/view","드라이브 영상",""],
    ["PART 3 · 레벨 2","새 슬라이드 (단어·정의 4장, 왕복 링크)","https://slides.new/","","실습 6"],
    ["PART 3 · 레벨 2","새 사이트 (자료 넣고 게시)","https://sites.new/","권한 · 게시 대상","실습 6"],
    ["PART 4 · 레벨 2 응시","GCE 시험 등록·응시","https://educertifications.google/","개인 Gmail로 로그인 → Level 2",""],
    ["PART 4 · 레벨 2 응시","합격 자랑판에 올리기","#board","합격하면 바로!!",""],
    ["PART 4 · 레벨 2 응시","만족도 조사","https://forms.gle/7aRaJanRhRrF6U158","마치기 전에 1분",""],
  ];
  if (lk.getLastRow() > 1) lk.getRange(2, 1, lk.getLastRow() - 1, 5).clearContent();
  lk.getRange(2, 1, L.length, 5).setValues(L);
  return "links " + L.length;
}
