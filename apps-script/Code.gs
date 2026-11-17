// 구글 시트 접수용 Apps Script (시트 > 확장 프로그램 > Apps Script 에 붙여넣기)
// 배포: 배포 > 배포 관리 > 연필 아이콘 > 버전: 새 버전 > 배포 (웹 앱 URL이 그대로 유지됨)
//   실행 계정: 나 / 액세스 권한: 모든 사용자

const NAMES_TAB = "시트1"; // A열에 이름. 행 번호 = 접수 페이지 경로 (/{번호})
const APPLY_TAB = "시트3"; // 기본 접수 탭 (루트 페이지 등 PAGE_TABS 에 없는 경로는 여기에 저장)
// 경로 이름 = 시트 탭 이름. /durepost 로 들어온 접수는 'durepost' 탭에 저장되고(없으면 자동 생성),
// /durepost/admin 에서 그 탭 하나만 보인다. 경로를 늘리려면 여기에 탭 이름을 추가한다.
const PAGE_TABS = ["durepost"];
// admin 페이지는 비밀번호가 필요하다:
//   Apps Script > 프로젝트 설정 > 스크립트 속성 > VIEW_KEY 를 추가하고 비밀번호를 넣는다.
const HEADERS = [
  "접수일시",
  "영업자",
  "이름",
  "휴대전화",
  "지역",
  "인터뷰 방식",
  "희망 유형",
  "개인정보 동의",
  "유입 페이지",
];
const TEL_COL = HEADERS.indexOf("휴대전화") + 1;

function doPost(e) {
  const p = e.parameter || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const ref = /^\d+$/.test(p.ref || "") ? Number(p.ref) : "";
    let listed = "";
    if (ref) {
      const names = ss.getSheetByName(NAMES_TAB) || ss.getSheets()[0];
      if (ref <= names.getLastRow())
        listed = String(names.getRange(ref, 1).getValue());
    }

    // PAGE_TABS 에 있는 탭만 허용 (임의 입력으로 탭이 생기거나 시트1이 오염되지 않게 함)
    const tabName = PAGE_TABS.indexOf(p.tab) >= 0 ? p.tab : APPLY_TAB;
    let sh = ss.getSheetByName(tabName);
    if (!sh) sh = ss.insertSheet(tabName);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
    }

    const row = sh.getLastRow() + 1;
    // 전화번호가 숫자로 변환돼 앞의 0이 사라지지 않도록 텍스트 서식으로 저장
    sh.getRange(row, TEL_COL).setNumberFormat("@");
    sh.getRange(row, 1, 1, HEADERS.length).setValues([
      [
        new Date(),
        listed,
        p.name || "",
        p.tel || "",
        p.sido || "",
        p.how || "",
        p.type || "",
        p.agree || "",
        p.page || "",
      ],
    ]);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput("ok");
}

// 접수 목록 조회: ?action=list&tab={탭이름}&key={비밀번호} → PAGE_TABS 의 탭 하나만 JSON 으로 반환
function doGet(e) {
  const p = (e && e.parameter) || {};
  const key = PropertiesService.getScriptProperties().getProperty("VIEW_KEY");
  if (!key || p.key !== key) return json_({ ok: false, error: "auth" });
  if (p.action !== "list" || PAGE_TABS.indexOf(p.tab) < 0)
    return json_({ ok: false, error: "tab" });
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(p.tab);
  if (!sh) return json_({ ok: false, error: "tab" });
  return json_({
    ok: true,
    tab: p.tab,
    rows: sh.getLastRow() ? sh.getDataRange().getDisplayValues() : [],
  });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// 편집기 테스트용: 함수 선택에서 testPost 를 골라 ▶ 실행 (doPost 는 직접 실행 불가)
function testPost() {
  doPost({
    parameter: {
      name: "테스트",
      tel: "01012345678",
      sido: "서울",
      how: "온라인",
      type: "테스트",
      agree: "Y",
      page: "test",
    },
  });
}
