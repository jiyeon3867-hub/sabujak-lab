/**
 * ====================================================================
 * [로기로그] 구글 스프레드시트 연동 Apps Script (Code.gs)
 * ====================================================================
 * 
 * [설정 및 배포 방법]
 * 1. 구글 스프레드시트 생성 (예: '로기로그_견적문의_접수대장')
 * 2. 상단 메뉴 [확장 프로그램] -> [Apps Script] 클릭
 * 3. 기존 코드를 모두 지우고 아래 코드를 그대로 붙여넣은 후 저장(Ctrl+S)
 * 4. 오른쪽 상단 [배포] -> [새 배포] 클릭
 * 5. 유형 선택: [웹 앱 (Web app)]
 *    - 설명: 로기로그 견적 문의 연동 API
 *    - 다음 사용자로 실행: [나 (내 계정)]
 *    - 액세스 권한이 있는 사용자: [모든 사용자 (Anyone)]  <-- 필수!
 * 6. [배포] 클릭 후 승인 절차(고급 -> 안전하지 않은 페이지로 이동 허용) 진행
 * 7. 생성된 "웹 앱 URL (exec로 끝나는 주소)"을 복사하여 웹사이트 SCRIPT_URL에 적용
 */

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // 첫 행(헤더)이 비어있다면 자동으로 열 헤더 생성
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "접수일시",
        "기업/단체명",
        "담당자 성함",
        "연락처",
        "이메일",
        "행사/희망 일정",
        "예상 인원(명)",
        "예산 범위",
        "세부 요청 사항",
        "개인정보 동의"
      ]);
      // 헤더 스타일 지정 (배경색 민트 톤, 볼드, 가운데 정렬)
      var headerRange = sheet.getRange(1, 1, 1, 10);
      headerRange.setBackground("#1BA89A");
      headerRange.setFontColor("#FFFFFF");
      headerRange.setFontWeight("bold");
      headerRange.setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    // 전송된 데이터 파싱
    var data = {};
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      data = e.parameter;
    }

    var submittedAt = data.submitted_at || Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    var company     = data.company || "";
    var name        = data.name || "";
    var phone       = data.phone || "";
    var email       = data.email || "";
    var eventDate   = data.event_date || "";
    var headcount   = data.headcount || "";
    var budget      = data.budget || "";
    var message     = data.message || "";
    var privacy     = data.privacy_agree || "동의";

    // 시트의 마지막 행에 데이터 추가
    sheet.appendRow([
      submittedAt,
      company,
      name,
      phone,
      email,
      eventDate,
      headcount,
      budget,
      message,
      privacy
    ]);

    // (선택 사항) 새 문의 접수 시 관리자 이메일 알림 받기
    /*
    var adminEmail = "your_email@gmail.com"; // 로기로그 관리자 이메일
    MailApp.sendEmail({
      to: adminEmail,
      subject: "[로기로그] 새로운 단체·기업 견적 문의가 접수되었습니다: " + company,
      body: "새로운 견적 문의 내역\n\n" +
            "- 기업/단체명: " + company + "\n" +
            "- 담당자: " + name + " (" + phone + " / " + email + ")\n" +
            "- 희망 일정: " + eventDate + "\n" +
            "- 예상 인원: " + headcount + "명\n" +
            "- 예산 범위: " + budget + "\n" +
            "- 세부 요청 사항:\n" + message + "\n\n" +
            "스프레드시트에서 전체 내역을 확인하세요."
    });
    */

    return ContentService.createTextOutput(JSON.stringify({
      result: "success",
      message: "정상적으로 등록되었습니다."
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// GET 요청 테스트용
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "ok",
    message: "로기로그 단체/기업 견적 문의 수신 API 정상 작동 중"
  })).setMimeType(ContentService.MimeType.JSON);
}
