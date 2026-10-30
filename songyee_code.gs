/**
 * @OnlyCurrentDoc
 */
function sendDormantEmails() {
  // 1. [테스트 모드 설정] 
  // true: 내 이메일로만 발송(안전한 테스트) / false: 실제 고객 이메일로 발송
  var TEST_MODE = true; 
  
  // 2. 현재 열려 있는 스프레드시트에서 '고객명단' 시트를 가져옵니다.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('고객명단');
  
  // 만약 '고객명단' 시트가 없다면 에러를 내고 중단합니다.
  if (!sheet) {
    Logger.log('오류: "고객명단" 시트를 찾을 수 없습니다. 시트 이름을 확인해주세요!');
    return;
  }
  
  // 3. 시트에서 데이터가 있는 마지막 행 번호를 가져옵니다.
  var lastRow = sheet.getLastRow();
  
  // 데이터가 제목행(1행)만 있고 내용이 없다면 중단합니다.
  if (lastRow < 2) {
    Logger.log('처리할 데이터가 없습니다.');
    return;
  }
  
  // 4. 2행부터 마지막 행까지, A열(1)부터 G열(7)까지의 데이터를 한 번에 가져옵니다.
  var dataRange = sheet.getRange(2, 1, lastRow - 1, 7);
  var values = dataRange.getValues();
  
  // 실행자의 이메일 주소 가져오기 (테스트 모드용)
  var myEmail = Session.getActiveUser().getEmail();
  
  // 총 몇 행이 처리되었는지 세기 위한 변수 (0에서 시작)
  var processedCount = 0;
  
  // 5. 가져온 데이터를 한 줄씩 반복(루프)하며 조건에 맞는지를 확인합니다.
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var actualRowNumber = i + 2; // 실제 스프레드시트의 행 번호 (2행부터 시작)
    
    var name = row[0];       // A열: 이름 (인덱스 0)
    var email = row[1];      // B열: 이메일 (인덱스 1)
    var grade = String(row[2]).trim(); // C열: 등급 (인덱스 2 / 보이지 않는 공백 제거)
    var sentStatus = row[6]; // G열: 발송여부 (인덱스 6)
    
    // 6. 조건 확인: C열이 '휴면'이고, G열(발송여부)이 비어 있는 경우
    if (grade === '휴면' && (!sentStatus || sentStatus === '')) {
      
      // 테스트 모드인지에 따라 보낼 이메일 주소 결정
      var recipient = TEST_MODE ? myEmail : email;
      
      // 메일 제목과 본문 작성 (A열 이름 포함)
      var subject = '[안내] 휴면 계정 전환 관련 안내 말씀 드립니다.';
      var body = name + ' 고객님 안녕하세요.\n\n' +
                 '오랫동안 서비스를 이용하지 않아 계정이 휴면 상태로 전환되었습니다.\n' +
                 '서비스를 다시 이용하시려면 로그인해 주세요.\n\n감사합니다.';
      
      // 테스트 모드일 때 메일 본문에 원래 수신자 정보 표시
      if (TEST_MODE) {
        body = '[테스트 모드 - 원래 수신자: ' + email + ']\n\n--------------------\n\n' + body;
      }
      
      try {
        // 7. 메일 발송하기
        MailApp.sendEmail(recipient, subject, body);
        
        // 8. G열(7번째 열)에 오늘 날짜 기록하기
        var today = new Date();
        sheet.getRange(actualRowNumber, 7).setValue(today);
        
        // 메일 발송이 성공할 때마다 처리된 행 수를 1씩 증가시킵니다.
        processedCount++;
        
        Logger.log(actualRowNumber + '행: ' + name + '(' + email + ')님에게 메일 발송 완료');
        
      } catch (error) {
        // 메일 발송 중 에러가 발생하면 로그에 남깁니다.
        Logger.log(actualRowNumber + '행 발송 실패: ' + error.toString());
      }
    }
  }
  
  // 9. 모든 반복이 끝난 후, 총 몇 행이 처리되었는지 로그로 출력합니다.
  Logger.log('========================================');
  Logger.log('작업 완료! 총 ' + processedCount + '개의 행에 메일이 발송되었습니다.');
  Logger.log('========================================');
}
