/**
 * STREAMING_CHUNK: Setting up custom menu on Spreadsheet open...
 * 스프레드시트가 열릴 때 상단에 커스텀 메뉴를 생성합니다.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📝 설문지 관리')
    .addItem('1. 시트 템플릿 및 기본 문항 생성', 'setupSheets')
    .addItem('2. 구글폼 동기화 (생성 및 업데이트)', 'syncForm')
    .addToUi();
}

/**
 * STREAMING_CHUNK: Creating and initializing settings & questions sheets...
 * 설문 제어에 필요한 '설정' 시트와 '문항관리' 시트를 자동으로 생성하고 기본 데이터를 채웁니다.
 */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 설정 시트 생성 또는 가져오기
  var settingsSheet = ss.getSheetByName('설정');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('설정');
  }
  settingsSheet.clear();
  
  var settingsData = [
    ["설정 항목", "값"],
    ["설문지 ID", ""], // 최초 생성 시 자동으로 채워집니다.
    ["설문지 제목", "펫장례문화 구독서비스 수요 및 인식 조사"],
    ["설문지 설명", "안녕하세요.\n본 설문은 반려동물 장례 문화의 변화와 새로운 구독서비스(상조 서비스 등) 도입에 대한 반려인 여러분의 인식과 선호도를 파악하기 위해 마련되었습니다.\n바쁘시더라도 소중한 의견을 나누어 주시면 감사하겠습니다. (소요 시간: 약 3분)"]
  ];
  settingsSheet.getRange(1, 1, settingsData.length, 2).setValues(settingsData);
  settingsSheet.getRange("A1:B1").setBackground("#E2EFDA").setFontWeight("bold");
  settingsSheet.autoResizeColumns(1, 2);

  // 2. 문항관리 시트 생성 또는 가져오기
  var qSheet = ss.getSheetByName('문항관리');
  if (!qSheet) {
    qSheet = ss.insertSheet('문항관리');
  }
  qSheet.clear();

  // 헤더 설정
  var headers = ["ID", "구역", "유형", "질문 제목", "도움말", "필수여부(Y/N)", "기타의견(Y/N)", "선택지 (엔터/줄바꿈으로 구분)"];
  qSheet.appendRow(headers);

  // 기본 11개 문항 데이터 정의
  var defaultQuestions = [
    ["Q0", "1", "객관식", "시작하기 전 확인: 현재 반려동물(강아지, 고양이 등)을 키우고 계십니까?", "", "Y", "N", "예, 있습니다. (설문 계속 진행)\n아니오, 없습니다. (설문 종료)"],
    ["Q1", "2", "객관식", "Q1. 귀하는 반려동물이 무지개다리를 건넜을 때(사후)의 대처 방안이나 장례 절차에 대해 생각해 보신 적이 있습니까?", "", "Y", "N", "구체적인 계획과 방법을 알아본 적이 있다.\n막연하게 생각만 해보았다.\n아직 한 번도 생각해 본 적이 없다."],
    ["Q2", "2", "체크박스", "Q2. 반려동물 장례 서비스(화장, 장례식 등)를 이용할 때 가장 염려되거나 부담스러운 부분은 무엇입니까? (중복 선택 가능)", "", "Y", "Y", "갑작스러운 이별로 인한 정신적 충격(정서적 대처 어려움)\n일시에 발생하는 장례 비용에 대한 경제적 부담\n믿을 만한 합법적 장례 업체 정보를 찾는 것의 어려움\n장례 절차 및 행정 처리(사망신고 등)에 대한 무지"],
    ["Q3", "2", "객관식", "Q3. 매달 일정 금액을 지불하고, 추후 반려동물 장례 시 필요한 비용과 서비스를 보장받는 '펫장례 구독(상조) 서비스'를 들어보신 적이 있습니까?", "", "Y", "N", "잘 알고 있으며, 필요성을 느낀다.\n들어본 적은 있으나 구체적인 내용은 모른다.\n이번에 처음 알게 되었다."],
    ["Q4", "2", "객관식", "Q4. 만약 펫장례 구독서비스를 이용한다면, '장례(화장/수목장 등)' 외에 가장 포함되었으면 하는 연계 서비스는 무엇입니까?", "", "Y", "N", "펫로스 증후군 케어 (심리 상담, 치유 프로그램 등)\n노령견/노령묘를 위한 사전 헬스케어 상담 및 건강 검진 연계\n메모리얼 스톤, 유골함, 초상화 등 추모 굿즈 제작 지원\n정기 펫 스튜디오 사진 촬영 (영정사진 또는 추억 소장용)\n유산 상속 및 법률/행정 절차 대행 지원"],
    ["Q5", "2", "객관식", "Q5. 펫장례 구독서비스 선택 시 가장 중요하게 고려할 기준은 무엇입니까?", "", "Y", "N", "서비스 제공 업체의 신뢰도 및 규모\n월 구독료 및 총 납입 금액의 가성비\n장례 시설의 위치 및 접근성\n장례 지도사의 전문성과 친절도\n중도 해지 시 환불 규정 및 안정성"],
    ["Q6", "2", "객관식", "Q6. 반려동물의 장례를 미리 준비하기 위해 지불할 수 있는 '적정 월 구독료'는 얼마라고 생각하십니까?", "", "Y", "N", "5,000원 미만\n5,000원 이상 ~ 10,000원 미만\n10,000원 이상 ~ 20,000원 미만\n20,000원 이상 ~ 30,000원 미만\n30,000원 이상"],
    ["Q7", "2", "객관식", "Q7. 위에서 선택하신 금액대나 혜택이 만족스럽다면, 해당 펫장례 구독서비스를 이용(가입)하실 의향이 있으십니까?", "", "Y", "N", "매우 있다 (바로 가입 고려)\n약간 있다 (조건이 맞으면 가입 고려)\n보통이다 (잘 모르겠다)\n별로 없다 (필요할 때 일시불로 결제하겠다)\n전혀 없다 (장례 서비스 자체를 이용할 생각이 없다)"],
    ["Q8", "2", "객관식", "Q8. 현재 함께 거주하고 계신 반려동물의 종류는 무엇입니까? (다수일 경우 주 반려동물 기준)", "", "Y", "N", "반려견 (강아지)\n반려묘 (고양이)\n소동물/특수동물 (햄스터, 새, 파충류 등)"],
    ["Q9", "2", "객관식", "Q9. 현재 키우고 계신 반려동물의 나이는 어떻게 됩니까? (다수일 경우 가장 나이가 많은 아이 기준)", "", "Y", "N", "1세 ~ 3세 (성장기/청년기)\n4세 ~ 7세 (성숙기)\n8세 ~ 12세 (노령기 진입)\n13세 이상 (초고령기)"],
    ["Q10", "2", "객관식", "Q10. 귀하의 연령대는 어떻게 되십니까?", "", "Y", "N", "20대 이하\n30대\n40대\n50대\n60대 이상"]
  ];

  qSheet.getRange(2, 1, defaultQuestions.length, 8).setValues(defaultQuestions);
  qSheet.getRange("A1:H1").setBackground("#DDEBF7").setFontWeight("bold");
  qSheet.autoResizeColumns(1, 8);

  ss.toast('설정 및 문항관리 시트 템플릿 생성이 성공적으로 완료되었습니다!', '알림');
}

/**
 * STREAMING_CHUNK: Reading configurations and loading Google Form...
 * 새롭게 생성되어 폼연동은 되나 한번 생성된 폼에 업데이트가 안되고 있음
 */
function syncForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('설정');
  var qSheet = ss.getSheetByName('문항관리');

  if (!settingsSheet || !qSheet) {
    ss.toast('먼저 [1. 시트 템플릿 및 기본 문항 생성]을 실행해 주세요.', '오류');
    return;
  }

  //설정 시트에서 값 읽기
  var title = settingsSheet.getRange("B3").getValue();
  var desc = settingsSheet.getRange("B4").getValue();
  var formId = settingsSheet.getRange("B2").getValue();

  var form;
  
  // 기존 폼이 존재하는지 ID로 확인 및 로드
  if (formId) {
    try {
      form = FormApp.openById(formId);
      form.setTitle(title);
      form.setDescription(desc);
    } catch (e) {
      // ID가 유효하지 않은 경우 새롭게 생성
      form = FormApp.create(title);
      form.setDescription(desc);
      settingsSheet.getRange("B2").setValue(form.getId());
    }
  } else {
    // ID가 아예 비어있는 경우 최초 생성
    form = FormApp.create(title);
    form.setDescription(desc);
    settingsSheet.getRange("B2").setValue(form.getId());
  }

  /* STREAMING_CHUNK: Clearing existing form items for clean overwrite... */
  // 기존 문항 전부 삭제 (덮어쓰기를 위함)
  var items = form.getItems();
  for (var i = items.length - 1; i >= 0; i--) {
    form.deleteItem(items[i]);
  }

  /* STREAMING_CHUNK: Parsing question rows and organizing structure... */
  // '문항관리' 시트에서 데이터 읽어오기
  var lastRow = qSheet.getLastRow();
  var qData = qSheet.getRange(2, 1, lastRow - 1, 8).getValues();

  // 구역에 맞춰 폼 구조 빌드하기
  // 1구역: 선별 구역 (시작 페이지)
  // 2구역: 메인 설문 (반려인 대상)
  // 3구역: 종료 안내 (비반려인 대상)
  
  var q0Item = null;
  var q0Choices = [];
  
  // 1. 선별 질문 추가 (우선 생성하고 나중에 네비게이션 지정)
  var q0Row = qData.filter(function(row) { return row[1] == 1; })[0];
  if (q0Row) {
    q0Item = form.addMultipleChoiceItem();
    q0Item.setTitle(q0Row[3]);
    if (q0Row[4]) q0Item.setHelpText(q0Row[4]);
    q0Item.setRequired(q0Row[5] === 'Y');
    q0Choices = q0Row[7].split('\n');
  }

  /* STREAMING_CHUNK: Adding Section 2 (Main Survey) and its questions... */
  // 2. 2단계 메인 설문 섹션 추가
  var pageTwo = form.addPageBreakItem()
      .setTitle('펫장례문화 구독서비스 수요 및 인식 조사')
      .setHelpText('반려 생활을 하시는 분들을 대상으로 진행하는 메인 설문 영역입니다.');

  // 2구역 질문들 필터링하여 추가
  var s2Questions = qData.filter(function(row) { return row[1] == 2; });
  s2Questions.forEach(function(row) {
    var type = row[2];
    var qTitle = row[3];
    var qHelp = row[4];
    var req = (row[5] === 'Y');
    var hasOther = (row[6] === 'Y');
    var rawChoices = row[7].split('\n').filter(function(c) { return c.trim() !== ""; });

    if (type === '객관식') {
      var mcItem = form.addMultipleChoiceItem().setTitle(qTitle).setRequired(req);
      if (qHelp) mcItem.setHelpText(qHelp);
      var choicesArr = rawChoices.map(function(choice) { return mcItem.createChoice(choice); });
      mcItem.setChoices(choicesArr);
      if (hasOther) mcItem.showOtherOption(true);
    } else if (type === '체크박스') {
      var cbItem = form.addCheckboxItem().setTitle(qTitle).setRequired(req);
      if (qHelp) cbItem.setHelpText(qHelp);
      var choicesArr = rawChoices.map(function(choice) { return cbItem.createChoice(choice); });
      cbItem.setChoices(choicesArr);
      if (hasOther) cbItem.showOtherOption(true);
    }
  });

  /* STREAMING_CHUNK: Adding Section 3 (Exit Page) and linking branch routing... */
  // 3. 3단계 비반려인용 감사 및 종료 섹션 추가
  var pageThree = form.addPageBreakItem()
      .setTitle('설문에 관심을 가져주셔서 감사합니다.')
      .setHelpText('본 설문은 현재 반려동물을 양육하고 계신 분들을 대상으로 진행하고 있습니다.\n소중한 시간을 내어 응답해 주셔서 대단히 고맙습니다.');

  // 각 섹션 종료 시 행동 지정 (바로 설문 제출 처리)
  pageTwo.setGoToPage(FormApp.PageNavigationType.SUBMIT);
  pageThree.setGoToPage(FormApp.PageNavigationType.SUBMIT);

  // 4. 선별 질문(Q0)에 페이지 분기 경로 연결
  if (q0Item && q0Choices.length >= 2) {
    q0Item.setChoices([
      q0Item.createChoice(q0Choices[0], pageTwo),
      q0Item.createChoice(q0Choices[1], pageThree)
    ]);
  }

  ss.toast('구글 폼 동기화 및 덮어쓰기가 완료되었습니다!', '성공');
}

