# 📊 Google Apps Script 기반 HTML 웹 폼 데이터 수집 프로젝트

Google Apps Script(GAS)를 활용하여 웹 페이지의 HTML `<form>` 태그로부터 입력된 데이터를 구글 스프레드시트(Google Sheets)로 실시간 전송하고 저장하는 토이 프로젝트입니다. 별도의 백엔드 서버 없이 구글 인프라를 활용하여 서버리스(Serverless) 데이터 수집 파이프라인을 구축할 수 있습니다.

---

## 🛠 기술 스택 (Tech Stack)

* **Frontend:** HTML5, CSS3, JavaScript (Fetch API / Axios)
* **Backend:** Google Apps Script (V8 Runtime)
* **Database:** Google Sheets

---

## 🚀 주요 기능 (Features)

* **서버리스 데이터 수집:** 웹 페이지 내 `<form>` 요소를 통해 제출된 데이터를 구글 시트에 실시간 적재
* **CORS 이슈 해결:** `HtmlService` 또는 `ContentService`를 활용하여 외부 도메인에서의 비동기 API 요청(`POST`) 처리
* **구글 API 확장성:** 데이터 저장과 동시에 구글 설문지 연동, 지메일(Gmail) 자동 알림 등 확장 가능

---

## 💻 설치 및 사용 방법 (Setup Instructions)

### 1. 구글 스프레드시트 및 스크립트 설정
1. 새 **구글 스프레드시트**를 생성합니다.
2. 상단 메뉴에서 **확장 프로그램** ➡️ **Apps Script**를 클릭합니다.
3. 기존 코드를 지우고 아래의 `Code.gs` 스크립트를 붙여넣습니다.

```javascript
// Code.gs
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var params = e.parameter;
    
    // 스프레드시트에 기록할 데이터 배열 매핑 (HTML form의 name 속성과 일치해야 함)
    var rowData = [
      new Date(), // 제출 시간
      params.name,
      params.email,
      params.message
    ];
    
    sheet.appendRow(rowData);
    
    return ContentService.createTextOutput(JSON.stringify({"result": "success"}))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"result": "error", "error": error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}
## 📌 커밋 메시지 컨벤션 (Commit Message Convention)

프로젝트의 일관된 히스토리 관리를 위해 아래의 말머리(Type)를 준수하여 커밋 메시지를 작성합니다.

| 말머리 (Type) | 의미 (Description) | 사용 예시 (Example) |
| :--- | :--- | :--- |
| **`Feat`** | 새로운 기능 추가 | `Feat: 로그인 구글 소셜 연동 기능 추가` |
| **`Fix`** | 버그 수정 | `Fix: 회원가입 시 이메일 중복 체크 오류 수정` |
| **`Docs`** | 문서 수정 (README.md, 주석 등) | `Docs: README.md 설치 방법 업데이트` |
| **`Style`** | 코드 포맷팅, 세미콜론 누락 등 (코드 변경 없는 포맷팅 수정) | `Style: 인덴트 오류 수정 및 세미콜론 추가` |
| **`Refactor`** | 코드 리팩토링 (기능 추가/버그 수정 없는 구조 개선) | `Refactor: 중복 코드를 공통 함수로 분리` |
| **`Test`** | 테스트 코드 추가 및 수정 | `Test: 인증 로직 단위 테스트 코드 추가` |
| **`Chore`** | 빌드 업무, 패키지 설정, 패키지 설치 등 (소스 코드 변경 없음) | `Chore: dotenv 라이브러리 패키지 설치` |
| **`Design`** | UI/UX 디자인 요소 변경 (CSS, 레이아웃 세부 조정) | `Design: 메인 폼 페이지 버튼 호버 색상 변경` |
| **`Rename`** | 파일 혹은 폴더명을 수정하거나 이동하는 경우 | `Rename: utils 폴더 이름을 helpers로 변경` |
| **`Remove`** | 파일을 삭제하는 작업만 수행한 경우 | `Remove: 사용하지 않는 레거시 이미지 파일 삭제` |

---

### 💡 커밋 메시지 기본 구조
```text
Type: Subject

- Body (왜, 무엇을 변경했는지 세부 내용 작성 - 생략 가능)
- Footer (이슈 트래커 ID 연동 시 사용 - 생략 가능)