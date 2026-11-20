# 문서 편집기 — Word · Excel · 한글을 브라우저에서

설치 없이 브라우저에서 **MS Word(.docx), Excel(.xlsx·.xls·.csv), 한글(.hwp·.hwpx)** 파일을 열어 보고, 고치고, 저장하는 웹사이트입니다.
파일은 서버로 보내지 않고 **브라우저 안에서만** 처리합니다.

## 실행 방법

빌드 과정이 없는 정적 사이트입니다.

- **바로 열기**: `index.html`을 브라우저로 열면 됩니다(인터넷 연결 없이도 동작).
- **로컬 서버**: `python3 -m http.server 8000` 실행 후 <http://localhost:8000> 접속
- **GitHub Pages 배포**: 저장소 **Settings → Pages → Build and deployment → Source**를 **GitHub Actions**로 바꾸면,
  `main` 브랜치에 푸시할 때마다 `.github/workflows/pages.yml`이 사이트를 배포합니다.

## 기능

### 공통
- 여러 파일을 탭으로 동시에 열기, 끌어다 놓기로 열기
- 저장하지 않은 변경 표시(탭의 주황 점), 창을 닫을 때 경고
- 단축키: `Ctrl+O` 열기 · `Ctrl+S` 저장 · `Ctrl+P` 인쇄(PDF로 저장)
- 휴대폰·태블릿 화면 지원

### Word(.docx)와 한글(.hwp·.hwpx) — 문서 편집기
| 항목 | 지원 내용 |
|---|---|
| 열기 | .docx, .hwpx, .hwp(한글 5.0 이상, 2002~현재 버전) |
| 서식 | 굵게·기울임·밑줄·취소선·위/아래 첨자, 글꼴, 글자 크기, 글자 색, 형광펜 |
| 문단 | 제목 1~4, 인용, 정렬(왼쪽·가운데·오른쪽·양쪽), 들여쓰기, 글머리 기호·번호 목록 |
| 개체 | 표(행/열 추가·삭제, 셀 배경색, 병합된 셀 유지), 그림 넣기/크기 조절, 링크, 쪽 나누기 |
| 도구 | 찾기/바꾸기, 확대/축소, 글자·단어 수, Word 기준 **페이지 미리보기** |
| 저장 | .docx · .hwpx · .html · .txt · PDF(인쇄) — Word 문서는 기본 .docx, 한글 문서는 기본 .hwpx |

### Excel(.xlsx·.xls·.csv·.ods) — 스프레드시트 편집기
| 항목 | 지원 내용 |
|---|---|
| 시트 | 여러 시트, 시트 추가·이름 바꾸기·복제·이동·삭제 |
| 입력 | 숫자·날짜·시간·백분율·통화 자동 인식, 한글 입력(IME), `Alt+Enter` 줄 바꿈 |
| 수식 | 110여 개 함수(SUM, AVERAGE, IF, IFS, IFERROR, VLOOKUP, XLOOKUP, INDEX/MATCH, SUMIF(S), COUNTIF(S), TEXT, DATE 등), 다른 시트 참조, 순환 참조 감지, 셀을 눌러 참조 넣기 |
| 서식 | 굵게·기울임·밑줄·취소선, 글자 크기·색, 채우기 색, 테두리, 맞춤, 자동 줄 바꿈, 표시 형식(숫자·통화·백분율·날짜), 소수 자릿수 |
| 편집 | 셀 병합/해제, 행·열 삽입/삭제(수식 참조 자동 조정), 정렬, 자동 합계, 열 너비/행 높이 조절(머리글 경계 두 번 클릭 시 자동 맞춤), 복사·잘라내기·붙여넣기(Excel과 주고받기), `Ctrl+D`/`Ctrl+R` 채우기, 실행 취소/다시 실행 |
| 저장 | .xlsx(서식·병합·열 너비·수식 유지), .csv(UTF-8) |

## 알아 둘 점

- **HWP(바이너리) 형식으로는 저장할 수 없습니다.** HWP 파일을 고치면 .hwpx(한글 2014 이상에서 열림) 또는 .docx로 저장합니다.
- HWPX 저장 파일은 한컴오피스가 만든 빈 문서 구조를 바탕으로 만들지만, 이 저장소에서는 한컴오피스로 직접 열어 보는 검증을 하지 못했습니다. 중요한 문서는 .docx로도 함께 저장해 두세요.
- 문서 편집기는 내용·서식 중심이라 머리말/꼬리말, 각주, 글상자 위치, 도형, 다단 같은 쪽 배치 요소는 단순화되거나 생략됩니다.
- 배포용·암호 문서(HWP), 예전 Word 형식(.doc)은 열 수 없습니다.
- 스프레드시트의 차트, 피벗 테이블, 조건부 서식, 매크로는 보존되지 않습니다. 계산할 수 없는 함수는 파일에 저장된 값을 그대로 보여 줍니다.

## 파일 구성

```
index.html            화면 틀
css/style.css         스타일
js/app.js             탭·열기·새로 만들기·저장·인쇄
js/editor.js          문서(Word·한글) 편집 화면
js/model.js           편집 내용(HTML) → 공통 문서 모델
js/word.js            DOCX 읽기/쓰기
js/hwpx.js            HWPX 읽기/쓰기
js/hwp.js             HWP 5.0 읽기
js/hwpx-template.js   HWPX 기본 머리 정보(빈 문서 골격)
js/sheet.js           스프레드시트 편집 화면
js/formula.js         수식 계산기
js/xlsx-io.js         XLSX/XLS/CSV 읽기, XLSX/CSV 쓰기
vendor/               외부 라이브러리
```

## 사용한 오픈소스

| 라이브러리 | 용도 | 라이선스 |
|---|---|---|
| [JSZip](https://github.com/Stuk/jszip) 3.10 | ZIP(docx·xlsx·hwpx) 읽기/쓰기 | MIT 또는 GPLv3 |
| [SheetJS Community Edition](https://github.com/SheetJS/sheetjs) 0.18.5 | 스프레드시트 읽기, HWP 복합 문서(CFB) 읽기 | Apache-2.0 |
| [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) 0.4 | 페이지 미리보기 | Apache-2.0 |
| [python-hwpx](https://github.com/airmang/python-hwpx)의 `Skeleton.hwpx` | HWPX 빈 문서 골격 | Apache-2.0 |

라이선스 원문은 `vendor/licenses/`에 있습니다. HWP 5.0 해석은 한글과컴퓨터가 공개한 *한글 문서 파일 구조 5.0* 문서와
[hwp.js](https://github.com/hahnlee/hwp.js)(Apache-2.0)의 레코드 해석 방식을 참고해 새로 작성했습니다.

> 본 제품은 한글과컴퓨터의 HWP 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.
