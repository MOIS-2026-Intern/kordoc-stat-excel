<p align="center">
  <img src="src/logo2.png" alt="한글 표 추출기 로고" width="96" />
</p>

<h1 align="center">HWP(X) → Excel 표 추출기</h1>

<p align="center">
  kordoc으로 한글 문서를 Markdown으로 변환한 뒤, 문서 안의 표를 Excel 파일로 추출하는 Streamlit 앱입니다.
</p>

<p align="center">
  <a href="https://mois-kordoc-parsing.streamlit.app/"><strong>배포된 사이트 바로가기</strong></a>
</p>

<p align="center">
  <img alt="Python" src="https://img.shields.io/badge/Python-3.x-3776AB?logo=python&logoColor=white">
  <img alt="Streamlit" src="https://img.shields.io/badge/Streamlit-App-FF4B4B?logo=streamlit&logoColor=white">
  <img alt="Excel" src="https://img.shields.io/badge/Output-XLSX-217346?logo=microsoftexcel&logoColor=white">
</p>

---

## 소개

행정안전부 통계 문서처럼 표가 많은 HWP/HWPX 문서를 Excel로 옮기는 작업을 줄이기 위한 도구입니다.

문서를 업로드하면 앱이 표를 찾아 `.xlsx` 파일로 변환하고, 변환된 파일들을 `.zip`으로 묶어 내려받을 수 있습니다.

## 주요 기능

- HWP, HWPX, PDF 등 kordoc이 처리할 수 있는 문서 업로드
- 여러 파일을 한 번에 업로드하고 각각 ZIP으로 다운로드
- 열 너비, 줄바꿈, 헤더 굵게 표시 등 기본 Excel 서식 적용

## 온라인에서 사용하기

가장 간단한 방법은 배포된 Streamlit 앱을 사용하는 것입니다.

[https://mois-kordoc-parsing.streamlit.app/](https://mois-kordoc-parsing.streamlit.app/)

1. 추출기를 선택합니다.
2. 변환할 파일을 업로드합니다.
3. 변환이 끝나면 ZIP 파일을 다운로드합니다.
4. ZIP 안의 Excel 파일을 확인합니다.

## 로컬에서 실행하기

로컬에서 직접 실행하려면 Python과 Node.js/npm이 필요합니다.


### 1. 저장소 준비

```bash
git clone <repository-url>
cd kordoc-stat-excel
```

이미 저장소를 받아둔 경우에는 프로젝트 폴더로 이동하면 됩니다.

### 2. Python 가상환경 만들기

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Windows에서는 다음 명령을 사용할 수 있습니다.

```bash
.venv\Scripts\activate
```

### 3. Python 패키지 설치

```bash
pip install -r requirements.txt
```

### 4. Node.js/npm 확인

```bash
node --version
npm --version
```

두 명령이 모두 버전을 출력해야 합니다.

### 5. Streamlit 앱 실행

```bash
streamlit run app.py
```

실행 후 브라우저에서 보통 아래 주소로 접속합니다.

```text
http://localhost:8501
```

## 로컬 사용 흐름

1. 브라우저에서 로컬 Streamlit 앱을 엽니다.
2. `추출기 선택`에서 문서에 맞는 추출기를 고릅니다.
3. `파일 업로드` 영역에 변환할 파일을 올립니다.
4. 진행률이 완료될 때까지 기다립니다.
5. `엑셀 다운로드` 버튼으로 ZIP 파일을 내려받습니다.

## 추출기 선택 기준

| 추출기 | 추천 상황 | 출력 방식 |
| --- | --- | --- |
| 범용 추출기 | 일반적인 Markdown/HTML 표가 있는 문서 | 표마다 별도 Excel 파일 생성 |
| 행정안전통계연보 추출기 | `1-1-1 제목`처럼 코드형 제목이 있는 통계연보 | 각 섹션 별 시트로 저장 |
| 주요통계집 추출기 | `1-1`, `참고-1` 같은 번호와 직전 제목 줄이 있는 문서 | 각 섹션 별 시트로 저장 |

기본값은 `범용 추출기`입니다. 문서 구조를 잘 모를 때는 먼저 범용 추출기로 실행해 보는 것이 좋습니다.

## 동작 원리

```text
업로드 파일
  ↓
임시 폴더에 원본 저장
  ↓
npx kordoc으로 Markdown 변환
  ↓
선택한 추출기로 표 탐색
  ↓
HTML/Markdown 표를 Excel용 그리드로 파싱
  ↓
openpyxl로 .xlsx 생성
  ↓
여러 Excel 파일을 ZIP으로 압축
  ↓
Streamlit 다운로드 버튼으로 제공
```



## 프로젝트 구조

```text
.
├── app.py                         # Streamlit UI 진입점
├── pipeline.py                    # 업로드 파일 → Markdown → Excel ZIP 변환 파이프라인
├── extract_tables_generic.py      # 범용 표 추출기
├── extract_tables_moisstat.py     # 행정안전통계연보 전용 추출기
├── extract_tables_keystat.py      # 주요통계집 전용 추출기
├── table_extractor.py             # 표 탐색과 Markdown 표 파싱
├── table_utils.py                 # HTML 표 파싱과 Excel 저장 유틸리티
├── requirements.txt               # Python 의존성
├── packages.txt                   # Streamlit 배포 환경의 Node/npm 의존성
└── src/                           # 로고 이미지
```

## 주의 사항

- 변환 품질은 원본 문서 구조와 kordoc 변환 결과에 영향을 받습니다.
- 표가 이미지로만 들어 있는 문서는 Excel 표로 추출되지 않을 수 있습니다.
- 첫 실행 시 `npx`가 kordoc 패키지를 내려받느라 시간이 걸릴 수 있습니다.
- 배포 앱은 베타 버전이며, 일부 문서는 변환이 실패할 수 있습니다.
