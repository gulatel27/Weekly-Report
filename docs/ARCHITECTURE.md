# 소스 구성과 동작 프로세스

## 전체 흐름

```text
run-final.bat
  -> scripts/read-report-dates.py
  -> npm run download
     -> scripts/download-rockpaper-weekly.js
  -> scripts/build-weekly-report.py
  -> 선택: scripts/apply-to-sharepoint-workbook.py
```

## 1. 실행 입력 처리

`run-final.bat`이 `scripts/read-report-dates.py`를 실행합니다.

`read-report-dates.py`는 사용자에게 아래 값을 입력받습니다.

- 조회 시작일
- 조회 종료일
- 요약 방식: LLM 또는 로컬

입력값은 `logs\report-dates.cmd`에 저장되고, `run-final.bat`은 이 파일을 `call`해서 환경 변수로 불러옵니다.

```text
REPORT_START_DATE=YYYY-MM-DD
REPORT_END_DATE=YYYY-MM-DD
REPORT_SUMMARY_MODE=llm|local
```

## 2. Rockpaper 다운로드

`scripts/download-rockpaper-weekly.js`는 Playwright로 Rockpaper에 접속합니다.

기본 동작:

1. `https://rockpaper.rockplace.com/` 접속
2. 좌측 `통계` 클릭
3. `주간업무보고` 클릭
4. 작업일 시작일/종료일 입력
5. 부서명에서 `DS G` 선택
6. 검색
7. 엑셀 다운로드
8. 다운로드 파일을 `D:\Downloads`에 저장

브라우저 프로필은 `.rockpaper-browser-profile`에 저장되어 로그인 상태를 재사용합니다.

## 3. 다운로드 데이터 정리

`scripts/build-weekly-report.py`는 `D:\Downloads`에서 최신 `주간업무보고*.xlsx`를 찾습니다.

처리 내용:

- `부서명`이 `DS 2T`인 행만 유지
- 요청 시작일/종료일 밖의 작업일시가 있으면 오류 처리
- `지원내역` 텍스트에서 `종류`, `지원내용`, `상세내용`, `이슈` 값을 추출
- `이슈 : "Yes"`인 행을 주간 주요 이슈 대상으로 분류

## 4. 최종 보고서 생성

최종 보고서 템플릿은 `D:\Downloads`의 최신 `락플레이스-지원부문_주간보고_*.xlsx`를 사용합니다.

생성 로직:

- `DS-2팀` 시트만 유지
- 보고서 화면 확대를 100%로 설정
- `2. 주간 주요 이슈` 작성
- `3. 주간 지원 요약` 작성
- `4. Presales 지원사항` 작성
- 전체 세로 정렬을 가운데 맞춤으로 정리
- 불필요한 중간 가로선 제거
- 우측 외곽선 유지
- 병합 셀 겹침 검증
- 저장 대상 파일이 열려 있으면 타임스탬프 파일명으로 자동 저장

## 5. 주간 주요 이슈 작성

대상 조건:

```text
지원내역 안에 이슈 : "Yes"
```

그룹 기준:

- 고객사명
- 지원내용 또는 상세내용 일부

출력 형식:

```text
- 고객사명 (엔지니어)
     . 지원 내용 : ...
     . 구성 : ...
     . 요청/이슈 : ...
     . 원인 : ...
     . 조치 : ...
```

LLM 요약을 선택하면 OpenAI 호환 Chat Completions API로 요약합니다. 실패하면 로컬 요약으로 전환합니다.

로컬 요약은 `상세내용` 안의 라벨을 기준으로 줄을 구성합니다.

- 지원 내용
- 구성
- 요청/이슈
- 원인
- 조치
- 기타

## 6. 주간 지원 요약 집계

시간대 집계:

| 항목 | 기준 |
| --- | --- |
| 주간 | 평일 06:00 이상 18:00 미만 |
| 야간 | 평일 18:00 이상 22:00 미만 |
| 심야 | 평일 22:00 이상 또는 06:00 미만 |
| 휴일 | 토요일, 일요일 전체 |

종류 집계:

- 설치
- 지원
- 점검
- Presales
- 파견

시간대 총합과 종류 총합이 다르면 집계 오류로 보고 중단합니다.

## 7. Presales 지원사항 작성

`종류`가 `Presales`인 행을 대상으로 작성합니다.

| 보고서 컬럼 | 원본 컬럼 |
| --- | --- |
| 일시 | 작업시작일시의 월/일 |
| 엔드유저 | 고객사명 |
| 종류 | Presales |
| 지원 내용 | 지원내역의 지원내용 |
| 담당 | 엔지니어 |
| 영업 | 담당영업 |
| 영업부서 | 공란 |
| 제품 | 제품명 |

기본 3행보다 Presales 건수가 많으면 행을 추가해서 작성합니다.

## 8. SharePoint 반영

`scripts/apply-to-sharepoint-workbook.py`는 생성된 DS-2팀 보고서 시트를 SharePoint 원본 엑셀 파일에 복사합니다.

Excel Online에 직접 붙여넣을 때 열 너비, 병합, 행 높이, 테두리가 깨지는 문제를 피하기 위한 방식입니다.

기본은 원본 파일을 직접 덮어쓰지 않고 `_DS2업데이트.xlsx` 사본을 생성합니다.

## 9. 파일 구성

```text
.
├─ run-final.bat
├─ run-download.bat
├─ run-download-headless.bat
├─ run-weekly-report.bat
├─ run-weekly-report-headless.bat
├─ run-weekly-report-lastweek-test.bat
├─ run-create-form.bat
├─ run-apply-to-sharepoint.bat
├─ package.json
├─ package-lock.json
├─ config
│  └─ llm.example.json
├─ scripts
│  ├─ read-report-dates.py
│  ├─ download-rockpaper-weekly.js
│  ├─ build-weekly-report.py
│  ├─ create-weekly-report-form.py
│  └─ apply-to-sharepoint-workbook.py
└─ docs
   ├─ USAGE.md
   └─ ARCHITECTURE.md
```

## 10. 커밋 제외 대상

아래 파일은 실행 환경 또는 개인정보가 포함될 수 있어 저장소에 올리지 않습니다.

- `config/llm.local.json`
- `.rockpaper-browser-profile`
- `node_modules`
- `logs`
- `outputs`
- `backups`
- 엑셀 산출물
- Python 캐시 파일
