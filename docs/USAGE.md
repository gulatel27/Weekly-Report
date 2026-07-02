# 사용법

## 1. 소스 받기

GitHub에서 ZIP으로 내려받거나 `git clone`으로 받은 뒤 프로젝트 폴더로 이동합니다.

```bat
cd Weekly-Report
```

## 2. 필수 프로그램 설치

아래 프로그램이 필요합니다.

- Python 3.10 이상
- Node.js LTS

프로젝트 의존성은 한 번만 설치합니다.

```bat
npm install
```

## 3. 템플릿 파일 준비

최종 주간보고 템플릿 파일을 `D:\Downloads`에 둡니다.

```text
D:\Downloads\락플레이스-지원부문_주간보고_YYYYMMDD.xlsx
```

프로그램은 `D:\Downloads`에서 `락플레이스-지원부문_주간보고_*.xlsx` 패턴의 최신 파일을 템플릿으로 사용합니다.

## 4. LLM 설정

LLM 요약을 사용하지 않을 경우 이 단계는 건너뛰어도 됩니다.

```bat
copy config\llm.example.json config\llm.local.json
```

`config\llm.local.json`의 `api_key`, `model`, `base_url`을 본인 환경에 맞게 수정합니다.

```json
{
  "enabled": true,
  "provider": "openai",
  "api_key": "YOUR_API_KEY_HERE",
  "model": "gpt-4o-mini",
  "base_url": "https://api.openai.com/v1",
  "timeout_seconds": 30,
  "temperature": 0.2,
  "max_tokens": 500
}
```

`config\llm.local.json`은 GitHub에 올리지 않는 개인 설정 파일입니다.

## 5. 최종 실행

```bat
run-final.bat
```

프롬프트에 기간을 `YYYYMMDD` 형식으로 입력합니다.

```text
시작일을 입력하세요 (YYYYMMDD): 20260629
종료일을 입력하세요 (YYYYMMDD): 20260705
```

요약 방식을 선택합니다.

```text
요약 방식을 선택하세요.
1. LLM 요약 사용
2. 로컬 요약 사용
번호를 입력하세요 (1/2):
```

결과 파일은 `outputs` 폴더에 생성됩니다.

```text
outputs\락플레이스-DS2_주간보고_YYYYMMDD.xlsx
```

동일한 파일이 Excel에서 열려 있으면 아래처럼 새 파일명으로 자동 저장됩니다.

```text
outputs\락플레이스-DS2_주간보고_YYYYMMDD_YYYYMMDDHHMMSS.xlsx
```

## 6. 첫 로그인

처음 실행하면 Rockpaper 로그인 상태가 없을 수 있습니다. 이 경우 화면이 보이는 실행 파일로 로그인 상태를 먼저 만들어 둡니다.

```bat
run-download.bat
```

로그인 후 브라우저 프로필은 `.rockpaper-browser-profile` 폴더에 저장됩니다. 이 폴더는 GitHub에 올리지 않습니다.

## 7. SharePoint 업로드용 파일 생성

OneDrive로 동기화된 SharePoint 주간보고 원본 파일이 있으면 인자로 넘깁니다.

```bat
run-final.bat "C:\path\to\sharepoint-weekly-report.xlsx"
```

원본 파일은 직접 수정하지 않고 같은 폴더에 아래 형식의 사본을 만듭니다.

```text
원본파일명_DS2업데이트.xlsx
```

이미 생성된 보고서를 특정 SharePoint 원본 파일에만 반영하려면 아래 파일을 사용합니다.

```bat
run-apply-to-sharepoint.bat "C:\path\to\sharepoint-weekly-report.xlsx"
```

## 8. 환경 변수 옵션

필요하면 실행 전에 환경 변수를 지정할 수 있습니다.

| 변수 | 기본값 | 설명 |
| --- | --- | --- |
| `REPORT_START_DATE` | 실행 시 입력 | 조회 시작일 |
| `REPORT_END_DATE` | 실행 시 입력 | 조회 종료일 |
| `REPORT_SUMMARY_MODE` | 실행 시 입력 | `llm` 또는 `local` |
| `ROCKPAPER_DEPARTMENT` | `DS G` | Rockpaper 부서 드롭다운 선택값 |
| `DOWNLOAD_DIR` | `D:\Downloads` | 다운로드 파일 저장 위치 |
| `ROCKPAPER_PROFILE_DIR` | `.rockpaper-browser-profile` | Playwright 로그인 프로필 위치 |
| `LLM_CONFIG_PATH` | `config\llm.local.json` | LLM 설정 파일 위치 |
| `HEADLESS` | `1` | 브라우저 창 표시 여부. `0`이면 화면 표시 |

## 9. 자주 나는 오류

### PermissionError: Permission denied

생성하려는 엑셀 파일이 Excel에서 열려 있는 경우입니다. 현재 버전은 자동으로 새 파일명에 타임스탬프를 붙여 저장합니다.

### 다운로드 데이터가 요청 주간 밖 날짜를 포함합니다

Rockpaper에서 내려받은 엑셀의 작업일시가 입력한 시작일/종료일 밖에 있는 경우입니다. 입력 기간 또는 Rockpaper 조회 조건을 확인합니다.

### LLM 요약 실패

API 키, 모델명, 잔여 한도, 네트워크를 확인합니다. 실패 시 프로그램은 로컬 요약으로 전환하고 계속 진행합니다.
