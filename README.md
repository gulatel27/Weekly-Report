# Weekly Report Automation

Rockpaper 주간업무보고 데이터를 내려받아 DS-2팀 주간보고 엑셀을 자동 생성하는 도구입니다.

## 빠른 실행

```bat
run-final.bat
```

실행하면 아래 순서로 진행됩니다.

1. 조회 시작일 입력
2. 조회 종료일 입력
3. 요약 방식 선택
4. Rockpaper 주간업무보고 엑셀 다운로드
5. DS-2팀 최종 주간보고 엑셀 생성

```text
시작일을 입력하세요 (YYYYMMDD): 20260629
종료일을 입력하세요 (YYYYMMDD): 20260705

요약 방식을 선택하세요.
1. LLM 요약 사용
2. 로컬 요약 사용
번호를 입력하세요 (1/2):
```

생성 파일은 `outputs` 폴더에 저장됩니다. 같은 이름의 엑셀 파일이 열려 있으면 덮어쓰지 못하므로, 자동으로 타임스탬프가 붙은 새 파일명으로 저장합니다.

## 설치

필수 항목:

- Windows
- Python 3.10 이상
- Node.js LTS
- Rockpaper 접속 권한
- 최종 주간보고 템플릿 엑셀 파일

처음 받은 뒤 프로젝트 폴더에서 한 번 실행합니다.

```bat
npm install
```

최종 주간보고 템플릿 파일은 기본적으로 `D:\Downloads`에서 아래 패턴으로 찾습니다.

```text
락플레이스-지원부문_주간보고_*.xlsx
```

## LLM 설정

LLM 요약은 선택 기능입니다. 사용하려면 예시 파일을 복사해서 로컬 설정을 만듭니다.

```bat
copy config\llm.example.json config\llm.local.json
```

`config\llm.local.json`에 개인 또는 회사 API 키를 입력합니다. 이 파일은 `.gitignore`에 포함되어 GitHub에 올라가지 않습니다.

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

LLM 호출이 실패하면 로컬 요약으로 자동 전환됩니다. 실행 중에는 사용 토큰과 rate limit 잔여 정보도 출력합니다.

## SharePoint 파일 만들기

OneDrive로 동기화된 SharePoint 원본 파일에 DS-2팀 시트를 반영한 업로드용 사본을 만들 수 있습니다.

```bat
run-final.bat "C:\path\to\sharepoint-weekly-report.xlsx"
```

원본 파일은 직접 덮어쓰지 않고 같은 폴더에 `원본파일명_DS2업데이트.xlsx` 파일을 생성합니다.

## 문서

- [사용법](docs/USAGE.md)
- [소스 구성과 동작 프로세스](docs/ARCHITECTURE.md)

## 주요 파일

| 파일 | 설명 |
| --- | --- |
| `run-final.bat` | 최종 실행 파일 |
| `scripts/download-rockpaper-weekly.js` | Rockpaper 접속 및 엑셀 다운로드 |
| `scripts/build-weekly-report.py` | 다운로드 엑셀 분석 및 최종 보고서 생성 |
| `scripts/apply-to-sharepoint-workbook.py` | SharePoint 원본 엑셀에 DS-2팀 시트 반영 |
| `scripts/read-report-dates.py` | 실행 시 기간/요약 방식 입력 처리 |
| `config/llm.example.json` | LLM 설정 예시 |

## 커밋 제외 파일

아래 항목은 GitHub에 올리지 않습니다.

- `config/llm.local.json`
- `node_modules`
- `.rockpaper-browser-profile`
- `logs`
- `outputs`
- `backups`
- 엑셀 산출물
