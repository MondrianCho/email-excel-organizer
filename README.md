# email-excel-organizer

이메일 접수 내역을 엑셀로 정리하고, 백업된 이메일 파일과 엑셀 기록을 서로 대조하는 자동화 스크립트 모음입니다.

## 구성

| 파일 | 설명 |
|------|------|
| `email_excel_organize.py` | 폴더 내 엑셀 파일명을 `yyyy-mm-dd 이메일 접수_초안.xlsx` 형식으로 변경하고, 제목(E열)을 기준으로 업무종류(K열) 등을 자동 입력합니다. |
| `email_excel_reconcile.py` | `오전 이메일 백업` 폴더의 이메일 파일 수와 엑셀에 기록된 행 수를 비교하여 누락/중복을 찾아냅니다. |
| `분류규칙.md` | 제목/발신 기준 자동 분류 규칙표. |
| `run_email_excel_organize.bat` | `email_excel_organize.py` 실행용 배치 파일 (폴더 경로 입력). |
| `run_email_excel_reconcile.bat` | `email_excel_reconcile.py` 실행용 배치 파일 (상위 폴더 경로 입력). |

## 요구 사항

- Python 3.9 이상 (`from __future__ import annotations` 및 `list[...]` 타입 힌트 사용)
- 의존 패키지 설치:

```bash
pip install -r requirements.txt
```

## 사용법

### 엑셀 정리

```bash
python email_excel_organize.py "폴더경로"
```

지정한 폴더 안의 `.xlsx`/`.xlsm` 파일을 찾아 제목(E열) 및 발신(C열)을 기준으로 K/L/N/O열을 채우고, 파일명을 접수일 기준으로 변경합니다. 분류 규칙은 [분류규칙.md](분류규칙.md)를 참고하세요.

### 이메일-엑셀 대조

```bash
python email_excel_reconcile.py "상위폴더경로"
```

`상위폴더경로/오전 이메일 백업` 안의 `.msg`/`.eml` 파일 개수와 엑셀에 기록된 제목 행 수를 비교하여 누락되거나 중복 접수된 항목을 보고합니다.

Windows에서는 `run_email_excel_organize.bat`, `run_email_excel_reconcile.bat`을 더블클릭해 폴더 경로만 입력해도 됩니다.

## 테스트

```bash
python -m unittest discover -s tests
```
