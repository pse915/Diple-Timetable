# 개발 및 검증 보고서

작성: 2026-10-21

## 실행 환경

- 코드: Python 3.12+ 호환 표준 라이브러리 기반 (`tkinter`, `sqlite3`, `zipfile`, `xml.etree`)
- 실제 테스트 실행 환경: Linux / Python 3.13
- Windows EXE 컴파일 환경: 제공된 GitHub Actions `windows-latest` 워크플로 (아직 실행 안 함)
- Google Sheets, Streamlit, Google Cloud 연결: 없음
- 운영 원본 GitHub, Streamlit 및 운영 시트: 변경 명령 수행하지 않음

## 검증 실행

```sh
python -W error::ResourceWarning -m unittest discover -s tests -v
```

결과: **11개 / 11개 통과**

| 테스트 | 결과 |
|---|---|
| 학교명을 입력하지 않은 배포 양식 가져오기 방지 | PASS |
| 샘플 엑셀에서 학교/교사/학급/수업 읽기 | PASS |
| 시간표 편집과 변경 이력 | PASS |
| 존재하지 않는 교사 지정 차단 | PASS |
| 교사 중복 배정 차단과 전체 롤백 | PASS |
| 가능한 맞교환 실행 및 이력 기록 | PASS |
| 불가능한 맞교환에서 데이터 변경 방지 | PASS |
| 수정 내용 엑셀 내보내기 / 재읽기 | PASS |
| 학교 자료 재업로드 시 기존 DB 백업 | PASS |
| 필수 시트 누락 시 오류 처리 | PASS |
| 서로 다른 학교 데이터 저장소 독립성 | PASS |

## 미검증

- Windows EXE 생성 결과 및 Windows 10/11 실제 GUI 실행
- Tkinter 대화상자/파일 선택/마우스 동작 실기기 테스트
- 원본 앱 전체 기능 동등성
- 원본 학교별 엑셀 양식 호환
- 실제 교직원 데이터/운영용 시간표에 대한 검증
- 복잡한 교육행정 규칙·NEIS·Gemini 연동

**테스트를 통과했다고 원본 기능 100%가 구현되었다는 의미는 아닙니다.**
