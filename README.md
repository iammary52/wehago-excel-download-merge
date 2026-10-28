# 위하고 엑셀 다운로드·통합 스킬

위하고의 재무상태표·손익계산서·분개장·총괄원장을 내려받아 보고서별 시트로 합치는 Codex 스킬입니다.

- 회사·연도·기준일을 지정해 재사용
- 총괄원장: 거래처잔액 → 프린트 → 엑셀 저장 → 시트분할 없음
- 금액은 숫자, 날짜는 실제 날짜, 각종 코드는 텍스트로 보존
- 결과는 **일반 셀 범위**에 자동 필터·틀 고정·서식 적용. 엑셀 데이터 테이블은 만들지 않습니다.

## 다른 PC에서 설치

GitHub에서 **Code → Download ZIP**을 선택하고 압축을 풉니다. `SKILL.md`가 들어 있는 폴더의 이름을 `wehago-excel-download-merge`로 맞춰 아래에 넣습니다.

```text
Windows: %USERPROFILE%\.codex\skills\wehago-excel-download-merge
macOS/Linux: ~/.codex/skills/wehago-excel-download-merge
```

CODEX_HOME을 별도 설정한 경우 그 아래 `skills` 폴더를 사용합니다. 기존 `wehago-audit-export` 스킬이 있으면 중복 등록을 피하도록 하나만 사용합니다.

Git을 사용할 수 있으면 다음과 같이 설치할 수도 있습니다(대상 폴더가 아직 없을 때).

```text
git clone https://github.com/iammary52/wehago-excel-download-merge.git ~/.codex/skills/wehago-excel-download-merge
```

이후 새 작업을 열고 다음처럼 요청합니다.

> $wehago-excel-download-merge 예시회사 2026년 중간감사 6월말과 9월말 자료를 다운로드하고 합쳐줘.

대상 PC에서 위하고에 로그인할 수 있어야 하며 브라우저/컴퓨터 제어 도구가 필요합니다. 통합 도우미를 실행할 때에는 Python과 openpyxl이 필요합니다. 구형 `.xls`를 읽는 작업에는 xlrd 등 별도 판독기가 필요할 수 있습니다.

이 저장소에는 스킬 지침과 통합 도우미만 포함됩니다. 회사 재무자료·로그인 정보는 포함하지 않습니다.
