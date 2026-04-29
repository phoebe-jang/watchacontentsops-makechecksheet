# 검수시트 자동 생성기

왓챠 OTT 콘텐츠 편성팀용 — 주간편성표 CSV를 업로드해서 검수시트(70컬럼) 양식으로 변환·다운로드하는 Streamlit 웹 대시보드.

## 기능

- 시작/끝 요일 범위 선택 (월~금. 토/일은 자동으로 금요일에 합쳐짐)
- 한 시트에 여러 요일 검수시트 이어 출력
- 처리 룰:
  - 신규 / 예약작 / 결방·홀드백 / 종영 / 회차 분할 / mapping_type(자막·더빙)
  - 전날 셋팅 (전주 금에 이미 들어간 행 제외)
  - 연휴지연편성 (헤더 + 콘텐츠 두 톤 회색)
  - 휴일 자동 판정 (KR_HOLIDAYS_2026 기준)
  - 종영 의심 행 옅은 파랑 + 시트 상단 알림
  - 금요일 EBS 별도 섹션
- 검수시트 템플릿 색상 매핑 (Arial 10pt)
- 파일명 인라인 편집 + xlsx 다운로드
- 팀 공용 비밀번호 게이트

## 로컬 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

비밀번호는 `.streamlit/secrets.toml`의 `app_password` 값(로컬 fallback: `watchacontents`).

## Streamlit Cloud 배포

1. GitHub repo 연결 → New app 생성
2. Settings → **Secrets**에 다음 추가:
   ```toml
   app_password = "팀-실제-비밀번호"
   ```
3. Deploy

## 라이선스

내부용.
