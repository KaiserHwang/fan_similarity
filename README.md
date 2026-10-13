# Fan Studio Mobile — 팬 상사법칙 예측·설계 (스마트폰 버전)

`fan_similarity_v10`의 계산 엔진을 그대로 유지하면서, 스마트폰에서 쓰기 쉽게 화면을 다시 구성한 Streamlit 앱입니다.

## 실행

```bash
pip install -r requirements.txt
streamlit run fan_similarity_mobile.py
```

같은 Wi‑Fi의 스마트폰에서 접속하려면:

```bash
streamlit run fan_similarity_mobile.py --server.address 0.0.0.0
# 휴대폰 브라우저에서 http://<PC IP>:8501
```

[Streamlit Community Cloud](https://share.streamlit.io)에 이 저장소를 연결하고 메인 파일을 `fan_similarity_mobile.py`로 지정하면 외부에서도 접속할 수 있습니다.

계산 회귀 테스트:

```bash
python fan_similarity_mobile.py --self-test
```

## 기능 (7개 섹션)

| 섹션 | 내용 |
|---|---|
| 📈 단품 PQ | 단품 PQ선도 상사 예측, 덕트 저항곡선 운전점 |
| 🏠 세트 예측 | 세트 raw data 보정, 부품·신규 팬 교체 후 RPM·정압·소음 예측 |
| 📐 경험식 | Cordier 선도 기반 경험식 설계와 σ 적합성 |
| 🧪 실험 기반 | 실험 DB에서 파생 설계 후보 탐색·순위 |
| ⚖️ 1:1 비교 | 기준 팬 ↔ 목표 팬 단일 운전점 비교 |
| 🔍 설계 탐색 | 직경·높이·RPM 2D/3D 스윕, 목표 사양 탐색 |
| 🌀 덕트 손실 | 배관·피팅 정압 손실 계산기 |

## v10 대비 모바일 변경점

- **설정 패널**: 사이드바 → 화면 상단 접이식 패널(팬 계열 / 단위 / 운전 환경 3개 탭). 현재 설정은 패널 아래 칩으로 항상 표시됩니다.
- **섹션 이동**: 아이콘이 붙은 가로 스크롤 알약형 탭.
- **폰트**: Pretendard(웹폰트), 입력창 16px로 iOS 자동 확대 방지, 터치 영역 44px 이상.
- **레이아웃**: 최대 폭 880px 단일 열, 좁은 화면에서 열이 자동으로 세로 배치되고 버튼은 전체 너비.
- **차트**: 드래그를 꺼서 화면 스크롤과 충돌하지 않게 하고, 탭하면 값이 표시됩니다. 제목은 차트 위로 옮겨 툴바와 겹치지 않으며, 세트 비교 차트(PQ·소음)는 위아래로 배치했습니다.
- **계산식·보고서(PPT)·저장(SQLite) 로직은 v10과 동일**합니다. 입력 저장 파일 이름만 `.fan_similarity_mobile_inputs.sqlite3`로 분리했습니다.
