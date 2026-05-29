# 공통부문 일반관리비 배분 웹앱 (하이브리드)

산출 엑셀 파일(`template/master.xlsx`)을 그대로 활용해
공통부문 일반관리비를 장금/흥아/한성 3사로 배분합니다.

## 동작 방식

1. **웹 화면(Streamlit)** 에서 기준값 입력 + 6개 파일 업로드
2. 서버가 `master.xlsx`의 지정 시트에 데이터를 주입하고 새 파일 생성
3. 사용자가 파일 다운로드 → **Excel로 열면 모든 수식이 자동 재계산**
4. 결과 확인 및 인쇄는 Excel에서 진행

> 이 파일은 Excel 365 전용 함수(LAMBDA, BYCOL, FILTER, SORT, HSTACK 등)를 다수 사용하기 때문에,
> 정확한 재계산은 **Excel 자체에서만** 보장됩니다. 그래서 웹은 입력·주입까지만 담당합니다.

## 파일 구조

```
expense-allocation-app/
├── app.py                 ← Streamlit 메인 화면
├── core/
│   └── inject.py          ← 기준값/업로드 파일을 master.xlsx에 주입
├── template/master.xlsx   ← 산출 엑셀 원본 (수식 포함)
├── requirements.txt
└── .streamlit/config.toml
```

## 로컬 실행

```bash
python -m venv .venv
.venv\Scripts\activate     # Windows
pip install -r requirements.txt
streamlit run app.py
```

## GitHub + Streamlit Cloud 배포

### 1) GitHub 저장소 생성

GitHub에서 새 저장소(`expense-allocation-app`) Public으로 생성 후, 로컬에서:

```bash
cd expense-allocation-app
git init
git add .
git commit -m "초기 배포"
git branch -M main
git remote add origin https://github.com/<사용자명>/expense-allocation-app.git
git push -u origin main
```

`template/master.xlsx`(약 10MB)도 함께 커밋됩니다. GitHub 단일 파일 제한 100MB 이내라 문제없습니다.

### 2) Streamlit Community Cloud 연결

1. https://share.streamlit.io 접속 → GitHub로 로그인
2. **New app** 클릭
3. Repository: `<사용자명>/expense-allocation-app`
4. Branch: `main`
5. Main file path: `app.py`
6. **Deploy** 클릭

배포 후 `https://<자동이름>.streamlit.app/` 주소로 접속합니다.

이후 GitHub에 push만 하면 자동 재배포됩니다.

## 사용 방법

1. 웹앱에서 **기준 5개 값** 입력
   - 기준 연도, 기준 분기
   - 기간평균환율
   - 장금상선/흥아라인 부가세 매출 신고액
2. **6개 파일** 업로드
   - 장금 급여 / 흥아 급여 (출처: 월급날 개인별수당공제받기)
   - 장금 비용 / 흥아 비용 (출처: 회계시스템 6001-6596 확정전표)
   - 장금 퇴직금 / 흥아 퇴직금 (출처: 월급날 퇴직금추계액)
3. **[배분 파일 생성]** 클릭
4. 다운로드 받은 xlsx를 **Excel로 열기** → 자동 재계산
5. `3사배분`, `배분계산서(수정)` 시트에서 결과 확인 후 인쇄

## 시트 매핑

| 업로드 슬롯 | 주입되는 시트 |
|---|---|
| 장금 급여 | `장금급여입력` |
| 흥아 급여 | `흥아급여입력` |
| 장금 비용 | `장금비용입력` |
| 흥아 비용 | `흥아비용입력` |
| 장금 퇴직금 | `장금퇴충` |
| 흥아 퇴직금 | `흥아퇴충` |

기준값은 `배분계산서(기준) ` 시트의 H2/I2/J2/L2/M2에 기록됩니다.

## 주의사항

- 업로드한 엑셀의 **첫 번째 시트**만 사용합니다.
- 업로드 파일의 **값(수식 결과)만** 가져와 대상 시트에 기록합니다.
- 대상 시트의 기존 값은 모두 지운 뒤 새 값이 기록됩니다.
- master.xlsx의 다른 시트(`1. 급여`, `3사배분` 등)에 있는 수식은 그대로 유지됩니다.
