"""공통부문 일반관리비 배분 - 하이브리드(Streamlit + Excel)."""

from __future__ import annotations

import tempfile
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

import streamlit as st

from core.inject import populate_workbook, UPLOAD_TARGET_SHEETS


TEMPLATE_PATH = Path(__file__).parent / "template" / "master.xlsx"


st.set_page_config(
    page_title="공통부문 일반관리비 배분",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .stApp { font-family: "맑은 고딕", "Malgun Gothic", sans-serif; }
    .section-card {
        border:1px solid #e5e7eb; border-radius:8px;
        padding:16px 20px; margin: 0 0 16px 0; background:#fafafa;
    }
    .section-title {
        font-size:16px; font-weight:700; color:#1f3a68;
        margin-bottom:8px;
    }
    .source-note {
        font-size:12px; color:#6b7280; margin-bottom:8px;
    }
    .info-banner {
        background:#eff6ff; border-left:4px solid #1f3a68;
        padding:10px 14px; margin: 8px 0 16px 0; font-size:14px;
        border-radius:4px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------
# 상단 안내
# ----------------------------------------------------------------------
st.title("공통부문 일반관리비 배분")
st.markdown(
    """
    <div class="info-banner">
    <strong>사용 흐름</strong>: 기준값을 입력하고 6개 파일을 업로드한 후
    <b>[배분 파일 생성]</b> 버튼을 누르세요. 데이터가 주입된 엑셀 파일이 생성되며,
    <b>Excel로 열면 모든 수식이 자동으로 재계산</b>됩니다.<br>
    결과 확인 및 인쇄는 Excel에서 진행하세요.
    </div>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------
# 1) 기준 입력
# ----------------------------------------------------------------------
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">1. 기준 입력</div>', unsafe_allow_html=True)

c1, c2, c3 = st.columns(3)
with c1:
    year = st.number_input("기준 연도", min_value=2000, max_value=2100, value=2026, step=1)
with c2:
    quarter = st.number_input("기준 분기", min_value=1, max_value=4, value=1, step=1)
with c3:
    fx_rate = st.number_input(
        "기간평균환율 (최초매매기준율)",
        min_value=0.0, value=1464.66, step=0.01, format="%.2f",
    )

c4, c5 = st.columns(2)
with c4:
    janggeum_vat = st.number_input(
        "장금상선 부가세 매출 신고액",
        min_value=0.0, value=614629993110.0, step=1.0, format="%.0f",
    )
with c5:
    heungah_vat = st.number_input(
        "흥아라인 부가세 매출 신고액",
        min_value=0.0, value=426036916131.0, step=1.0, format="%.0f",
    )
st.markdown('</div>', unsafe_allow_html=True)

# ----------------------------------------------------------------------
# 2) 파일 업로드
# ----------------------------------------------------------------------
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">2. 파일 업로드 (총 6개)</div>', unsafe_allow_html=True)

uploads: Dict[str, Optional[object]] = {}


def section(title: str, source_note: str, slots: list[tuple[str, str]]) -> None:
    st.markdown(f"**{title}**")
    st.markdown(
        f'<div class="source-note">{source_note}</div>',
        unsafe_allow_html=True,
    )
    cols = st.columns(len(slots))
    for col, (slot_key, label) in zip(cols, slots):
        with col:
            uploads[slot_key] = st.file_uploader(
                label, type=["xlsx", "xlsm"], key=f"up_{slot_key}",
            )


section(
    "급여 부분",
    "출처 : 월급날 개인별수당공제받기(월별, 대장별) / Excel파일",
    [("janggeum_salary", "장금상선 급여"), ("heungah_salary", "흥아라인 급여")],
)
st.markdown("")
section(
    "비용 부분",
    "출처 : 회계시스템 6001-6596 확정전표 / Excel파일",
    [("janggeum_cost", "장금상선 비용"), ("heungah_cost", "흥아라인 비용")],
)
st.markdown("")
section(
    "퇴직금 부분",
    "출처 : 월급날 퇴직금추계액 / Excel파일",
    [("janggeum_retire", "장금상선 퇴직금"), ("heungah_retire", "흥아라인 퇴직금")],
)

st.markdown('</div>', unsafe_allow_html=True)

# ----------------------------------------------------------------------
# 3) 실행 버튼
# ----------------------------------------------------------------------
run = st.button("배분 파일 생성", type="primary", use_container_width=True)


def _save_upload_to_temp(uploaded_file) -> Optional[Path]:
    if uploaded_file is None:
        return None
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(uploaded_file.getvalue())
    tmp.flush()
    tmp.close()
    return Path(tmp.name)


if "result_bytes" not in st.session_state:
    st.session_state.result_bytes = None
    st.session_state.result_name = None

if run:
    missing_labels = {
        "janggeum_salary": "장금 급여",
        "heungah_salary": "흥아 급여",
        "janggeum_cost": "장금 비용",
        "heungah_cost": "흥아 비용",
        "janggeum_retire": "장금 퇴직금",
        "heungah_retire": "흥아 퇴직금",
    }
    missing = [
        missing_labels[s] for s in UPLOAD_TARGET_SHEETS
        if uploads.get(s) is None
    ]
    if missing:
        st.error("다음 파일이 업로드되지 않았습니다: " + ", ".join(missing))
    elif not TEMPLATE_PATH.exists():
        st.error(f"템플릿 파일이 없습니다: {TEMPLATE_PATH}")
    else:
        with st.spinner("데이터를 master.xlsx에 주입하는 중..."):
            upload_paths = {
                slot: _save_upload_to_temp(f) for slot, f in uploads.items()
            }
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as out_tmp:
                out_path = Path(out_tmp.name)
            populate_workbook(
                template_path=TEMPLATE_PATH,
                output_path=out_path,
                year=int(year),
                quarter=int(quarter),
                fx_rate=float(fx_rate),
                janggeum_vat_sales=float(janggeum_vat),
                heungah_vat_sales=float(heungah_vat),
                uploads=upload_paths,
            )
            st.session_state.result_bytes = out_path.read_bytes()
            st.session_state.result_name = (
                f"{int(year)}년_{int(quarter)}분기_공통부문_일반관리비_배분"
                f"_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
        st.success("파일 생성 완료. 아래에서 다운로드 후 Excel에서 열어주세요.")

# ----------------------------------------------------------------------
# 4) 다운로드 + 사용 안내
# ----------------------------------------------------------------------
if st.session_state.result_bytes:
    st.download_button(
        "배분 엑셀 다운로드",
        data=st.session_state.result_bytes,
        file_name=st.session_state.result_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.markdown(
        """
        <div class="info-banner">
        <strong>다음 단계</strong><br>
        ① 다운로드한 파일을 <b>Excel로 열기</b><br>
        ② Excel이 자동으로 모든 수식을 재계산합니다 (수초 ~ 수십초)<br>
        ③ <code>3사배분</code> 시트와 <code>배분계산서(수정)</code> 시트에서 결과 확인<br>
        ④ 인쇄는 Excel의 [파일] → [인쇄] 메뉴로 진행
        </div>
        """,
        unsafe_allow_html=True,
    )
