"""
CA Dashboard — Streamlit 웹 대시보드
=====================================
Monthly 종합 성과 테이블 · QANDA 디자인 시스템(QDS3) 적용
Progress Board with editable targets + What-if simulator

실행: python3 -m streamlit run dashboard.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import json
import datetime
from pathlib import Path

# ── 기존 분석 모듈 임포트 ──
from ca_dashboard_analysis import (
    load_all_blocks,
    merge_all_blocks,
    compute_kpis,
    VALID_MANAGERS,
)

# ============================================================
# 목표치 저장/로드 (targets.json)
# ============================================================
TARGETS_FILE = Path(__file__).resolve().parent / "targets.json"

# 프로그레스 보드에 표시할 10개 핵심 KPI (대분류/중분류 그룹 포함)
PROGRESS_KPIS = [
    {"cat1": "리드",  "cat2": "유입", "label": "리드 수(전체)",              "key": "lead_cnt",               "unit": "건", "fmt": "int",   "agg": "sum_mgr"},
    {"cat1": "활동",  "cat2": "발신", "label": "발신 횟수(전체)",            "key": "total_calls",            "unit": "건", "fmt": "int",   "agg": "sum_mgr"},
    {"cat1": "활동",  "cat2": "통화", "label": "통화 연결 성공 건수",        "key": "success_calls",          "unit": "건", "fmt": "int",   "agg": "sum_mgr"},
    {"cat1": "활동",  "cat2": "효율", "label": "통화 시도/전체 리드(비율)",  "key": "_attempt_per_total",     "unit": "",   "fmt": "ratio", "agg": "derived"},
    {"cat1": "성과",  "cat2": "신규", "label": "NewStdnt(a/c+organic)",      "key": "newstdt_total",          "unit": "건", "fmt": "int",   "agg": "sum_dedup"},
    {"cat1": "성과",  "cat2": "신규", "label": "NewStdnt(a/c)",              "key": "newstdt_ac",             "unit": "건", "fmt": "int",   "agg": "sum_dedup"},
    {"cat1": "성과",  "cat2": "전환", "label": "Up/Cross Selling(#)",        "key": "up_cross_selling_count", "unit": "건", "fmt": "int",   "agg": "sum_dedup"},
    {"cat1": "성과",  "cat2": "전환", "label": "상담 결제율(CVR_Total)",     "key": "_cvr_total",             "unit": "%",  "fmt": "pct",   "agg": "derived"},
    {"cat1": "성과",  "cat2": "전환", "label": "결제 CVR(허수 제거)",        "key": "_cvr_clean",             "unit": "%",  "fmt": "pct",   "agg": "derived"},
    {"cat1": "성과",  "cat2": "매출", "label": "Booking(M)",                 "key": "net_booking_in_krw",     "unit": "원", "fmt": "krw",   "agg": "sum_dedup"},
]


def _load_targets() -> dict:
    if TARGETS_FILE.exists():
        try:
            return json.loads(TARGETS_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _save_targets(data: dict):
    TARGETS_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _get_current_month() -> str:
    return datetime.date.today().strftime("%Y-%m")


def _get_business_days_in_month(year: int, month: int) -> int:
    import calendar
    cal = calendar.Calendar()
    return sum(1 for d in cal.itermonthdays2(year, month)
               if d[0] != 0 and d[1] < 5)


def _get_remaining_business_days(year: int, month: int) -> int:
    import calendar
    today = datetime.date.today()
    cal = calendar.Calendar()
    return sum(1 for d in cal.itermonthdays2(year, month)
               if d[0] != 0 and d[1] < 5
               and datetime.date(year, month, d[0]) >= today)


# ============================================================
# QDS3 Color System
# ============================================================
COLORS = {
    "text":         "#222222",
    "sub_text":     "#5D5D5D",
    "inactive":     "#999999",
    "border":       "#D0D0D0",
    "border_light": "#F0F0F0",
    "canvas":       "#F9F9F9",
    "card":         "#FFFFFF",
    "orange":       "#FF5500",
    "orange_mid":   "#FF8C4B",
    "orange_soft":  "#FFB899",
    "orange_light": "#FEF1EB",
    "success":      "#0D9974",
    "success_bg":   "#ECF7F4",
    "negative":     "#FB2D36",
    "negative_bg":  "#FFEEEF",
}

# ============================================================
# 페이지 설정
# ============================================================
st.set_page_config(page_title="CA Dashboard", page_icon="", layout="wide")

# ============================================================
# QDS3 기반 커스텀 CSS
# ============================================================
st.markdown("""
<style>
  .stApp { background-color: #F9F9F9; font-family: 'Pretendard','Calibri',-apple-system,sans-serif; }
  [data-testid="stSidebar"] { background-color:#FFF; border-right:1px solid #F0F0F0; }
  h1 { color:#222 !important; font-weight:700 !important; }
  h2,h3,h4,h5 { color:#222 !important; font-weight:600 !important; }
  p,li,span { color:#5D5D5D; }
  [data-testid="stVerticalBlock"]>div { border-radius:8px; background:transparent !important; }
  thead th { background-color:#F0F0F0 !important; color:#222 !important;
             font-size:0.8rem !important; font-weight:700 !important;
             text-align:center !important; white-space:nowrap !important;
             border-bottom:1px solid #E0E0E0 !important; }
  tbody td { text-align:center !important; font-size:0.8rem !important;
             color:#222 !important; white-space:nowrap !important;
             border-bottom:1px solid #F0F0F0 !important; }
  [data-testid="stMetricValue"] { font-size:1.6rem; font-weight:700; color:#FF5500 !important; }
  [data-testid="stMetricLabel"] { color:#5D5D5D !important; font-size:0.85rem; }
  details { border:1px solid #F0F0F0 !important; border-radius:8px !important; background:#FFF !important; }
  summary { color:#222 !important; font-weight:600 !important; }
  hr { border-color:#F0F0F0 !important; }
  [data-testid="stDataFrame"] { max-height:none !important; }
</style>
""", unsafe_allow_html=True)


# ============================================================
# 데이터 로드 (캐싱)
# ============================================================
@st.cache_data(ttl=300, show_spinner="구글 시트에서 데이터를 가져오는 중...", hash_funcs={str: hash})
def load_data():
    blocks, source = load_all_blocks()
    weekly_mgr, monthly_mgr = merge_all_blocks(blocks)
    weekly_kpi = compute_kpis(weekly_mgr)
    return weekly_kpi, monthly_mgr, source


# ============================================================
# 지표 트리 정의 (대분류 / 중분류 / 소분류)
# ============================================================
METRIC_TREE = [
    ("리드", "유입", [
        ("리드 수(전체)",              "lead_cnt",         "sum_mgr",  "int",   None,     True,  True),
        ("유효 리드 수(허수 제거)",     "clean_lead_cnt",   "sum_mgr",  "int",   None,     True,  True),
        ("허수 비율(%)",               "_invalid_ratio",   "derived",  "pct",   "purple", True,  True),
    ]),
    ("리드", "진행", [
        ("상담 완료 수",                       "consulted_cnt",    "sum_mgr",  "int",   None,     True,  True),
        ("상담 완료 상태가 아닌 리드 비율(%)",  "_nonclosed_ratio", "derived",  "pct",   "purple", True,  True),
    ]),
    ("리드", "관리", [
        ("결제 링크 접속 후 미관리된 리드 비율(%)", "_b_unmanaged", "derived", "pct", "yellow", True, True),
        ("(B) 미관리 리드 건수",        "f_b_push_target",    "sum_mgr",  "int",   "yellow", False, True),
        ("(B) 관리했으나 FAIL 건수",    "f_b_push_fail",      "sum_mgr",  "int",   "yellow", False, True),
        ("결제 진행중 미관리된 리드 비율(%)", "_c_unmanaged", "derived", "pct", "yellow", True, True),
        ("(C) 미관리 리드 건수",        "f_c_push_target",    "sum_mgr",  "int",   "yellow", False, True),
        ("(C) 관리했으나 FAIL 건수",    "f_c_push_fail",      "sum_mgr",  "int",   "yellow", False, True),
    ]),
    ("활동", "발신", [
        ("발신 횟수(전체)",        "total_calls",      "sum_mgr", "int", None, True,  True),
        ("발신 횟수(정상 번호)",   "mobile_010_calls", "sum_mgr", "int", None, True,  True),
        ("통화 연결 성공 건수",    "success_calls",    "sum_mgr", "int", None, True,  True),
    ]),
    ("활동", "효율", [
        ("리드당 발신(정상 번호)",              "_calls_per_lead_clean",  "derived", "ratio", None, True,  True),
        ("리드당 발신(허수 제거)",              "_calls_per_lead_total",  "derived", "ratio", None, True,  True),
        ("통화 시도/전체 리드(비율)",           "_attempt_per_total",     "derived", "ratio", None, True,  True),
        ("통화 시도/유효 리드(비율)",           "_attempt_per_clean",     "derived", "ratio", None, True,  True),
        ("총 노동 환산 시간(Effort)",           "_effort_hours",          "derived", "hours", None, True,  True),
    ]),
    ("성과", "전환", [
        ("상담 후 결제 전환 건수",               "success_cnt",     "sum_mgr",  "int",   None,      True,  True),
        ("상담 결제율(CVR_Total)",              "_cvr_total",      "derived",  "pct",   "mkt_tag", True,  True),
        ("결제 CVR(허수 제거)",                 "_cvr_clean",      "derived",  "pct",   None,      True,  True),
        ("결제 CVR(상담 완료 기준)",            "_cvr_consulted",  "derived",  "pct",   None,      True,  True),
    ]),
    ("성과", "매출", [
        ("결제 금액",              "paid_cost_w",          "sum_mgr",  "krw",  None, True,  True),
        ("coupon 사용 비율",       "_coupon_ratio",        "derived",  "pct",  None, True,  True),
        ("환불 금액",              "total_refund_amount",  "sum_mgr",  "krw",  None, True,  True),
        ("환불 금액의 비율",       "_refund_ratio",        "derived",  "pct",  None, True,  True),
    ]),
    ("성과", "신규", [
        ("NewStdnt(a/c+organic)", "newstdt_total",          "sum_dedup", "int", None, True,  False),
        ("NewStdnt(a/c)",         "newstdt_ac",             "sum_dedup", "int", None, True,  False),
        ("Up/Cross Selling(#)",   "up_cross_selling_count", "sum_dedup", "int", None, True,  False),
    ]),
    ("성과", "부킹", [
        ("Booking(M)",  "net_booking_in_krw", "sum_dedup", "krw", None, True,  False),
    ]),
]

_SUM_MGR_KEYS = set()
_SUM_DEDUP_KEYS = set()
for _, _, metrics in METRIC_TREE:
    for item in metrics:
        if item[2] == "sum_mgr":
            _SUM_MGR_KEYS.add(item[1])
        elif item[2] == "sum_dedup":
            _SUM_DEDUP_KEYS.add(item[1])
_SUM_MGR_KEYS.update([
    "f_b_total", "f_b_push_target", "f_b_push_fail",
    "f_c_total", "f_c_push_target", "f_c_push_fail",
    "effort_seconds_est", "coupon_cnt_w", "order_cnt_w",
])


# ============================================================
# 월별 집계 함수 (테이블 + Progress Board 공용)
# ============================================================
def build_monthly_summary_transposed(weekly_kpi: pd.DataFrame,
                                      view_type: str = "monthly") -> pd.DataFrame:
    df = weekly_kpi.copy()

    mgr_agg = {}
    for key in _SUM_MGR_KEYS:
        if key in df.columns:
            mgr_agg[key] = (key, "sum")
    mgr_monthly = df.groupby("monthly").agg(**mgr_agg).reset_index() if mgr_agg else pd.DataFrame()

    dedup_cols = [k for k in _SUM_DEDUP_KEYS if k in df.columns]
    if dedup_cols:
        deduped = df.drop_duplicates(subset=["weekly", "monthly"])[["weekly", "monthly"] + dedup_cols]
        dedup_monthly = deduped.groupby("monthly").agg({c: "sum" for c in dedup_cols}).reset_index()
        if not mgr_monthly.empty:
            mgr_monthly = mgr_monthly.merge(dedup_monthly, on="monthly", how="left")
        else:
            mgr_monthly = dedup_monthly

    m = mgr_monthly
    if m.empty:
        return pd.DataFrame()

    def _c(name):
        return m[name] if name in m.columns else pd.Series(0, index=m.index, dtype=float)

    def _safe_div(num, den):
        n = np.asarray(num, dtype=float)
        d = np.asarray(den, dtype=float)
        return np.where(d > 0, n / np.where(d == 0, 1, d), np.nan)

    m["_invalid_ratio"] = np.round(_safe_div(_c("lead_cnt") - _c("clean_lead_cnt"), _c("lead_cnt")) * 100, 1)
    m["_nonclosed_ratio"] = np.round(_safe_div(_c("clean_lead_cnt") - _c("consulted_cnt"), _c("clean_lead_cnt")) * 100, 1)
    m["_b_unmanaged"] = np.round(_safe_div(_c("f_b_push_target"), _c("f_b_total")) * 100, 1)
    m["_c_unmanaged"] = np.round(_safe_div(_c("f_c_push_target"), _c("f_c_total")) * 100, 1)
    m["_calls_per_lead_clean"] = np.round(_safe_div(_c("mobile_010_calls"), _c("lead_cnt")), 1)
    m["_calls_per_lead_total"] = np.round(_safe_div(_c("total_calls"), _c("clean_lead_cnt")), 1)
    m["_attempt_per_total"] = np.round(_safe_div(_c("total_calls"), _c("lead_cnt")), 2)
    m["_attempt_per_clean"] = np.round(_safe_div(_c("total_calls"), _c("clean_lead_cnt")), 2)
    m["_effort_hours"] = np.round(_c("effort_seconds_est") / 3600, 1)
    m["_cvr_total"] = np.round(_safe_div(_c("success_cnt"), _c("lead_cnt")) * 100, 1)
    m["_cvr_clean"] = np.round(_safe_div(_c("success_cnt"), _c("clean_lead_cnt")) * 100, 1)
    m["_cvr_consulted"] = np.round(_safe_div(_c("success_cnt"), _c("consulted_cnt")) * 100, 1)
    m["_coupon_ratio"] = np.round(_safe_div(_c("coupon_cnt_w"), _c("order_cnt_w")) * 100, 1)
    m["_refund_ratio"] = np.round(_safe_div(_c("total_refund_amount"), _c("paid_cost_w")) * 100, 1)

    months = sorted(m["monthly"].unique(), reverse=True)
    show_monthly = (view_type == "monthly")

    rows = []
    for cat1, cat2, metrics_list in METRIC_TREE:
        for item in metrics_list:
            label, key, agg, fmt, color_tag, in_monthly, in_manager = item
            if show_monthly and not in_monthly:
                continue
            if not show_monthly and not in_manager:
                continue
            row = {"대분류": cat1, "중분류": cat2, "지표": label, "_fmt": fmt, "_color": color_tag}
            for month in months:
                month_row = m[m["monthly"] == month]
                if month_row.empty or key not in m.columns:
                    row[month] = np.nan
                else:
                    row[month] = float(month_row[key].iloc[0])
            rows.append(row)

    return pd.DataFrame(rows)


# ============================================================
# 수식 사전 모달
# ============================================================
@st.dialog("수식 사전", width="large")
def _render_formula_dialog():
    left, right = st.columns(2)
    _tbl = 'style="width:100%; border-collapse:collapse; font-size:0.82rem; text-align:center;"'
    _th = 'style="padding:6px; border:1px solid #F0F0F0; background:#F0F0F0; font-weight:600;"'
    _td = 'style="padding:5px; border:1px solid #F0F0F0;"'
    _tr2 = 'style="background:#F9F9F9;"'

    with left:
        st.markdown("##### 파생 지표")
        st.markdown(f"""
<table {_tbl}>
<thead><tr><th {_th}>지표</th><th {_th}>수식</th><th {_th}>블록</th></tr></thead>
<tbody>
<tr><td {_td}><b>허수비율(%)</b></td><td {_td}>(리드전체 − 리드유효) ÷ 리드전체 × 100</td><td {_td}>A</td></tr>
<tr {_tr2}><td {_td}><b>미종결비율(%)</b></td><td {_td}>(리드유효 − 상담완료) ÷ 리드유효 × 100</td><td {_td}>A</td></tr>
<tr><td {_td}><b>리드당발신(전체)</b></td><td {_td}>발신전체 ÷ 리드전체</td><td {_td}>E ÷ A</td></tr>
<tr {_tr2}><td {_td}><b>리드당발신(유효)</b></td><td {_td}>발신전체 ÷ 리드유효</td><td {_td}>E ÷ A</td></tr>
<tr><td {_td}><b>CVR전체(%)</b></td><td {_td}>결제전환 ÷ 리드전체 × 100</td><td {_td}>A</td></tr>
<tr {_tr2}><td {_td}><b>CVR유효(%)</b></td><td {_td}>결제전환 ÷ 리드유효 × 100</td><td {_td}>A</td></tr>
<tr><td {_td}><b>CVR상담(%)</b></td><td {_td}>결제전환 ÷ 상담완료 × 100</td><td {_td}>A</td></tr>
<tr {_tr2}><td {_td}><b>환불비율(%)</b></td><td {_td}>환불금액 ÷ 결제금액 × 100</td><td {_td}>I ÷ H</td></tr>
</tbody></table>""", unsafe_allow_html=True)

    with right:
        st.markdown("##### 원본 지표")
        st.markdown(f"""
<table {_tbl}>
<thead><tr><th {_th}>지표</th><th {_th}>블록</th><th {_th}>시트 컬럼명</th></tr></thead>
<tbody>
<tr><td {_td}>리드(전체)</td><td {_td}>A</td><td {_td}><code>lead_cnt</code></td></tr>
<tr {_tr2}><td {_td}>리드(유효)</td><td {_td}>A</td><td {_td}><code>clean_lead_cnt</code></td></tr>
<tr><td {_td}>상담완료</td><td {_td}>A</td><td {_td}><code>consulted_cnt</code></td></tr>
<tr {_tr2}><td {_td}>결제전환</td><td {_td}>A</td><td {_td}><code>success_cnt</code></td></tr>
<tr><td {_td}>발신(전체)</td><td {_td}>E</td><td {_td}><code>total_calls</code></td></tr>
<tr {_tr2}><td {_td}>발신(정상)</td><td {_td}>E</td><td {_td}><code>mobile_010_calls</code></td></tr>
<tr><td {_td}>통화성공</td><td {_td}>E</td><td {_td}><code>success_calls</code></td></tr>
<tr {_tr2}><td {_td}>NewStdnt(a/c)</td><td {_td}>C</td><td {_td}><code>newstdt_ac</code></td></tr>
<tr><td {_td}>Booking(원)</td><td {_td}>B</td><td {_td}><code>net_booking_in_krw</code></td></tr>
<tr {_tr2}><td {_td}>결제금액(원)</td><td {_td}>H</td><td {_td}><code>paid_cost_w</code></td></tr>
<tr><td {_td}>환불금액(원)</td><td {_td}>I</td><td {_td}><code>total_refund_amount</code></td></tr>
</tbody></table>""", unsafe_allow_html=True)


# ============================================================
# 목표 설정 모달 (10개 KPI, 2행 × 5열 그리드)
# ============================================================
@st.dialog("목표 설정", width="large")
def _render_target_dialog():
    cur_month = _get_current_month()
    all_targets = _load_targets()
    month_targets = all_targets.get(cur_month, {})

    st.markdown(f"##### {cur_month} 목표치 입력")
    st.caption("각 KPI의 월간 목표를 입력하세요.")

    new_targets = {}
    # 2행 × 5열 그리드
    for row_start in range(0, len(PROGRESS_KPIS), 5):
        row_kpis = PROGRESS_KPIS[row_start:row_start + 5]
        cols = st.columns(len(row_kpis))
        for i, kpi in enumerate(row_kpis):
            with cols[i]:
                label = kpi["label"]
                fmt = kpi["fmt"]
                default = month_targets.get(label, 0)
                if fmt == "krw":
                    # Booking 목표: M(백만) 단위로 입력/저장
                    val = st.number_input(f"{label} (M)", value=int(default), step=1,
                                          format="%d", key=f"tgt_{label}",
                                          help="백만 단위로 입력 (예: 550 = 5.5억)")
                elif fmt == "pct":
                    val = st.number_input(label, value=float(default), step=0.5,
                                          format="%.1f", key=f"tgt_{label}")
                elif fmt == "ratio":
                    val = st.number_input(label, value=float(default), step=0.1,
                                          format="%.2f", key=f"tgt_{label}")
                else:
                    val = st.number_input(label, value=int(default), step=1,
                                          format="%d", key=f"tgt_{label}")
                new_targets[label] = val

    _s1, _s2, _s3 = st.columns([2, 1, 2])
    with _s2:
        _saved = st.button("저장", type="secondary", use_container_width=True)
    if _saved:
        all_targets[cur_month] = new_targets
        _save_targets(all_targets)
        st.success("저장 완료")
        st.rerun()


# ============================================================
# 프로그레스 보드 — MTD / 목표 / 도달율 / 남은수치 / 진척대비
# ============================================================
def _render_progress_board(weekly_kpi: pd.DataFrame):
    cur_month = _get_current_month()
    all_targets = _load_targets()
    month_targets = all_targets.get(cur_month, {})

    today = datetime.date.today()
    # 영업일 (카드 표시용)
    total_biz = _get_business_days_in_month(today.year, today.month)
    remain_biz = _get_remaining_business_days(today.year, today.month)
    # 달력일 기준 Month Progress = (오늘-1) / 월 총일수  (시트 공식과 동일)
    import calendar
    days_in_month = calendar.monthrange(today.year, today.month)[1]
    elapsed_days = max(today.day - 1, 1)
    remain_cal_days = days_in_month - today.day + 1  # 오늘 포함 남은 일수
    month_progress = (today.day - 1) / days_in_month  # 시트: =(DAY(TODAY())-1)/DAY(EOMONTH(TODAY(),0))

    # 전월 정보
    if today.month > 1:
        prev_year, prev_m = today.year, today.month - 1
    else:
        prev_year, prev_m = today.year - 1, 12
    prev_month_str = f"{prev_year}-{prev_m:02d}"

    df = weekly_kpi.copy()
    df_cur = df[df["monthly"] == cur_month] if "monthly" in df.columns else pd.DataFrame()
    df_prev = df[df["monthly"] == prev_month_str] if "monthly" in df.columns else pd.DataFrame()

    # ── 합산 헬퍼 ──
    def _sum_month(df_m, key, agg):
        if df_m.empty or key not in df_m.columns:
            return 0.0
        if agg == "sum_dedup":
            return float(df_m.drop_duplicates(subset=["weekly", "monthly"])[key].sum())
        return float(df_m[key].sum())

    # 파생지표 계산용 기초 데이터
    _base = ["lead_cnt", "clean_lead_cnt", "total_calls", "success_cnt", "success_calls"]
    cur_b = {k: _sum_month(df_cur, k, "sum_mgr") for k in _base}
    prev_b = {k: _sum_month(df_prev, k, "sum_mgr") for k in _base}

    def _derive(base_dict, key):
        if key == "_attempt_per_total":
            return round(base_dict["total_calls"] / base_dict["lead_cnt"], 2) if base_dict["lead_cnt"] > 0 else 0.0
        elif key == "_cvr_total":
            return round(base_dict["success_cnt"] / base_dict["lead_cnt"] * 100, 1) if base_dict["lead_cnt"] > 0 else 0.0
        elif key == "_cvr_clean":
            return round(base_dict["success_cnt"] / base_dict["clean_lead_cnt"] * 100, 1) if base_dict["clean_lead_cnt"] > 0 else 0.0
        return 0.0

    def _get_mtd(kpi):
        if kpi["agg"] == "derived":
            return _derive(cur_b, kpi["key"])
        return _sum_month(df_cur, kpi["key"], kpi["agg"])

    def _get_prev(kpi):
        if kpi["agg"] == "derived":
            return _derive(prev_b, kpi["key"])
        return _sum_month(df_prev, kpi["key"], kpi["agg"])

    # 포맷 헬퍼 (krw는 백만단위 "M" 표시)
    def _fmt_val(val, fmt, with_sign=False):
        prefix = "+" if with_sign and val > 0 else ""
        if fmt == "krw":
            m = val / 1_000_000
            if abs(m) >= 1:
                return f"{prefix}{m:,.0f}M"
            return f"{prefix}{int(val):,}"
        elif fmt == "pct":
            return f"{prefix}{val:.1f}%"
        elif fmt == "ratio":
            return f"{prefix}{val:.2f}"
        else:
            return f"{prefix}{int(val):,}"

    # ── 상단 헤더 바 ──
    month_num = cur_month.split("-")[1]
    pct_str = f"{month_progress * 100:.1f}"

    _hdr, _btn = st.columns([8, 1])
    with _hdr:
        st.markdown(f"**{cur_month} Progress Board** &nbsp;&nbsp;"
                    f'<span style="font-size:0.78rem; color:#999;">'
                    f'{month_num}/{today.day:02d} ({pct_str}%) · 영업일 {remain_biz}일 · 잔여 {remain_cal_days}일</span>',
                    unsafe_allow_html=True)
    with _btn:
        if st.button("목표 설정", type="secondary", use_container_width=True):
            _render_target_dialog()

    # ── KPI 계산 ──
    mtd_cache = {}
    kpi_data = []
    for kpi in PROGRESS_KPIS:
        actual = _get_mtd(kpi)
        prev_total = _get_prev(kpi)
        target = month_targets.get(kpi["label"], 0)
        mtd_cache[kpi["label"]] = actual
        kpi_data.append({**kpi, "actual": actual, "prev_total": prev_total, "target": target})

    has_any_target = any(d["target"] > 0 for d in kpi_data)

    # ── 상단 요약 카드 (st.columns로 렌더링) ──
    _card_indices = [0, 9, 7]  # 리드, Booking(M), CVR
    card_cols = st.columns(4)
    for i, idx in enumerate(_card_indices):
        ck = kpi_data[idx]
        a, t, fm = ck["actual"], ck["target"], ck["fmt"]
        v = _fmt_val(a, fm)
        if t > 0:
            # Booking(M) 목표는 백만 단위로 저장되므로 비교 시 환산
            t_raw = t * 1_000_000 if fm == "krw" else t
            p = a / t_raw * 100 if t_raw > 0 else 0
            p_color = "#0D9974" if p >= 100 else "#FF5500"
            sub_line = f'<span style="color:{p_color}; font-size:0.78rem; font-weight:600;">{p:.1f}%</span>'
        else:
            sub_line = ""
        with card_cols[i]:
            st.markdown(
                f'<div style="background:#FFF; border:1px solid #EAEAEA; border-radius:8px; padding:12px 16px;'
                f' min-height:100px; display:flex; flex-direction:column; justify-content:center; align-items:center; text-align:center;">'
                f'<div style="font-size:0.72rem; color:#999; margin-bottom:2px;">{ck["label"]}</div>'
                f'<div style="font-size:1.4rem; font-weight:700; color:#222; line-height:1.2;">{v}</div>'
                f'{sub_line}</div>',
                unsafe_allow_html=True)
    with card_cols[3]:
        mp_pct = month_progress * 100
        # 프로그레스 바 색상
        mp_color = "#FF5500" if mp_pct < 50 else "#FF8C4B" if mp_pct < 80 else "#0D9974"
        st.markdown(
            f'<div style="background:#FFF; border:1px solid #EAEAEA; border-radius:8px; padding:12px 16px;'
            f' min-height:100px; display:flex; flex-direction:column; justify-content:center;">'
            f'<div style="font-size:0.72rem; color:#999; margin-bottom:2px;">월 진행률</div>'
            f'<div style="font-size:1.4rem; font-weight:700; color:#222; line-height:1.2;">{mp_pct:.1f}%</div>'
            f'<div style="background:#F0F0F0;border-radius:3px;height:4px;margin:4px 0 6px;overflow:hidden;">'
            f'<div style="background:{mp_color};height:100%;width:{mp_pct:.0f}%;border-radius:3px;"></div></div>'
            f'<span style="font-size:0.68rem; color:#999;">영업일 {remain_biz}일 · 주말포함 {remain_cal_days}일 남음</span></div>',
            unsafe_allow_html=True)

    # ── KPI 테이블 (그룹 헤더 없이 플랫, 항상 6컬럼) ──
    rows_html = []

    for d in kpi_data:
        label, fmt = d["label"], d["fmt"]
        actual, target, prev_total = d["actual"], d["target"], d["prev_total"]
        has_target = target > 0

        # krw 목표는 M 단위로 저장됨 → 비교/계산 시 raw KRW로 환산
        target_raw = target * 1_000_000 if (fmt == "krw" and has_target) else target

        mtd_str = _fmt_val(actual, fmt)

        # 진척대비 (시트 공식 기준)
        # 패턴1 (건수/금액): MTD - 목표 × MonthProgress
        # 패턴2 (비율 pct/ratio): -남은수치 (= -(목표 - MTD) = MTD - 목표)
        if not has_target:
            delta_html = '<span style="color:#CCC;">—</span>'
        elif fmt in ("pct", "ratio"):
            delta = -(target_raw - actual)  # = -남은수치
            if abs(delta) < 0.005:
                delta_html = '<span style="color:#CCC;">—</span>'
            else:
                d_color = "#0D9974" if delta >= 0 else "#FB2D36"
                delta_html = f'<span style="color:{d_color}; font-weight:600;">{_fmt_val(delta, fmt, with_sign=True)}</span>'
        else:
            delta = actual - target_raw * month_progress
            if abs(delta) < 0.005:
                delta_html = '<span style="color:#CCC;">—</span>'
            else:
                d_color = "#0D9974" if delta >= 0 else "#FB2D36"
                delta_html = f'<span style="color:{d_color}; font-weight:600;">{_fmt_val(delta, fmt, with_sign=True)}</span>'

        # 목표/도달율/남은수치
        if has_target:
            pct = actual / target_raw * 100 if target_raw > 0 else 0
            daily_avg = actual / elapsed_days
            bar_pct = min(pct, 100)
            pct_c = "#0D9974" if pct >= 100 else "#FF5500" if pct >= 50 else "#FB2D36"
            bar_c = "#0D9974" if pct >= 100 else "#FF5500" if pct >= 50 else "#FFB899"
            if fmt == "krw":
                m_actual = actual / 1_000_000
                m_daily = daily_avg / 1_000_000
                rt = f"{m_actual:,.0f} ({pct:.1f}%, 일평균 {m_daily:.1f})"
            elif fmt == "pct":
                rt = f"{actual:.1f}% ({pct:.1f}%)"
            elif fmt == "ratio":
                rt = f"{actual:.2f} ({pct:.1f}%)"
            else:
                rt = f"{int(actual)} ({pct:.1f}%, 일평균 {daily_avg:.1f})"
            bar_td = (
                f'<div style="display:flex;align-items:center;gap:8px;">'
                f'<div style="flex:1;background:#F0F0F0;border-radius:3px;height:6px;overflow:hidden;">'
                f'<div style="background:{bar_c};height:100%;width:{bar_pct:.0f}%;border-radius:3px;"></div></div>'
                f'<span style="font-size:0.75rem;color:#888;white-space:nowrap;">{rt}</span>'
                f'<span style="font-weight:700;color:{pct_c};min-width:40px;text-align:right;">{pct:.1f}%</span>'
                f'</div>'
            )
            # 목표 표시: krw는 M 단위로 저장된 값 그대로 + "M" 표시
            tgt_str = f"{int(target)}M" if fmt == "krw" else _fmt_val(target_raw, fmt)
            rem_str = _fmt_val(target_raw - actual, fmt)
        else:
            tgt_str = ""
            bar_td = ""
            rem_str = ""

        _c = "padding:5px 12px; font-size:0.82rem; border-bottom:1px solid #F5F5F5;"
        rows_html.append(
            f'<tr>'
            f'<td style="{_c} color:#444; white-space:nowrap;">{label}</td>'
            f'<td style="{_c} text-align:right; color:#FF5500; font-weight:700; white-space:nowrap;">{mtd_str}</td>'
            f'<td style="{_c} text-align:right; color:#999;">{tgt_str}</td>'
            f'<td style="{_c} text-align:right; color:#555;">{rem_str}</td>'
            f'<td style="{_c} text-align:right;">{delta_html}</td>'
            f'<td style="{_c}">{bar_td}</td>'
            f'</tr>'
        )

    _th = "padding:5px 12px; font-size:0.72rem; color:#BBB; font-weight:600; border-bottom:1px solid #EAEAEA;"
    thead = (
        f'<th style="{_th} text-align:left; width:190px; min-width:190px;">지표</th>'
        f'<th style="{_th} text-align:right; width:80px; min-width:80px;">MTD</th>'
        f'<th style="{_th} text-align:right; width:80px; min-width:80px;">{month_num}월 목표</th>'
        f'<th style="{_th} text-align:right; width:90px; min-width:90px;">남은수치</th>'
        f'<th style="{_th} text-align:right; width:90px; min-width:90px;">진척대비</th>'
        f'<th style="{_th} text-align:left;">도달율</th>'
    )

    st.markdown(
        f'<div style="border:1px solid #EAEAEA; border-radius:10px; background:#FFF; margin-top:8px;">'
        f'<table style="width:100%;border-collapse:collapse;font-family:Pretendard,Calibri,sans-serif;">'
        f'<thead><tr>{thead}</tr></thead>'
        f'<tbody>{"".join(rows_html)}</tbody>'
        f'</table></div>',
        unsafe_allow_html=True)

    # ── What-if 시뮬레이터 — 구현 예정 (슬라이더 UI 미리보기) ──
    with st.expander("What-if 시뮬레이터 — 구현 예정", expanded=False):
        st.caption("추가 달성 건수를 조절하면 달성률이 실시간으로 변합니다. (현재 데이터 연동 미구현)")
        for row_start in range(0, len(PROGRESS_KPIS), 5):
            row_kpis = PROGRESS_KPIS[row_start:row_start + 5]
            sim_cols = st.columns(len(row_kpis))
            for i, kpi in enumerate(row_kpis):
                label, fmt = kpi["label"], kpi["fmt"]
                target = month_targets.get(label, 0)
                actual = mtd_cache.get(label, 0)
                with sim_cols[i]:
                    if fmt in ("pct", "ratio") or target <= 0:
                        st.markdown(f"<div style='text-align:center;font-size:0.72rem;color:#CCC;padding:4px;'>{label}</div>", unsafe_allow_html=True)
                        continue
                    max_v = max(int(target * 0.5), 50) if fmt != "krw" else max(int(target * 0.5), 10_000_000)
                    step = 1 if fmt != "krw" else 1_000_000
                    # 슬라이더는 조작 가능하지만 데이터 반영 없음 (껍데기)
                    add_val = st.slider(f"{label} 추가", 0, max_v, 0, step=step, key=f"sim_{label}")
                    # 시뮬레이션 결과 표시 (실제 데이터 미반영 — 로컬 계산만)
                    new_a = actual + add_val
                    new_r = max(0, target - new_a)
                    new_p = (new_a / target * 100) if target > 0 else 0
                    sc = "#0D9974" if new_p >= 100 else "#FF5500"
                    st.markdown(f'<div style="text-align:center;font-size:0.76rem;"><span style="color:{sc};font-weight:700;">{new_p:.0f}%</span> · 잔여 {_fmt_val(new_r, fmt)}</div>', unsafe_allow_html=True)


# ============================================================
# 메인 화면
# ============================================================
def main():
    st.title("CA Dashboard")

    try:
        weekly_kpi, monthly_mgr, source = load_data()
    except Exception as e:
        import traceback
        st.error(f"데이터 로드 실패: {e}")
        with st.expander("상세 에러 로그"):
            st.code(traceback.format_exc())
        st.stop()

    if source != "google_sheet":
        st.warning("샘플 데이터 사용 중 — 구글 시트 연결을 확인하세요")

    # ══ Progress Board ══
    _render_progress_board(weekly_kpi)

    # ── 섹션 간 여백 확보 (#12) ──
    st.markdown("<div style='height:32px'></div>", unsafe_allow_html=True)

    # ══ Monthly 종합 성과 ══
    _title_col, _btn_col = st.columns([8, 1])
    with _title_col:
        st.subheader("Monthly 종합 성과")
    with _btn_col:
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        # #11: 수식사전 — 작은 텍스트 링크 스타일
        _show_formula = st.button("수식 사전", key="btn_formula")

    summary = build_monthly_summary_transposed(weekly_kpi)

    if summary.empty:
        st.warning("데이터가 없어 지표 트리를 생성할 수 없습니다.")
    else:
        meta_cols = {"대분류", "중분류", "지표", "_fmt", "_color"}
        months = [c for c in summary.columns if c not in meta_cols]

        # ── Sticky 컬럼 오프셋 ──
        W_CAT1 = 72
        W_CAT2 = 70
        W_METRIC = 200
        STK1 = f"position:sticky; left:0; z-index:3;"
        STK2 = f"position:sticky; left:{W_CAT1}px; z-index:3;"
        STK3 = f"position:sticky; left:{W_CAT1 + W_CAT2}px; z-index:2;"
        STK1_H = f"position:sticky; left:0; z-index:5;"
        STK2_H = f"position:sticky; left:{W_CAT1}px; z-index:5;"
        STK3_H = f"position:sticky; left:{W_CAT1 + W_CAT2}px; z-index:4;"

        # ── HTML 테이블 빌드 ──
        html_rows = []
        prev_cat1 = None
        prev_cat2 = None

        for idx, row in summary.iterrows():
            cat1, cat2 = row["대분류"], row["중분류"]
            label, fmt, color_tag = row["지표"], row["_fmt"], row["_color"]

            # 중분류 그룹 변경 시 상단 굵은 보더 (#10)
            group_changed = (cat1 != prev_cat1) or (cat2 != prev_cat2)
            top_border = "border-top:2px solid #D0D0D0;" if group_changed and prev_cat1 is not None else ""

            # 행 배경 (#8: mkt_tag만 초록)
            row_bg = "#EDF7ED" if color_tag == "mkt_tag" else "#FFFFFF"

            # 대분류 셀
            if cat1 != prev_cat1:
                n_cat1 = len(summary[summary["대분류"] == cat1])
                html_rows.append(
                    f'<tr>'
                    f'<td rowspan="{n_cat1}" style="'
                    f'{STK1} background:#F0F0F0; color:#222; font-weight:700; '
                    f'text-align:center; padding:8px 6px; font-size:0.82rem; '
                    f'border-right:1px solid #E0E0E0; border-bottom:1px solid #E0E0E0; '
                    f'vertical-align:middle; width:{W_CAT1}px; min-width:{W_CAT1}px; max-width:{W_CAT1}px; {top_border}">'
                    f'{cat1}</td>'
                )
                prev_cat1 = cat1
                prev_cat2 = None
            else:
                html_rows.append(f'<tr>')

            # 중분류 셀
            if cat2 != prev_cat2:
                n_cat2 = len(summary[(summary["대분류"] == cat1) & (summary["중분류"] == cat2)])
                html_rows.append(
                    f'<td rowspan="{n_cat2}" style="'
                    f'{STK2} background:#F9F9F9; color:#5D5D5D; font-weight:600; '
                    f'text-align:center; padding:6px 6px; font-size:0.78rem; '
                    f'border-right:1px solid #F0F0F0; border-bottom:1px solid #F0F0F0; '
                    f'vertical-align:middle; width:{W_CAT2}px; min-width:{W_CAT2}px; max-width:{W_CAT2}px; {top_border}">'
                    f'{cat2}</td>'
                )
                prev_cat2 = cat2

            # 지표 셀
            html_rows.append(
                f'<td style="{STK3} background:{row_bg}; padding:5px 10px; '
                f'font-size:0.8rem; color:#222; text-align:left; '
                f'border-right:1px solid #F0F0F0; border-bottom:1px solid #F0F0F0; '
                f'white-space:nowrap; width:{W_METRIC}px; min-width:{W_METRIC}px;">'
                f'{label}</td>'
            )

            # 월별 데이터 셀
            for month in months:
                val = row[month]
                if pd.isna(val):
                    cell_str = "-"
                elif fmt == "pct":
                    cell_str = f"{val:.1f}%"
                elif fmt == "ratio":
                    cell_str = f"{val:.1f}"
                elif fmt == "hours":
                    cell_str = f"{val:.1f}h"
                elif fmt == "krw":
                    cell_str = f"{int(val):,}"
                else:
                    cell_str = f"{int(val):,}"

                html_rows.append(
                    f'<td style="background:{row_bg}; text-align:center; '
                    f'padding:5px 10px; font-size:0.8rem; color:#222; '
                    f'border-right:1px solid #F0F0F0; border-bottom:1px solid #F0F0F0; '
                    f'white-space:nowrap; min-width:85px;">'
                    f'{cell_str}</td>'
                )
            html_rows.append('</tr>')

        # 헤더 (#6: #F0F0F0 배경)
        _hdr_th = "background:#F0F0F0; color:#222; font-weight:700; padding:8px 10px; font-size:0.8rem; text-align:center; border-bottom:2px solid #D0D0D0; white-space:nowrap;"
        month_ths = "".join(
            f'<th style="{_hdr_th} min-width:85px;">{m}</th>' for m in months
        )

        full_html = f"""
        <div style="overflow-x:auto; border:1px solid #E0E0E0; border-radius:8px; background:#FFF;">
        <table style="width:max-content; border-collapse:collapse; font-family:'Pretendard','Calibri',sans-serif;">
        <thead><tr>
            <th style="{_hdr_th} {STK1_H} width:{W_CAT1}px; min-width:{W_CAT1}px;">대분류</th>
            <th style="{_hdr_th} {STK2_H} width:{W_CAT2}px; min-width:{W_CAT2}px;">중분류</th>
            <th style="{_hdr_th} {STK3_H} width:{W_METRIC}px; min-width:{W_METRIC}px;">지표</th>
            {month_ths}
        </tr></thead>
        <tbody>{"".join(html_rows)}</tbody>
        </table></div>"""

        st.markdown(full_html, unsafe_allow_html=True)

    if _show_formula:
        _render_formula_dialog()


if __name__ == "__main__":
    main()
