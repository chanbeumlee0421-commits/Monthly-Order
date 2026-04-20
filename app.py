import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import date, timedelta

st.set_page_config(
    page_title="경보제약 월별 주문 분석",
    page_icon="💊",
    layout="wide"
)

st.markdown("""
<style>
    .stMetric { background-color: white; padding: 16px; border-radius: 10px; border: 1px solid #e9ecef; }
    div[data-testid="stMetricValue"] { font-size: 1.6rem; font-weight: 700; }
</style>
""", unsafe_allow_html=True)

st.title("경보제약 동물병원 월별 주문 분석")
st.caption("Raw 탭 엑셀 파일을 업로드하면 자동으로 분석됩니다.")
st.divider()

uploaded = st.file_uploader("📂 엑셀 파일 업로드 (경보제약_dashboard.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("👆 위에서 엑셀 파일을 업로드해 주세요.")
    st.stop()

@st.cache_data
def load_data(file):
    df = pd.read_excel(file, sheet_name="Raw")
    df = df[df["유통"] == "직거래"].copy()
    df["매출일"] = pd.to_datetime(df["매출일(배송완료일)"], errors="coerce")
    df = df.dropna(subset=["매출일"])
    df["매출수량"] = pd.to_numeric(df["매출수량"], errors="coerce").fillna(0)
    df["매출액"] = pd.to_numeric(df["매출액(vat 포함)"], errors="coerce").fillna(0)
    df["제품명"] = df["제품명"].fillna("기타")
    df["담당자"] = df["담당자"].fillna("미지정")
    return df

df = load_data(uploaded)

all_hospitals = sorted(df["거래처명"].dropna().unique().tolist())
all_managers  = sorted(df["담당자"].dropna().unique().tolist())

# 우선순위 제품명 순서
PRIORITY = [
    "티스템 펫 2mL",
    "티스템 크림펫 30g",
    "티스템 크림펫 10g * 10개입",
    "레나크린 120캡슐",
    "레나톡스캅",
    "바이오플로라 300g",
    "벳에이다 플러스",
    "벳에이다 테이스티",
    "벳에이다 하이포",
    "벳에이다 카디오",
    "제로디 55g (5.5g x 10ea)",
    "제로디 55g(5.5gx10ea)",
    "모보플렉스 2g 정 X 30정/통",
    "이지앱",
    "듀라하트 SR-3 주사액(목시덱틴) 6mL 6V (분말부3V, 희석액 3V)",
]

all_products_raw = df["제품명"].dropna().unique().tolist()
priority_products = [p for p in PRIORITY if p in all_products_raw]
rest_products = sorted([p for p in all_products_raw if p not in PRIORITY])
all_products = priority_products + rest_products

min_date = df["매출일"].min().date()
max_date = df["매출일"].max().date()
today    = date.today()

with st.sidebar:
    st.header("🔍 필터 설정")

    st.subheader("📅 주문 기간")
    d_col1, d_col2, d_col3 = st.columns([5, 5, 3])
    with d_col1:
        start_date = st.date_input("시작일", value=date(2024, 1, 1),
                                   min_value=min_date, max_value=max_date)
    with d_col3:
        use_today = st.checkbox("오늘", value=False)
    with d_col2:
        if use_today:
            end_date = min(today, max_date)
            st.date_input("종료일", value=end_date, disabled=True)
        else:
            end_date = st.date_input("종료일", value=date(2024, 12, 31),
                                     min_value=min_date, max_value=max_date)

    st.subheader("👤 담당자")
    selected_managers = st.multiselect(
        "담당자 선택 (미선택 시 전체)",
        options=all_managers,
        placeholder="담당자를 선택하세요..."
    )

    st.subheader("🏥 동물병원")
    selected_hospitals = st.multiselect(
        "병원 선택 (미선택 시 전체)",
        options=all_hospitals,
        placeholder="병원명을 검색하거나 선택하세요..."
    )

    # 기간 + 담당자 + 병원 조건 모두 반영해서 실제 거래된 제품만 표시
    filtered_df = df[
        (df["매출일"].dt.date >= start_date) &
        (df["매출일"].dt.date <= end_date)
    ].copy()
    if selected_managers:
        filtered_df = filtered_df[filtered_df["담당자"].isin(selected_managers)]
    if selected_hospitals:
        filtered_df = filtered_df[filtered_df["거래처명"].isin(selected_hospitals)]
    available_products_raw = filtered_df["제품명"].dropna().unique().tolist()
    available_products = [p for p in all_products if p in available_products_raw]

    st.subheader("💊 품목")
    col_all, col_syringe = st.columns(2)
    with col_all:
        select_all = st.checkbox("전체 선택", value=True)
    with col_syringe:
        exclude_syringe = st.checkbox("주사기 제외", value=False)

    syringe_keywords = ["syringe", "주사기"]
    non_syringe = [
        p for p in available_products
        if not any(k in p.lower() for k in syringe_keywords)
    ]
    base_list = non_syringe if exclude_syringe else available_products

    if select_all:
        selected_products = base_list
    else:
        selected_products = st.multiselect(
            "제품명 선택",
            options=base_list,
            default=base_list
        )

# ── 필터링 ────────────────────────────────────────────
mask = (
    (df["매출일"].dt.date >= start_date) &
    (df["매출일"].dt.date <= end_date) &
    (df["제품명"].isin(selected_products))
)
if selected_managers:
    mask &= df["담당자"].isin(selected_managers)
if selected_hospitals:
    mask &= df["거래처명"].isin(selected_hospitals)

fdf = df[mask].copy()

if fdf.empty:
    st.warning("⚠️ 선택한 조건에 맞는 데이터가 없습니다. 필터를 조정해주세요.")
    st.stop()

delta_days = (end_date - start_date).days + 1
months = max(delta_days / 30.44, 0.1)

# ── 집계 ─────────────────────────────────────────────
agg = (
    fdf.groupby(["거래처명", "제품명"])
    .agg(총수량=("매출수량", "sum"), 총금액=("매출액", "sum"))
    .reset_index()
)
agg["월평균수량"] = agg["총수량"] / months

hosp_agg = (
    fdf.groupby("거래처명")
    .agg(총수량=("매출수량", "sum"), 총금액=("매출액", "sum"))
    .reset_index()
)
hosp_agg["월평균수량"] = hosp_agg["총수량"] / months

# ── 요약 지표 ────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.metric("분석 병원 수", f"{fdf['거래처명'].nunique():,}개")
with c2:
    st.metric("총 주문 건수", f"{len(fdf):,}건")
with c3:
    st.metric("총 매출수량", f"{fdf['매출수량'].sum():,.0f}개")
with c4:
    st.metric("총 매출액", f"₩{fdf['매출액'].sum()/1e6:,.1f}M")

st.divider()

# ── 메인 테이블 ──────────────────────────────────────
st.subheader("📊 병원별 × 제품별 월평균 주문수량")
st.caption(f"분석 기간: {start_date} ~ {end_date} ({months:.1f}개월 기준)")

pivot = agg.pivot_table(
    index="거래처명",
    columns="제품명",
    values="월평균수량",
    fill_value=0
)

# 선택된 제품만, 우선순위 순서 유지
show_cols = [p for p in selected_products if p in pivot.columns]
pivot = pivot[show_cols]

pivot = pivot.merge(
    hosp_agg[["거래처명", "총수량", "총금액"]],
    left_index=True, right_on="거래처명"
).set_index("거래처명")

pivot = pivot.sort_values("총수량", ascending=False)

def color_cell(val):
    if isinstance(val, (int, float)):
        if val >= 10:
            return "background-color: #d4edda; color: #155724; font-weight: 600"
        elif val >= 3:
            return "background-color: #fff3cd; color: #856404;"
        elif val > 0:
            return "background-color: #f8f9fa; color: #495057;"
    return ""

display_pivot = pivot.copy()
display_pivot.columns = [
    "[총구매수량]" if c == "총수량"
    else "[총구매액]" if c == "총금액"
    else c
    for c in display_pivot.columns
]

fmt2 = {}
for col in display_pivot.columns:
    if col == "[총구매수량]":
        fmt2[col] = "{:,.0f}"
    elif col == "[총구매액]":
        fmt2[col] = "₩{:,.0f}"
    else:
        fmt2[col] = "{:.2f}"

styled = (
    display_pivot.style
    .format(fmt2)
    .map(color_cell, subset=[c for c in display_pivot.columns if c not in ["[총구매수량]", "[총구매액]"]])
    .set_properties(**{"text-align": "right", "font-size": "13px"})
    .set_table_styles([
        {"selector": "th", "props": [("background-color", "#f1f3f5"), ("font-weight", "600"), ("font-size", "12px"), ("text-align", "center")]},
        {"selector": "td:first-child", "props": [("font-weight", "500"), ("text-align", "left"), ("min-width", "180px")]},
    ])
)

st.dataframe(styled, use_container_width=True, height=500)

st.markdown("""
<div style="display:flex; gap:20px; font-size:12px; margin-top:4px;">
  <span>🟢 <b>월 10개 이상</b> — 고빈도 구매</span>
  <span>🟡 <b>월 3~9개</b> — 중빈도 구매</span>
  <span>⬜ <b>월 3개 미만</b> — 저빈도 구매</span>
</div>
""", unsafe_allow_html=True)

st.divider()

# ── 최근 1년 이탈 감지 ───────────────────────────────
st.subheader("⚠️ 최근 1년 구매 이탈 병원")
st.caption(f"과거에 구매했지만 최근 1년({(max_date - timedelta(days=365)).strftime('%Y-%m-%d')} 이후) 주문이 없는 병원 × 제품 조합")

cutoff = max_date - timedelta(days=365)
recent_mask   = df["매출일"].dt.date >= cutoff
recent_buyers = set(zip(df[recent_mask]["거래처명"], df[recent_mask]["제품명"]))

base_df = df.copy()
if selected_managers:
    base_df = base_df[base_df["담당자"].isin(selected_managers)]
if selected_hospitals:
    base_df = base_df[base_df["거래처명"].isin(selected_hospitals)]
base_df = base_df[base_df["제품명"].isin(selected_products)]

past_buyers = set(zip(base_df["거래처명"], base_df["제품명"]))
churned = past_buyers - recent_buyers

if churned:
    churn_df = pd.DataFrame(list(churned), columns=["병원명", "제품명"])
    last_order = (
        df.groupby(["거래처명", "제품명"])["매출일"]
        .max().reset_index()
        .rename(columns={"거래처명": "병원명", "매출일": "마지막_주문일"})
    )
    churn_df = churn_df.merge(last_order, on=["병원명", "제품명"], how="left")
    churn_df["마지막 주문일"] = churn_df["마지막_주문일"].dt.strftime("%Y-%m-%d")
    churn_df = churn_df[["병원명", "제품명", "마지막 주문일"]].sort_values("마지막 주문일")
    st.warning(f"총 **{len(churn_df)}개** 병원×제품 조합이 최근 1년간 주문 없음")
    st.dataframe(churn_df, use_container_width=True, height=300)
else:
    st.success("✅ 최근 1년 이탈 병원이 없습니다!")

st.divider()

# ── TOP 20 병원 ──────────────────────────────────────
st.subheader("🏆 TOP 20 병원 (총 구매수량 기준)")

top20 = hosp_agg.nlargest(20, "총수량").copy()
top20["월평균"] = (top20["총수량"] / months).round(1)

fig2 = px.bar(
    top20.sort_values("총수량"),
    x="총수량", y="거래처명", orientation="h",
    template="plotly_white", color="총수량",
    color_continuous_scale="Blues", text="월평균",
    labels={"총수량": "총 구매수량", "거래처명": ""}
)
fig2.update_traces(texttemplate="%{text}개/월", textposition="outside")
fig2.update_layout(
    height=600, showlegend=False,
    coloraxis_showscale=False,
    margin=dict(l=0, r=60, t=10, b=0)
)
st.plotly_chart(fig2, use_container_width=True)
