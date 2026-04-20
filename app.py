import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import date, timedelta

st.set_page_config(
    page_title="경보제약 월별 주문 분석",
    page_icon="💊",
    layout="wide"
)

st.markdown("""
<style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: white; padding: 16px; border-radius: 10px; border: 1px solid #e9ecef; }
    .insight-box {
        background: linear-gradient(135deg, #e8f4fd, #f0f8ff);
        border-left: 4px solid #2196F3;
        padding: 16px 20px;
        border-radius: 0 10px 10px 0;
        margin: 16px 0;
    }
    div[data-testid="stMetricValue"] { font-size: 1.6rem; font-weight: 700; }
    .section-title { font-size: 1rem; font-weight: 600; color: #495057; margin-bottom: 8px; }
</style>
""", unsafe_allow_html=True)

# ── 헤더 ──────────────────────────────────────────────
st.title("💊 경보제약 동물병원 월별 주문 분석")
st.caption("Raw 탭 엑셀 파일을 업로드하면 자동으로 분석됩니다.")
st.divider()

# ── 파일 업로드 ────────────────────────────────────────
uploaded = st.file_uploader("📂 엑셀 파일 업로드 (경보제약_dashboard.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("👆 위에서 엑셀 파일을 업로드해 주세요.")
    st.stop()

# ── 데이터 로드 ────────────────────────────────────────
@st.cache_data
def load_data(file):
    df = pd.read_excel(file, sheet_name="Raw")
    df = df[df["유통"] == "직거래"].copy()
    df["매출일"] = pd.to_datetime(df["매출일(배송완료일)"], errors="coerce")
    df = df.dropna(subset=["매출일"])
    df["매출수량"] = pd.to_numeric(df["매출수량"], errors="coerce").fillna(0)
    df["매출액"] = pd.to_numeric(df["매출액(vat 포함)"], errors="coerce").fillna(0)
    df["품명요약"] = df["품명요약"].fillna("기타")
    return df

df = load_data(uploaded)

all_hospitals = sorted(df["거래처명"].dropna().unique().tolist())
all_products  = sorted(df["품명요약"].dropna().unique().tolist())
min_date = df["매출일"].min().date()
max_date = df["매출일"].max().date()

# ── 필터 사이드바 ──────────────────────────────────────
with st.sidebar:
    st.header("🔍 필터 설정")

    st.subheader("📅 주문 기간")
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("시작일", value=date(2024, 1, 1), min_value=min_date, max_value=max_date)
    with col2:
        end_date = st.date_input("종료일", value=date(2024, 12, 31), min_value=min_date, max_value=max_date)

    st.subheader("🏥 동물병원")
    hospital_search = st.text_input("병원명 검색", placeholder="검색어 입력...")
    filtered_hospitals = [h for h in all_hospitals if hospital_search.lower() in h.lower()] if hospital_search else all_hospitals
    selected_hospitals = st.multiselect(
        "병원 선택 (미선택 시 전체)",
        options=filtered_hospitals,
        placeholder="병원을 선택하세요..."
    )

    st.subheader("💊 품목")
    select_all = st.checkbox("전체 선택", value=True)
    if select_all:
        selected_products = all_products
    else:
        selected_products = st.multiselect(
            "품목 선택",
            options=all_products,
            default=all_products
        )

# ── 데이터 필터링 ──────────────────────────────────────
mask = (
    (df["매출일"].dt.date >= start_date) &
    (df["매출일"].dt.date <= end_date) &
    (df["품명요약"].isin(selected_products))
)
if selected_hospitals:
    mask &= df["거래처명"].isin(selected_hospitals)

fdf = df[mask].copy()

if fdf.empty:
    st.warning("⚠️ 선택한 조건에 맞는 데이터가 없습니다. 필터를 조정해주세요.")
    st.stop()

# ── 기간 계산 (개월 수) ────────────────────────────────
delta_days = (end_date - start_date).days + 1
months = max(delta_days / 30.44, 0.1)

# ── 집계 ───────────────────────────────────────────────
# 병원 × 품목 집계
agg = (
    fdf.groupby(["거래처명", "품명요약"])
    .agg(총수량=("매출수량", "sum"), 총금액=("매출액", "sum"))
    .reset_index()
)
agg["월평균수량"] = agg["총수량"] / months

# 병원별 전체 집계
hosp_agg = (
    fdf.groupby("거래처명")
    .agg(총수량=("매출수량", "sum"), 총금액=("매출액", "sum"))
    .reset_index()
)
hosp_agg["월평균수량"] = hosp_agg["총수량"] / months

# ── 요약 지표 ──────────────────────────────────────────
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

# ── 메인 테이블 ────────────────────────────────────────
st.subheader("📊 병원별 × 품목별 월평균 주문수량")
st.caption(f"분석 기간: {start_date} ~ {end_date} ({months:.1f}개월 기준)")

# 피벗 테이블 생성
pivot = agg.pivot_table(
    index="거래처명",
    columns="품명요약",
    values="월평균수량",
    fill_value=0
)

# 선택한 품목만 표시
show_cols = [c for c in selected_products if c in pivot.columns]
pivot = pivot[show_cols]

# 총수량, 총금액 추가
pivot = pivot.merge(
    hosp_agg[["거래처명", "총수량", "총금액"]].rename(columns={"거래처명": "거래처명"}),
    left_index=True, right_on="거래처명"
).set_index("거래처명")

# 정렬
pivot = pivot.sort_values("총수량", ascending=False)

# 스타일 함수
def color_cell(val):
    if isinstance(val, (int, float)):
        if val >= 10:
            return "background-color: #d4edda; color: #155724; font-weight: 600"
        elif val >= 3:
            return "background-color: #fff3cd; color: #856404;"
        elif val > 0:
            return "background-color: #f8f9fa; color: #495057;"
    return ""

# 포맷 딕셔너리
fmt = {}
for col in show_cols:
    fmt[col] = "{:.2f}"
fmt["총수량"] = "{:,.0f}"
fmt["총금액"] = "₩{:,.0f}"

# 총금액, 총수량 컬럼 이름 변경해서 표시
display_pivot = pivot.copy()
display_pivot.columns = [
    f"[총구매수량]" if c == "총수량"
    else f"[총구매액]" if c == "총금액"
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

# 범례
st.markdown("""
<div style="display:flex; gap:20px; font-size:12px; margin-top:4px;">
  <span>🟢 <b>월 10개 이상</b> — 고빈도 구매</span>
  <span>🟡 <b>월 3~9개</b> — 중빈도 구매</span>
  <span>⬜ <b>월 3개 미만</b> — 저빈도 구매</span>
</div>
""", unsafe_allow_html=True)

st.divider()

# ── 인사이트 — 이탈 감지 ──────────────────────────────
st.subheader("💡 인사이트: 최근 구매 이탈 감지")
st.caption("예전엔 구매했지만 최근 3개월 동안 주문이 없는 병원 × 품목 조합")

cutoff = max_date - timedelta(days=90)
recent_mask = df["매출일"].dt.date >= cutoff
recent_buyers = set(zip(df[recent_mask]["거래처명"], df[recent_mask]["품명요약"]))

# 필터 범위 내에서 과거에 구매한 병원×품목
if selected_hospitals:
    past_df = df[df["거래처명"].isin(selected_hospitals) & df["품명요약"].isin(selected_products)]
else:
    past_df = df[df["품명요약"].isin(selected_products)]

past_buyers = set(zip(past_df["거래처명"], past_df["품명요약"]))
churned = past_buyers - recent_buyers

if churned:
    churn_df = pd.DataFrame(list(churned), columns=["병원명", "품목"])
    # 마지막 구매일 추가
    last_order = (
        df.groupby(["거래처명", "품명요약"])["매출일"]
        .max()
        .reset_index()
        .rename(columns={"거래처명": "병원명", "품명요약": "품목", "매출일": "마지막_주문일"})
    )
    churn_df = churn_df.merge(last_order, on=["병원명", "품목"], how="left")
    churn_df["마지막_주문일"] = churn_df["마지막_주문일"].dt.strftime("%Y-%m-%d")
    churn_df = churn_df.sort_values("마지막_주문일")

    st.markdown(f"""
    <div class="insight-box">
    📌 <b>총 {len(churn_df)}개 조합</b>이 최근 3개월({cutoff.strftime('%Y-%m-%d')} 이후) 주문 없음<br>
    이 병원들에 재방문 또는 프로모션 연락을 고려해보세요.
    </div>
    """, unsafe_allow_html=True)

    st.dataframe(
        churn_df.rename(columns={"마지막_주문일": "마지막 주문일"}),
        use_container_width=True,
        height=300
    )
else:
    st.success("✅ 최근 3개월 이탈 병원이 없습니다!")

st.divider()

# ── 품목별 월별 트렌드 차트 ────────────────────────────
st.subheader("📈 품목별 월별 매출수량 트렌드")

monthly = (
    fdf.groupby([pd.Grouper(key="매출일", freq="ME"), "품명요약"])
    ["매출수량"].sum()
    .reset_index()
)
monthly["월"] = monthly["매출일"].dt.strftime("%Y-%m")

if not monthly.empty:
    fig = px.line(
        monthly,
        x="월", y="매출수량", color="품명요약",
        markers=True,
        template="plotly_white",
        color_discrete_sequence=px.colors.qualitative.Set2
    )
    fig.update_layout(
        height=400,
        legend_title="품목",
        xaxis_title="",
        yaxis_title="매출수량 (개)",
        hovermode="x unified",
        margin=dict(l=0, r=0, t=20, b=0)
    )
    st.plotly_chart(fig, use_container_width=True)

st.divider()

# ── TOP 병원 ───────────────────────────────────────────
st.subheader("🏆 TOP 20 병원 (총 구매수량 기준)")

top20 = hosp_agg.nlargest(20, "총수량").copy()
top20["월평균"] = (top20["총수량"] / months).round(1)

fig2 = px.bar(
    top20.sort_values("총수량"),
    x="총수량", y="거래처명",
    orientation="h",
    template="plotly_white",
    color="총수량",
    color_continuous_scale="Blues",
    text="월평균",
    labels={"총수량": "총 구매수량", "거래처명": ""}
)
fig2.update_traces(texttemplate="%{text}개/월", textposition="outside")
fig2.update_layout(
    height=600,
    showlegend=False,
    coloraxis_showscale=False,
    margin=dict(l=0, r=60, t=10, b=0)
)
st.plotly_chart(fig2, use_container_width=True)

# ── 푸터 ───────────────────────────────────────────────
st.divider()
st.caption("경보제약 동물사업부 | Raw 탭 기준 직거래 거래처만 집계 | VAT 포함 매출액 기준")
