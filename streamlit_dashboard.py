
# streamlit_dashboard.py (fixed)
import io
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st
from plotly.subplots import make_subplots
from pathlib import Path

import streamlit as st

try:
    import plotly.graph_objects as go
    import plotly.express as px
except ModuleNotFoundError:
    st.error("Plotly가 설치되지 않았습니다. `pip install plotly` 또는 `pip install -r requirements.txt` 실행 후 앱을 다시 시작하세요.")
    st.stop()
  # ensure Path is available

st.set_page_config(page_title="엑셀 대시보드", layout="wide")

st.title("📊 엑셀 대시보드 (시트별 그래프 모음)")
st.caption("업로드한 엑셀 파일의 6개 시트를 한 번에 시각화합니다.")

# ---- File input ----
DEFAULT_PATH = Path("/mnt/data/0. 그래프_최종_과제용.xlsx")

uploaded = st.file_uploader("엑셀 파일(.xlsx) 업로드", type=["xlsx"])

# 샘플 파일 토글 (샘플 파일이 존재할 때만 노출)
use_sample = False
if uploaded is None and DEFAULT_PATH.exists():
    use_sample = st.toggle("샘플 파일 사용 (업로드 대신)", value=True, help="현재 세션에 제공된 샘플 엑셀 파일을 사용합니다.")

def get_file_like(uploaded, use_sample):
    """Return a file-like object (BytesIO) for Excel loading or stop the app with a message."""
    if uploaded is not None:
        # Uploaded file is already a file-like object
        return uploaded
    if use_sample and DEFAULT_PATH.exists():
        # Read sample file into BytesIO so no file handle remains open
        data = DEFAULT_PATH.read_bytes()
        return io.BytesIO(data)
    # No file available -> stop with message
    st.warning("엑셀 파일을 업로드하거나, '샘플 파일 사용'을 켜주세요.")
    st.stop()

file_like = get_file_like(uploaded, use_sample)

@st.cache_data(show_spinner=False)
def load_excel(file_like_obj):
    xls = pd.ExcelFile(file_like_obj)
    sheets = xls.sheet_names
    needed = ["바차트_히스토그램", "시계열차트", "파이차트", "산점도", "파레토차트", "버블차트"]
    missing = [s for s in needed if s not in sheets]
    if missing:
        raise ValueError(f"다음 시트가 누락되었습니다: {missing}")

    bar_df = pd.read_excel(xls, sheet_name="바차트_히스토그램")
    time_df = pd.read_excel(xls, sheet_name="시계열차트")
    pie_df = pd.read_excel(xls, sheet_name="파이차트")
    scatter_df = pd.read_excel(xls, sheet_name="산점도")
    pareto_df = pd.read_excel(xls, sheet_name="파레토차트")
    bubble_df = pd.read_excel(xls, sheet_name="버블차트")

    # 전처리
    bar_df["월"] = pd.to_datetime(bar_df["월"]).dt.strftime("%Y-%m")
    time_df["월"] = pd.to_datetime(time_df["월"]).dt.strftime("%Y-%m")
    # 파이차트 컬럼명 정규화 (첫 두 열을 '제품', '1분기 매출'로)
    pie_cols = list(pie_df.columns[:2])
    pie_df = pie_df.rename(columns={pie_cols[0]: "제품", pie_cols[1]: "1분기 매출"})

    # Pareto 계산
    pareto_sorted = pareto_df.sort_values("매출", ascending=False).reset_index(drop=True)
    pareto_sorted["누적 비율"] = pareto_sorted["매출"].cumsum() / pareto_sorted["매출"].sum() * 100

    # Bubble size 기준
    max_customer = bubble_df["고객 수"].max()
    desired_max_px = 60  # 최대 버블 픽셀 크기
    sizeref = (2.0 * max_customer) / (desired_max_px ** 2) if max_customer and desired_max_px else 1.0

    return {
        "bar_df": bar_df,
        "time_df": time_df,
        "pie_df": pie_df,
        "scatter_df": scatter_df,
        "pareto_sorted": pareto_sorted,
        "bubble_df": bubble_df,
        "sizeref": sizeref
    }

# 안전하게 로드
try:
    dfs = load_excel(file_like)
except Exception as e:
    st.error(f"파일을 불러오면서 오류가 발생했습니다: {e}")
    st.stop()

# 해체할당
bar_df = dfs["bar_df"]
time_df = dfs["time_df"]
pie_df = dfs["pie_df"]
scatter_df = dfs["scatter_df"]
pareto_sorted = dfs["pareto_sorted"]
bubble_df = dfs["bubble_df"]
sizeref = dfs["sizeref"]

# ---- Sidebar Filters ----
st.sidebar.header("⚙️ 옵션")
all_products = list(time_df.columns[1:])
sel_products = st.sidebar.multiselect("시계열 표시 제품 선택", options=all_products, default=all_products)

# ---- 1) 바차트 ----
with st.container():
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("월별 총 매출 (바차트)")
        fig_bar = go.Figure(data=[go.Bar(
            x=bar_df["월"],
            y=bar_df["총 매출"],
            marker=dict(line=dict(color="rgba(0,0,0,0.25)", width=0.5))
        )])
        fig_bar.update_layout(
            margin=dict(l=40, r=20, t=40, b=40),
            xaxis_title="월", yaxis_title="총 매출",
            height=380
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    # ---- 2) 시계열 ----
    with c2:
        st.subheader("제품별 월별 매출 추세 (시계열)")
        fig_time = go.Figure()
        for col in all_products:
            if col in sel_products:
                fig_time.add_trace(go.Scatter(
                    x=time_df["월"], y=time_df[col],
                    mode="lines+markers", name=col
                ))
        fig_time.update_layout(
            margin=dict(l=40, r=20, t=40, b=40),
            xaxis_title="월", yaxis_title="매출",
            height=380,
            legend=dict(orientation="h", yanchor="bottom", y=-0.25)
        )
        st.plotly_chart(fig_time, use_container_width=True)

# ---- 3) 파이차트 / 4) 산점도 / 5) 파레토 ----
c3, c4, c5 = st.columns(3)

with c3:
    st.subheader("제품별 1분기 매출 비중 (도넛 파이)")
    fig_pie = go.Figure(data=[go.Pie(
        labels=pie_df["제품"],
        values=pie_df["1분기 매출"],
        hole=0.45
    )])
    fig_pie.update_traces(textinfo="label+percent",
                          hovertemplate="%{label}: %{value:,} (%{percent})<extra></extra>")
    fig_pie.update_layout(margin=dict(l=10, r=10, t=30, b=10), height=380)
    st.plotly_chart(fig_pie, use_container_width=True)

with c4:
    st.subheader("제품 A 매출 vs 비용 (산점도)")
    fig_scatter = go.Figure(data=[go.Scatter(
        x=scatter_df["제품 A 매출"],
        y=scatter_df["비용"],
        mode="markers",
        marker=dict(size=10, line=dict(color="rgba(0,0,0,0.35)", width=0.8))
    )])
    fig_scatter.update_layout(
        xaxis_title="제품 A 매출", yaxis_title="비용",
        margin=dict(l=40, r=20, t=40, b=40),
        height=380
    )
    st.plotly_chart(fig_scatter, use_container_width=True)

with c5:
    st.subheader("부서별 매출 파레토")
    fig_pareto = make_subplots(specs=[[{"secondary_y": True}]])
    fig_pareto.add_trace(go.Bar(
        x=pareto_sorted["부서"],
        y=pareto_sorted["매출"],
        name="매출"
    ), secondary_y=False)
    fig_pareto.add_trace(go.Scatter(
        x=pareto_sorted["부서"],
        y=pareto_sorted["누적 비율"],
        name="누적 비율",
        mode="lines+markers"
    ), secondary_y=True)
    fig_pareto.update_yaxes(title_text="매출", secondary_y=False)
    fig_pareto.update_yaxes(title_text="누적 비율(%)", range=[0, 100], secondary_y=True)
    fig_pareto.add_hline(y=80, line_dash="dash", line_color="gray", secondary_y=True)
    fig_pareto.update_layout(margin=dict(l=40, r=20, t=40, b=40),
                             height=380,
                             legend=dict(orientation="h", y=-0.2))
    st.plotly_chart(fig_pareto, use_container_width=True)

# ---- 6) 버블차트 ----
st.subheader("제품별 비용 vs 마진 (버블: 크기=고객 수)")
fig_bubble = go.Figure(data=[go.Scatter(
    x=bubble_df["제품별 비용"],
    y=bubble_df["마진"],
    mode="markers",
    text=bubble_df["제품"],
    hovertemplate="%{text}<br>비용: %{x:,}<br>마진: %{y:,}<br>고객 수: %{marker.size:,}<extra></extra>",
    marker=dict(
        size=bubble_df["고객 수"],
        sizemode="area",
        sizeref=sizeref,
        sizemin=6,
        line=dict(color="rgba(0,0,0,0.35)", width=0.8)
    )
)])
fig_bubble.update_layout(
    xaxis_title="제품별 비용", yaxis_title="마진",
    margin=dict(l=40, r=20, t=20, b=40),
    height=420
)
st.plotly_chart(fig_bubble, use_container_width=True)

st.caption("© Streamlit + Plotly • 범례/드래그/더블클릭으로 인터랙션")
