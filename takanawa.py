import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(
    page_title="居住地ダッシュボード（前年比対応版）",
    layout="wide"
)

st.title("居住地ダッシュボード - 前年比対応版")
st.caption("サンバースト / パレート / ランキング / 前年比")

# =========================
# 1) ファイル読み込み
# =========================

@st.cache_data
def load_default_data():
    return pd.read_excel("takanawapt.xlsx")


uploaded = st.file_uploader(
    "別のExcelで分析する場合はこちら",
    type=["xlsx", "csv"]
)

if uploaded is not None:
    if uploaded.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded)
else:
    df = load_default_data()

df.columns = [str(c).strip() for c in df.columns]

# =========================
# 2) 年月の標準化
# =========================
def to_month_start(x):
    try:
        return pd.to_datetime(x).to_period("M").to_timestamp()
    except Exception:
        return pd.NaT

if "年月" in df.columns:
    df["年月"] = df["年月"].apply(to_month_start)

elif {"年", "月"}.issubset(df.columns):

    df["年"] = pd.to_numeric(df["年"], errors="coerce").astype("Int64")
    df["月"] = pd.to_numeric(df["月"], errors="coerce").astype("Int64")

    df = df.dropna(subset=["年", "月"])

    if len(df) == 0:
        st.error("『年』『月』の有効な値がありません。")
        st.stop()

    df["年月"] = pd.to_datetime(
        dict(
            year=df["年"].astype(int),
            month=df["月"].astype(int),
            day=1
        )
    )

else:
    st.error("『年月』列 または 『年』『月』列が必要です。")
    st.stop()

if df["年月"].isna().all():
    st.error("『年月』が日付として解釈できません。")
    st.stop()

# =========================
# 3) 居住地 → 都道府県 / 市区町村
# =========================
if not {"都道府県", "市区町村"}.issubset(df.columns):

    if "居住地" not in df.columns:
        st.error("『居住地』 または 『都道府県』『市区町村』列が必要です。")
        st.stop()

    pref_pattern = re.compile(r"^(北海道|.+?県|.+?府|.+?都)")

    def extract_pref(x):
        if not isinstance(x, str):
            return np.nan

        m = pref_pattern.match(x.strip())

        return m.group(0) if m else np.nan

    def extract_city(x):

        if not isinstance(x, str):
            return np.nan

        p = extract_pref(x)

        if not isinstance(p, str):
            return np.nan

        rest = x.strip()[len(p):].lstrip()

        if rest == "":
            return np.nan

        m = re.match(r"(.+?(市|区|郡|町|村))", rest)

        return m.group(1) if m else rest

    df["都道府県"] = df["居住地"].apply(extract_pref)
    df["市区町村"] = df["居住地"].apply(extract_city)

df = df.dropna(subset=["都道府県", "市区町村"]).copy()

df["都道府県"] = df["都道府県"].astype(str)
df["市区町村"] = df["市区町村"].astype(str)

# =========================
# 4) 性別処理
# =========================
def normalize_gender(x):

    if pd.isna(x):
        return None

    s = str(x).strip()

    if s in ["男", "男性", "M", "m", "male"]:
        return "男"

    if s in ["女", "女性", "F", "f", "female"]:
        return "女"

    return None

def detect_value_col(columns):

    for c in ["人数", "件数", "数", "count", "n", "value", "合計"]:
        if c in columns:
            return c

    return None

has_wide = ("男" in df.columns) and ("女" in df.columns)
has_long = ("性別" in df.columns)

if has_wide:

    df["男"] = pd.to_numeric(df["男"], errors="coerce").fillna(0)
    df["女"] = pd.to_numeric(df["女"], errors="coerce").fillna(0)

else:

    if has_long:

        val_col = detect_value_col(df.columns)

        df["_性別norm"] = df["性別"].apply(normalize_gender)

        if val_col is None:
            df["_row_count"] = 1
            val_col = "_row_count"

        else:
            df[val_col] = pd.to_numeric(df[val_col], errors="coerce").fillna(0)

        idx_cols = ["年月", "都道府県", "市区町村"]

        wide = (
            df.dropna(subset=["_性別norm"])
            .pivot_table(
                index=idx_cols,
                columns="_性別norm",
                values=val_col,
                aggfunc="sum",
                fill_value=0
            )
            .reset_index()
        )

        if "男" not in wide.columns:
            wide["男"] = 0

        if "女" not in wide.columns:
            wide["女"] = 0

        df = wide.copy()

    else:
        df["_row_count"] = 1
        df["男"] = 0
        df["女"] = 0

# =========================
# 合計列
# =========================
if "合計" in df.columns:
    df["合計"] = pd.to_numeric(df["合計"], errors="coerce")

else:
    df["合計"] = np.nan

mask_nan = df["合計"].isna()

df.loc[mask_nan, "合計"] = (
    df.get("男", 0) + df.get("女", 0)
).where(mask_nan)

mask_nan = df["合計"].isna()

if "_row_count" in df.columns:
    df.loc[mask_nan, "合計"] = df.loc[mask_nan, "_row_count"]

df["合計"] = pd.to_numeric(df["合計"], errors="coerce").fillna(0)

# =========================
# 5) サイドバー
# =========================
with st.sidebar:

    st.header("設定")

    metric = st.radio(
        "指標",
        ["合計", "男", "女"],
        horizontal=True
    )

    mode = st.radio(
        "集計期間",
        ["単月", "期間合計", "期間指定"],
        index=0,
        horizontal=True
    )

    ym_opts = sorted(df["年月"].dropna().unique())

    ym_pick = None
    start_ym = None
    end_ym = None

    if mode == "単月":

        ym_pick = st.selectbox(
            "対象年月",
            options=ym_opts,
            index=max(len(ym_opts)-1, 0),
            format_func=lambda d: pd.to_datetime(d).strftime("%Y-%m")
        )

    elif mode == "期間指定":

        start_ym, end_ym = st.select_slider(
            "対象期間",
            options=ym_opts,
            value=(ym_opts[0], ym_opts[-1]),
            format_func=lambda d: pd.to_datetime(d).strftime("%Y-%m")
        )

    pareto_unit = st.radio(
        "粒度",
        ["都道府県", "市区町村"],
        index=0,
        horizontal=True
    )

    top_k = st.number_input(
        "ランキング上位K",
        min_value=5,
        max_value=200,
        value=20,
        step=1
    )

    show_yoy = st.checkbox(
        "前年比を表示",
        value=True
    )

# =========================
# 6) 対象期間抽出
# =========================
if mode == "単月":

    base = df[df["年月"] == ym_pick].copy()

elif mode == "期間指定":

    base = df[
        (df["年月"] >= start_ym) &
        (df["年月"] <= end_ym)
    ].copy()

else:
    base = df.copy()

if len(base) == 0:
    st.warning("該当データがありません。")
    st.stop()

value_col = metric

base[value_col] = pd.to_numeric(
    base[value_col],
    errors="coerce"
).fillna(0)

# =========================
# 7) 前年比
# =========================
prev_base = pd.DataFrame()

if show_yoy:

    if mode == "単月":

        prev_ym = pd.to_datetime(ym_pick) - pd.DateOffset(years=1)

        prev_base = df[
            df["年月"] == prev_ym
        ].copy()

    elif mode == "期間指定":

        prev_start = pd.to_datetime(start_ym) - pd.DateOffset(years=1)
        prev_end = pd.to_datetime(end_ym) - pd.DateOffset(years=1)

        prev_base = df[
            (df["年月"] >= prev_start) &
            (df["年月"] <= prev_end)
        ].copy()

# =========================
# 8) KPI表示
# =========================
if show_yoy and len(prev_base) > 0:

    current_total = base[value_col].sum()
    prev_total = prev_base[value_col].sum()

    yoy = 0

    if prev_total != 0:
        yoy = (current_total / prev_total - 1) * 100

    diff = current_total - prev_total

    col1, col2, col3, col4 = st.columns(4)

    col1.metric(
        "現在",
        f"{current_total:,.0f}"
    )

    col2.metric(
        "前年",
        f"{prev_total:,.0f}"
    )

    col3.metric(
        "前年差",
        f"{diff:,.0f}"
    )

    col4.metric(
        "前年比",
        f"{yoy:.1f}%"
    )

# =========================
# KPI追加
# =========================

total_patients = int(base["合計"].sum())

pref_count = base["都道府県"].nunique()

outside_patients = int(
    base.loc[
        base["都道府県"] != "宮城県",
        "合計"
    ].sum()
)

outside_ratio = 0

if total_patients > 0:
    outside_ratio = (
        outside_patients /
        total_patients * 100
    )

city_count = base["市区町村"].nunique()

k1, k2, k3, k4 = st.columns(4)

k1.metric(
    "総患者数",
    f"{total_patients:,}"
)

k2.metric(
    "県外患者数",
    f"{outside_patients:,}"
)

k3.metric(
    "県外比率",
    f"{outside_ratio:.1f}%"
)

k4.metric(
    "居住市区町村数",
    f"{city_count:,}"
)

# =========================
# 9) サンバースト
# =========================
st.subheader(f"サンバースト（{value_col}）")

sun = (
    base.groupby(
        ["都道府県", "市区町村"],
        as_index=False
    )[value_col]
    .sum()
)

if float(sun[value_col].sum()) <= 0:

    st.info("表示データがありません。")

else:

    fig_sb = px.sunburst(
        sun,
        path=["都道府県", "市区町村"],
        values=value_col,
        color="都道府県",
        height=550
    )

    st.plotly_chart(
        fig_sb,
        use_container_width=True
    )

# =========================
# 10) パレート図
# =========================
st.subheader(f"パレート図（{pareto_unit}）")

grp = (
    base.groupby(
        pareto_unit,
        as_index=False
    )[value_col]
    .sum()
    .sort_values(value_col, ascending=False)
)

total = float(grp[value_col].sum())

if total <= 0:

    st.info("表示データがありません。")

else:

    grp["累積"] = grp[value_col].cumsum()
    grp["累積比率"] = grp["累積"] / total * 100

    fig_pt = make_subplots(
        specs=[[{"secondary_y": True}]]
    )

    fig_pt.add_trace(
        go.Bar(
            x=grp[pareto_unit],
            y=grp[value_col],
            name=value_col
        ),
        secondary_y=False
    )

    fig_pt.add_trace(
        go.Scatter(
            x=grp[pareto_unit],
            y=grp["累積比率"],
            mode="lines+markers",
            name="累積比率"
        ),
        secondary_y=True
    )

    fig_pt.update_yaxes(
        title_text=value_col,
        secondary_y=False
    )

    fig_pt.update_yaxes(
        title_text="累積比率(%)",
        range=[0, 105],
        secondary_y=True
    )

    fig_pt.update_layout(
        height=550,
        hovermode="x unified"
    )

    st.plotly_chart(
        fig_pt,
        use_container_width=True
    )

# =========================
# 11) ランキング
# =========================
st.subheader(f"ランキング（{pareto_unit}）")

rank = grp[
    [pareto_unit, value_col]
].copy().head(int(top_k))

if rank.empty:

    st.info("データがありません。")

else:

    fig_bar = go.Figure(
        data=[
            go.Bar(
                x=rank[value_col].iloc[::-1],
                y=rank[pareto_unit].iloc[::-1],
                orientation="h"
            )
        ]
    )

    fig_bar.update_layout(
        height=max(350, 24 * len(rank)),
        xaxis_title=value_col,
        yaxis_title=pareto_unit
    )

    st.plotly_chart(
        fig_bar,
        use_container_width=True
    )

    st.dataframe(
        rank.reset_index(drop=True),
        use_container_width=True
    )

    st.download_button(
        label="CSVダウンロード",
        data=rank.to_csv(index=False).encode("utf-8-sig"),
        file_name="ranking.csv",
        mime="text/csv"
    )

# =========================
# 都道府県別 月間件数
# =========================
st.subheader("都道府県別 月間件数")

pref_order = (
    df.groupby("都道府県")["合計"]
      .sum()
      .sort_values(ascending=False)
      .index
)

pref_monthly = pd.pivot_table(
    df,
    index="年月",
    columns="都道府県",
    values="合計",
    aggfunc="sum",
    fill_value=0
)

pref_monthly = pref_monthly.reindex(
    columns=pref_order,
    fill_value=0
)

pref_monthly.loc["合計"] = pref_monthly.sum()

pref_monthly.index = [
    f"{d.year}年{d.month}月" if isinstance(d, pd.Timestamp) else d
    for d in pref_monthly.index
]

pref_monthly.index.name = "年月"

st.dataframe(pref_monthly, use_container_width=True)


# =========================
# 市区町村別 月間件数
# =========================
st.subheader("市区町村別 月間件数")

city_order = (
    df.groupby("市区町村")["合計"]
      .sum()
      .sort_values(ascending=False)
      .index
)

city_monthly = pd.pivot_table(
    df,
    index="年月",
    columns="市区町村",
    values="合計",
    aggfunc="sum",
    fill_value=0
)

city_monthly = city_monthly.reindex(
    columns=city_order,
    fill_value=0
)

city_monthly.loc["合計"] = city_monthly.sum()

city_monthly.index = [
    f"{d.year}年{d.month}月" if isinstance(d, pd.Timestamp) else d
    for d in city_monthly.index
]

city_monthly.index.name = "年月"

st.dataframe(city_monthly, use_container_width=True)

pref_monthly.index.name = "年月"
city_monthly.index.name = "年月"

# =========================
# ダウンロード
# =========================

from io import BytesIO

output = BytesIO()

with pd.ExcelWriter(output, engine="openpyxl") as writer:

    pref_monthly.to_excel(
        writer,
        sheet_name="都道府県別月間件数"
    )

    city_monthly.to_excel(
        writer,
        sheet_name="市区町村別月間件数"
    )

st.download_button(
    "月間集計Excelダウンロード",
    data=output.getvalue(),
    file_name="月間集計.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)