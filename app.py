# -*- coding: utf-8 -*-
"""
RENKEI - 地域連携室 予約業務実績ダッシュボード

元データ: GitHub RENKEI/latest.xlsx または画面アップロードExcel
必須列: 年度, 月, 月番号, 稼働日数, 予約件数
仕様: 予約件数が空欄の未来月・未入力月は 0 件扱いせず、計算対象外にする。
"""

from __future__ import annotations

import io
import re
from typing import Iterable, Optional

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

# =========================================================
# 基本設定
# =========================================================
st.set_page_config(
    page_title="RENKEI｜地域連携室 予約業務実績",
    page_icon="🔗",
    layout="wide",
)

DEFAULT_OWNER = "faker101730-droid"
DEFAULT_REPO = "RENKEI"
DEFAULT_BRANCH = "main"
DEFAULT_FILE_PATH = "latest.xlsx"

FISCAL_MONTHS = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
MONTH_LABELS = {m: f"{m}月" for m in FISCAL_MONTHS}
MONTH_ORDER = {m: i + 1 for i, m in enumerate(FISCAL_MONTHS)}


# =========================================================
# CSS
# =========================================================
st.markdown(
    """
    <style>
    .main .block-container { padding-top: 1.3rem; padding-bottom: 2rem; }
    .renkei-title {
        font-size: 2.0rem;
        font-weight: 800;
        letter-spacing: 0.02em;
        margin-bottom: 0.15rem;
    }
    .renkei-subtitle {
        color: rgba(120,120,120,0.95);
        font-size: 0.98rem;
        margin-bottom: 1.1rem;
    }
    .note-box {
        border: 1px solid rgba(120,120,120,0.25);
        border-radius: 14px;
        padding: 0.9rem 1.0rem;
        background: rgba(120,120,120,0.06);
        margin-bottom: 1rem;
        line-height: 1.65;
    }
    div[data-testid="stMetric"] {
        border: 1px solid rgba(120,120,120,0.23);
        border-radius: 16px;
        padding: 0.85rem 0.95rem;
        background: rgba(120,120,120,0.055);
    }
    div[data-testid="stMetric"] label { font-weight: 700; }
    .small-caption { color: rgba(120,120,120,0.95); font-size: 0.88rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# ユーティリティ
# =========================================================
def safe_secret(key: str, default: str) -> str:
    """Streamlit secrets が未設定でも落ちないように取得。"""
    try:
        return str(st.secrets.get(key, default))
    except Exception:
        return default


def normalize_text(value: object) -> str:
    text = str(value).strip()
    text = text.replace(" ", "").replace("　", "")
    text = text.replace("_", "").replace("-", "")
    return text.lower()


def find_column(columns: Iterable[object], keywords: list[str], required_name: str) -> str:
    """列名の表記ゆれを吸収して、該当列を返す。"""
    normalized_columns = [(str(col), normalize_text(col)) for col in columns]
    normalized_keywords = [normalize_text(k) for k in keywords]

    # 完全一致を優先
    for original, norm_col in normalized_columns:
        if norm_col in normalized_keywords:
            return original

    # 部分一致
    for original, norm_col in normalized_columns:
        if any(key in norm_col for key in normalized_keywords):
            return original

    raise ValueError(
        f"必須列「{required_name}」が見つかりません。Excelの列名を確認してください。"
    )


def fiscal_year_sort_key(year: object) -> int:
    """R6, R7 のような年度表記を自然順にする。"""
    text = str(year)
    match = re.search(r"(\d+)", text)
    if match:
        return int(match.group(1))
    return 999


def parse_month(value: object) -> Optional[int]:
    if pd.isna(value):
        return None
    text = str(value).strip()
    match = re.search(r"(\d{1,2})", text)
    if not match:
        return None
    month = int(match.group(1))
    return month if 1 <= month <= 12 else None


def format_number(value: Optional[float], digits: int = 0, suffix: str = "") -> str:
    if value is None or pd.isna(value):
        return "-"
    if digits == 0:
        return f"{value:,.0f}{suffix}"
    return f"{value:,.{digits}f}{suffix}"


def format_delta(value: Optional[float], digits: int = 0, suffix: str = "") -> Optional[str]:
    if value is None or pd.isna(value):
        return None
    sign = "+" if value > 0 else ""
    if digits == 0:
        return f"{sign}{value:,.0f}{suffix}"
    return f"{sign}{value:,.{digits}f}{suffix}"


def get_month_range(start_month: int, end_month: int) -> list[int]:
    start_idx = FISCAL_MONTHS.index(start_month)
    end_idx = FISCAL_MONTHS.index(end_month)
    if start_idx <= end_idx:
        return FISCAL_MONTHS[start_idx : end_idx + 1]
    # 通常はUI順で起きないが、念のため年度内で一周できるようにする
    return FISCAL_MONTHS[start_idx:] + FISCAL_MONTHS[: end_idx + 1]


# =========================================================
# データ読込・整形
# =========================================================
@st.cache_data(ttl=600, show_spinner=False)
def load_from_github(owner: str, repo: str, branch: str, file_path: str) -> pd.DataFrame:
    raw_url = f"https://raw.githubusercontent.com/{owner}/{repo}/{branch}/{file_path}"
    response = requests.get(raw_url, timeout=20)
    response.raise_for_status()
    return pd.read_excel(io.BytesIO(response.content), sheet_name=0, engine="openpyxl")


def load_from_upload(uploaded_file) -> pd.DataFrame:
    return pd.read_excel(uploaded_file, sheet_name=0, engine="openpyxl")


def standardize_renkei_data(raw_df: pd.DataFrame) -> pd.DataFrame:
    if raw_df is None or raw_df.empty:
        raise ValueError("Excelにデータがありません。")

    # 空行・空列を整理
    df = raw_df.copy()
    df = df.dropna(how="all").dropna(axis=1, how="all")
    df.columns = [str(c).strip() for c in df.columns]

    year_col = find_column(df.columns, ["年度"], "年度")
    month_col = find_column(df.columns, ["月ツキ", "月"], "月")
    month_no_col = find_column(df.columns, ["月番号", "月番"], "月番号")
    working_days_col = find_column(df.columns, ["稼働日数", "稼働"], "稼働日数")
    reservation_col = find_column(df.columns, ["予約件数", "予約"], "予約件数")

    df = df.rename(
        columns={
            year_col: "年度",
            month_col: "月",
            month_no_col: "月番号",
            working_days_col: "稼働日数",
            reservation_col: "予約件数",
        }
    )

    df = df[["年度", "月", "月番号", "稼働日数", "予約件数"]].copy()
    df["年度"] = df["年度"].astype(str).str.strip().str.upper()
    df["年度"] = df["年度"].replace({"NAN": pd.NA, "NONE": pd.NA, "": pd.NA})

    month_from_no = pd.to_numeric(df["月番号"], errors="coerce")
    month_from_label = df["月"].map(parse_month)
    df["月番号"] = month_from_no.fillna(month_from_label)
    df["月番号"] = pd.to_numeric(df["月番号"], errors="coerce")

    df["稼働日数"] = pd.to_numeric(df["稼働日数"], errors="coerce")
    df["予約件数"] = pd.to_numeric(df["予約件数"], errors="coerce")

    df = df[df["年度"].notna() & df["月番号"].isin(FISCAL_MONTHS)].copy()
    df["月番号"] = df["月番号"].astype(int)
    df["月"] = df["月番号"].map(MONTH_LABELS)
    df["年度内順"] = df["月番号"].map(MONTH_ORDER)

    # 同じ年度・月が複数ある場合は合算。稼働日数は最大値を採用。
    df = (
        df.groupby(["年度", "月番号", "月", "年度内順"], as_index=False)
        .agg({"稼働日数": "max", "予約件数": "sum"})
        .sort_values(["年度", "年度内順"])
    )

    # 注意：groupby sum は全欠損を 0 にしがちなので、元データで全欠損だった行は空欄へ戻す
    blank_keys = (
        raw_df.copy()
        .rename(columns={year_col: "年度", month_col: "月", month_no_col: "月番号", reservation_col: "予約件数"})
    )
    blank_keys["年度"] = blank_keys["年度"].astype(str).str.strip().str.upper()
    blank_keys["月番号"] = pd.to_numeric(blank_keys["月番号"], errors="coerce").fillna(blank_keys["月"].map(parse_month))
    blank_keys = blank_keys[blank_keys["予約件数"].isna()][["年度", "月番号"]].drop_duplicates()
    nonblank_keys = (
        raw_df.copy()
        .rename(columns={year_col: "年度", month_col: "月", month_no_col: "月番号", reservation_col: "予約件数"})
    )
    nonblank_keys["年度"] = nonblank_keys["年度"].astype(str).str.strip().str.upper()
    nonblank_keys["月番号"] = pd.to_numeric(nonblank_keys["月番号"], errors="coerce").fillna(nonblank_keys["月"].map(parse_month))
    nonblank_keys = nonblank_keys[nonblank_keys["予約件数"].notna()][["年度", "月番号"]].drop_duplicates()
    only_blank = blank_keys.merge(nonblank_keys, on=["年度", "月番号"], how="left", indicator=True)
    only_blank = only_blank[only_blank["_merge"] == "left_only"][["年度", "月番号"]]
    if not only_blank.empty:
        blank_set = set(map(tuple, only_blank[["年度", "月番号"]].values.tolist()))
        df.loc[df.apply(lambda r: (r["年度"], r["月番号"]) in blank_set, axis=1), "予約件数"] = pd.NA

    return df


def filter_period(df: pd.DataFrame, years: list[str], months: list[int]) -> pd.DataFrame:
    return df[df["年度"].isin(years) & df["月番号"].isin(months)].copy()


def calc_metrics(df: pd.DataFrame) -> dict[str, Optional[float]]:
    actual = df[df["予約件数"].notna()].copy()
    if actual.empty:
        return {
            "latest_count": None,
            "latest_month": None,
            "total": None,
            "monthly_average": None,
            "daily_average": None,
            "actual_months": 0,
            "working_days": None,
        }

    actual = actual.sort_values("年度内順")
    total = float(actual["予約件数"].sum())
    actual_months = int(actual["月番号"].nunique())
    working_days = actual["稼働日数"].sum(skipna=True)
    working_days = float(working_days) if pd.notna(working_days) and working_days > 0 else None

    latest_row = actual.tail(1).iloc[0]
    return {
        "latest_count": float(latest_row["予約件数"]),
        "latest_month": str(latest_row["月"]),
        "total": total,
        "monthly_average": total / actual_months if actual_months > 0 else None,
        "daily_average": total / working_days if working_days else None,
        "actual_months": actual_months,
        "working_days": working_days,
    }


# =========================================================
# グラフ
# =========================================================
def apply_common_layout(fig: go.Figure, title: str, y_title: str) -> go.Figure:
    fig.update_layout(
        title={"text": title, "x": 0.02, "xanchor": "left"},
        height=390,
        margin=dict(l=40, r=20, t=60, b=40),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        hovermode="x unified",
    )
    fig.update_xaxes(title_text="月", categoryorder="array", categoryarray=[MONTH_LABELS[m] for m in FISCAL_MONTHS])
    fig.update_yaxes(title_text=y_title, rangemode="tozero")
    return fig


def build_monthly_trend_chart(chart_df: pd.DataFrame, years: list[str]) -> go.Figure:
    fig = go.Figure()
    actual = chart_df[chart_df["予約件数"].notna()].sort_values(["年度", "年度内順"])
    for year in years:
        d = actual[actual["年度"] == year]
        fig.add_trace(
            go.Scatter(
                x=d["月"],
                y=d["予約件数"],
                mode="lines+markers",
                name=year,
                connectgaps=False,
            )
        )
    return apply_common_layout(fig, "月推移｜予約件数", "予約件数")


def build_monthly_average_chart(chart_df: pd.DataFrame, years: list[str]) -> go.Figure:
    bars = []
    for year in years:
        d = chart_df[(chart_df["年度"] == year) & chart_df["予約件数"].notna()]
        avg = d["予約件数"].mean() if not d.empty else None
        bars.append({"年度": year, "月平均": avg})
    bar_df = pd.DataFrame(bars)

    fig = go.Figure()
    fig.add_trace(go.Bar(x=bar_df["年度"], y=bar_df["月平均"], text=bar_df["月平均"], texttemplate="%{text:,.1f}", textposition="outside"))
    fig.update_layout(
        title={"text": "月平均｜期間内の実績月平均", "x": 0.02, "xanchor": "left"},
        height=390,
        margin=dict(l=40, r=20, t=60, b=40),
        showlegend=False,
    )
    fig.update_yaxes(title_text="月平均予約件数", rangemode="tozero")
    return fig


def build_daily_average_chart(chart_df: pd.DataFrame, years: list[str]) -> go.Figure:
    actual = chart_df[
        chart_df["予約件数"].notna() & chart_df["稼働日数"].notna() & (chart_df["稼働日数"] > 0)
    ].copy()
    actual["1日平均"] = actual["予約件数"] / actual["稼働日数"]
    actual = actual.sort_values(["年度", "年度内順"])

    fig = go.Figure()
    for year in years:
        d = actual[actual["年度"] == year]
        fig.add_trace(go.Scatter(x=d["月"], y=d["1日平均"], mode="lines+markers", name=year, connectgaps=False))
    return apply_common_layout(fig, "1日平均｜予約件数 ÷ 稼働日数", "1日平均予約件数")


def build_cumulative_chart(chart_df: pd.DataFrame, years: list[str]) -> go.Figure:
    actual = chart_df[chart_df["予約件数"].notna()].sort_values(["年度", "年度内順"]).copy()
    actual["累計"] = actual.groupby("年度")["予約件数"].cumsum()

    fig = go.Figure()
    for year in years:
        d = actual[actual["年度"] == year]
        fig.add_trace(go.Scatter(x=d["月"], y=d["累計"], mode="lines+markers", name=year, connectgaps=False))
    return apply_common_layout(fig, "累計｜期間内予約件数", "累計予約件数")


# =========================================================
# 画面本体
# =========================================================
st.markdown('<div class="renkei-title">RENKEI</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="renkei-subtitle">地域連携室 予約業務実績ダッシュボード｜月次・累計・1日平均を確認</div>',
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("データ読込")
    uploaded_file = st.file_uploader("Excelをアップロード（任意）", type=["xlsx"])

    owner = safe_secret("GITHUB_OWNER", DEFAULT_OWNER)
    repo = safe_secret("GITHUB_REPO", DEFAULT_REPO)
    branch = safe_secret("GITHUB_BRANCH", DEFAULT_BRANCH)
    file_path = safe_secret("GITHUB_FILE_PATH", DEFAULT_FILE_PATH)

    st.caption("アップロードがない場合はGitHubの latest.xlsx を読み込みます。")
    st.code(f"{owner}/{repo}/{file_path}", language="text")

try:
    if uploaded_file is not None:
        raw = load_from_upload(uploaded_file)
        data_source = "アップロードExcel"
    else:
        raw = load_from_github(owner, repo, branch, file_path)
        data_source = f"GitHub: {owner}/{repo}/{file_path}"
    df = standardize_renkei_data(raw)
except Exception as e:
    st.error("データを読み込めませんでした。Excel様式またはGitHub保存先を確認してください。")
    st.exception(e)
    st.stop()

all_years = sorted(df["年度"].dropna().unique().tolist(), key=fiscal_year_sort_key)
years_with_actual = sorted(df[df["予約件数"].notna()]["年度"].dropna().unique().tolist(), key=fiscal_year_sort_key)

if not all_years:
    st.warning("年度データがありません。テンプレートに年度・月番号・稼働日数・予約件数を入力してください。")
    st.stop()

# デフォルト年度
if years_with_actual:
    default_target = years_with_actual[-1]
else:
    default_target = all_years[-1]

default_comp = all_years[max(0, all_years.index(default_target) - 1)] if default_target in all_years else all_years[0]

with st.sidebar:
    st.header("分析条件")
    target_year = st.selectbox("対象年度", all_years, index=all_years.index(default_target))

    comparison_candidates = [y for y in all_years if y != target_year]
    if not comparison_candidates:
        comparison_candidates = [target_year]
    comparison_index = comparison_candidates.index(default_comp) if default_comp in comparison_candidates else 0
    comparison_year = st.selectbox("比較年度", comparison_candidates, index=comparison_index)

    month_options = [MONTH_LABELS[m] for m in FISCAL_MONTHS]
    start_label = st.selectbox("開始月", month_options, index=0)

    # 実績がある最後の月を終了月の初期値にする
    target_actual = df[(df["年度"] == target_year) & df["予約件数"].notna()].sort_values("年度内順")
    if not target_actual.empty:
        default_end_month = int(target_actual.tail(1).iloc[0]["月番号"])
    else:
        default_end_month = 3
    end_label = st.selectbox("終了月", month_options, index=FISCAL_MONTHS.index(default_end_month))

start_month = int(start_label.replace("月", ""))
end_month = int(end_label.replace("月", ""))
selected_months = get_month_range(start_month, end_month)
selected_years = [target_year, comparison_year]
period_df = filter_period(df, selected_years, selected_months)

st.markdown(
    f"""
    <div class="note-box">
    <b>読込元：</b>{data_source}<br>
    <b>分析条件：</b>{target_year} vs {comparison_year} ／ {start_label}〜{end_label}<br>
    <span class="small-caption">予約件数が空欄の月は、未来月・未入力月として0件扱いせず計算対象外にしています。</span>
    </div>
    """,
    unsafe_allow_html=True,
)

target_df = period_df[period_df["年度"] == target_year]
comp_df = period_df[period_df["年度"] == comparison_year]
target_metrics = calc_metrics(target_df)
comp_metrics = calc_metrics(comp_df)

if target_metrics["actual_months"] == 0:
    st.warning(f"{target_year} の {start_label}〜{end_label} に予約件数の実績がありません。Excelに予約件数を入力してください。")

# KPI cards
col1, col2, col3, col4 = st.columns(4)
with col1:
    latest_label = "最新月件数"
    if target_metrics["latest_month"]:
        latest_label += f"（{target_metrics['latest_month']}）"
    latest_delta = None
    if comp_metrics["latest_count"] is not None and target_metrics["latest_count"] is not None:
        latest_delta = target_metrics["latest_count"] - comp_metrics["latest_count"]
    st.metric(latest_label, format_number(target_metrics["latest_count"], 0, "件"), format_delta(latest_delta, 0, "件"))

with col2:
    total_delta = None
    if comp_metrics["total"] is not None and target_metrics["total"] is not None:
        total_delta = target_metrics["total"] - comp_metrics["total"]
    st.metric("期間累計", format_number(target_metrics["total"], 0, "件"), format_delta(total_delta, 0, "件"))

with col3:
    avg_delta = None
    if comp_metrics["monthly_average"] is not None and target_metrics["monthly_average"] is not None:
        avg_delta = target_metrics["monthly_average"] - comp_metrics["monthly_average"]
    st.metric("月平均", format_number(target_metrics["monthly_average"], 1, "件"), format_delta(avg_delta, 1, "件"))

with col4:
    day_delta = None
    if comp_metrics["daily_average"] is not None and target_metrics["daily_average"] is not None:
        day_delta = target_metrics["daily_average"] - comp_metrics["daily_average"]
    st.metric("1日平均", format_number(target_metrics["daily_average"], 1, "件"), format_delta(day_delta, 1, "件"))

st.divider()

left, right = st.columns(2)
with left:
    st.plotly_chart(build_monthly_trend_chart(period_df, selected_years), use_container_width=True)
with right:
    st.plotly_chart(build_monthly_average_chart(period_df, selected_years), use_container_width=True)

left, right = st.columns(2)
with left:
    st.plotly_chart(build_daily_average_chart(period_df, selected_years), use_container_width=True)
with right:
    st.plotly_chart(build_cumulative_chart(period_df, selected_years), use_container_width=True)

with st.expander("集計データを確認", expanded=False):
    display_df = period_df.sort_values(["年度", "年度内順"])[["年度", "月", "月番号", "稼働日数", "予約件数"]]
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    csv = display_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "集計データCSVをダウンロード",
        data=csv,
        file_name="renkei_filtered_data.csv",
        mime="text/csv",
    )

st.caption("RENKEI 初期復活版：元データは月次集計済みExcel。詳細な紹介元・診療科別分析は、将来的にLINK/STRIKE側と連携する想定。")
