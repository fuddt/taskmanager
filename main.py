import os
import tomllib
from pathlib import Path

import pandas as pd
import streamlit as st

# ============================================================
# Streamlit 進捗管理ビュー
# - SSOT: Excel の「進捗一覧」シート
# - 目的: 会議で見やすい進捗サマリ + ガントチャート表示
# - 方針:
#   1. Altair を本命にする
#   2. Altair が使えない場合は text 疑似ガントへフォールバック
# ============================================================

st.set_page_config(page_title="進捗ガントビュー", layout="wide")
st.title("進捗ガントビュー")
st.caption("Excel の『進捗一覧』を正本として読み込み、会議向けに可視化します。")

REQUIRED_COLUMNS = [
    "ID",
    "機能",
    "タスク区分",
    "タスク名",
    "担当",
    "開始日",
    "期限",
    "進捗率",
    "ステータス",
    "優先度",
]

STATUS_ORDER = ["未着手", "進行中", "完了", "保留"]
STATUS_COLOR_MAP = {
    "未着手": "#BDBDBD",
    "進行中": "#1F77B4",
    "完了": "#2CA02C",
    "保留": "#7F7F7F",
}

# Altair は Streamlit 環境に同梱されていることが多いが、
# 念のため import 可否を見ておく。
try:
    import altair as alt
    ALTAIR_AVAILABLE = True
except ImportError:
    ALTAIR_AVAILABLE = False


def normalize_progress(value) -> float:
    """
    進捗率を 0-100 の数値に正規化する。
    Excel 側で 40 / 40% / 0.4 のようにブレても吸収するための関数。
    """
    if pd.isna(value):
        return 0.0

    if isinstance(value, str):
        cleaned = value.replace("%", "").strip()
        if cleaned == "":
            return 0.0
        try:
            num = float(cleaned)
        except ValueError:
            return 0.0
    else:
        try:
            num = float(value)
        except (TypeError, ValueError):
            return 0.0

    if 0 <= num <= 1:
        num *= 100

    return max(0.0, min(100.0, num))


def detect_header_row(file_path: str, sheet_name: str, preview_rows: int = 15) -> int:
    """
    シート先頭をヘッダなしで仮読みし、
    REQUIRED_COLUMNS が最も多く並んでいる行をヘッダ行として判定する。
    先頭にタイトルや説明行があっても壊れないようにする。
    """
    preview_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=preview_rows)

    best_row_index = 0
    best_match_count = -1

    for row_idx in range(len(preview_df)):
        row_values = {str(v).strip() for v in preview_df.iloc[row_idx].tolist() if pd.notna(v)}
        match_count = sum(1 for col in REQUIRED_COLUMNS if col in row_values)
        if match_count > best_match_count:
            best_match_count = match_count
            best_row_index = row_idx

    return best_row_index


@st.cache_data
def load_progress_excel(file_path: str, sheet_name: str, mtime: float) -> pd.DataFrame:
    """
    Excel ファイルを読み込み、進捗一覧 DataFrame を返す。
    タイトル行や説明行が先頭にあってもヘッダ行を自動検出する。
    mtime はキャッシュ無効化用（ファイル更新時に自動バスト）。
    """
    header_row = detect_header_row(file_path, sheet_name)
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
    return df



def validate_columns(df: pd.DataFrame) -> list[str]:
    """必須列の不足を確認する。"""
    return [col for col in REQUIRED_COLUMNS if col not in df.columns]



def prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    表示用に DataFrame を整形する。
    - 文字列正規化
    - 日付型変換
    - 進捗率正規化
    - 遅延判定再計算
    - ガント表示用列作成
    """
    work = df.copy()

    # 文字列系の列を正規化
    string_cols = [
        "ID",
        "機能",
        "タスク区分",
        "タスク名",
        "担当",
        "ステータス",
        "優先度",
        "課題/懸念",
        "次アクション",
    ]
    for col in string_cols:
        if col in work.columns:
            work[col] = work[col].astype("string").str.strip()

    # 日付型に変換
    work["開始日"] = pd.to_datetime(work["開始日"], errors="coerce")
    work["期限"] = pd.to_datetime(work["期限"], errors="coerce")
    if "更新日" in work.columns:
        work["更新日"] = pd.to_datetime(work["更新日"], errors="coerce")

    # 進捗率を正規化
    work["進捗率"] = work["進捗率"].apply(normalize_progress)

    # ステータスを文字列として正規化し、空や nan は未着手扱いに寄せる
    work["ステータス"] = work["ステータス"].astype("string").str.strip()
    work["ステータス"] = work["ステータス"].replace({
        pd.NA: "未着手",
        "": "未着手",
        "nan": "未着手",
        "None": "未着手",
        "<NA>": "未着手",
    })
    work["ステータス"] = work["ステータス"].fillna("未着手")

    # 優先度も空を吸収
    work["優先度"] = work["優先度"].replace({pd.NA: "", "nan": "", "None": "", "<NA>": ""}).fillna("")

    today = pd.Timestamp.today().normalize()
    work["遅延_再計算"] = work.apply(
        lambda row: "遅延"
        if pd.notna(row["期限"]) and row["期限"] < today and row["ステータス"] != "完了"
        else "",
        axis=1,
    )

    # ガント用開始日: 開始日が無ければ期限日、それも無ければ今日
    work["表示開始日"] = work["開始日"]
    work.loc[work["表示開始日"].isna(), "表示開始日"] = work.loc[
        work["表示開始日"].isna(), "期限"
    ]
    work["表示開始日"] = work["表示開始日"].fillna(today)

    # ガント用期限: 期限が無ければ開始日と同日
    work["表示期限"] = work["期限"].fillna(work["表示開始日"])

    # 開始 > 期限 の異常値を吸収
    mask = work["表示開始日"] > work["表示期限"]
    work.loc[mask, "表示期限"] = work.loc[mask, "表示開始日"]

    # 表示用ラベル
    work["表示名"] = work.apply(
        lambda row: f"[{row['機能']}] {row['タスク名']}（{row['担当']}）",
        axis=1,
    )

    # ID 順ソート（数値IDは数値比較、英数字IDは文字列比較にフォールバック）
    _id_num = pd.to_numeric(work["ID"], errors="coerce")
    work["_id_sort"] = _id_num
    work = work.sort_values(by=["_id_sort", "ID"], ascending=[True, True], na_position="last")
    work = work.drop(columns=["_id_sort"])

    return work



def build_altair_gantt(df: pd.DataFrame):
    """
    Altair でガントチャートを作る。
    ブラウザ描画なので、日本語表示の相性が matplotlib より良い。
    """
    chart_df = df.copy()
    chart_df["表示名_遅延反映"] = chart_df.apply(
        lambda row: f"[遅延] {row['表示名']}" if row["遅延_再計算"] == "遅延" else row["表示名"],
        axis=1,
    )

    # 上から下へ見たいので、今の並び順をそのまま使う
    y_order = chart_df["表示名_遅延反映"].tolist()

    # 横軸を日単位に強制する（90日以内→1日刻み、〜180日→3日、それ以上→7日）
    _date_min = chart_df["表示開始日"].min()
    _date_max = chart_df["表示期限"].max()
    _day_range = max(1, (_date_max - _date_min).days)
    if _day_range <= 90:
        _tick_step = 1
    elif _day_range <= 180:
        _tick_step = 3
    else:
        _tick_step = 7

    bars = (
        alt.Chart(chart_df)
        .mark_bar(size=24)
        .encode(
            x=alt.X(
                "表示開始日:T",
                title="日付",
                axis=alt.Axis(
                    format="%m/%d",
                    tickCount={"interval": "day", "step": _tick_step},
                    labelAngle=-45,
                    labelFontSize=10,
                ),
            ),
            x2="表示期限:T",
            y=alt.Y(
                "表示名_遅延反映:N",
                title="タスク",
                sort=y_order,
                axis=alt.Axis(labelLimit=400),
            ),
            color=alt.Color(
                "ステータス:N",
                title="ステータス",
                scale=alt.Scale(
                    domain=["未着手", "進行中", "完了", "保留"],
                    range=["#BDBDBD", "#1F77B4", "#2CA02C", "#7F7F7F"],
                ),
            ),
            tooltip=[
                alt.Tooltip("ID:N", title="ID"),
                alt.Tooltip("機能:N", title="機能"),
                alt.Tooltip("タスク区分:N", title="タスク区分"),
                alt.Tooltip("タスク名:N", title="タスク名"),
                alt.Tooltip("担当:N", title="担当"),
                alt.Tooltip("表示開始日:T", title="開始日"),
                alt.Tooltip("表示期限:T", title="期限"),
                alt.Tooltip("進捗率:Q", title="進捗率", format=".1f"),
                alt.Tooltip("優先度:N", title="優先度"),
                alt.Tooltip("ステータス:N", title="ステータス"),
                alt.Tooltip("遅延_再計算:N", title="遅延"),
            ],
        )
    )

    # 日単位の縦グリッド線（mark_rule で明示的に描画 ─ 軸ティックと連動しないので幅に左右されない）
    _grid_dates = pd.date_range(start=_date_min, end=_date_max, freq="D")
    _grid_df = pd.DataFrame({"date": _grid_dates})
    day_grid = (
        alt.Chart(_grid_df)
        .mark_rule(color="#d0d0d0", strokeWidth=0.5, opacity=0.6)
        .encode(x="date:T")
    )

    # 各タスク行の横区切り線
    row_rules = (
        alt.Chart(chart_df)
        .mark_rule(color="#cccccc", strokeWidth=0.5, opacity=0.7)
        .encode(y=alt.Y("表示名_遅延反映:N", sort=y_order))
    )

    # 今日の縦線
    today_df = pd.DataFrame({"today": [pd.Timestamp.today().normalize()]})
    today_rule = alt.Chart(today_df).mark_rule(strokeDash=[6, 4], size=2).encode(x="today:T")

    chart = (day_grid + row_rules + bars + today_rule).properties(
        height=max(400, len(chart_df) * 34),
        title="進捗ガントチャート",
    )

    return chart



def build_text_gantt_table(df: pd.DataFrame) -> pd.DataFrame:
    """
    最終フォールバック用の疑似ガント。
    依存をほぼ増やさず、最低限の視認性を確保する。
    """
    chart_df = df.copy()

    if chart_df.empty:
        return chart_df

    base_date = chart_df["表示開始日"].min()

    def make_bar(start_date, end_date, status, delayed):
        if pd.isna(start_date) or pd.isna(end_date):
            return ""

        offset = max(0, (start_date - base_date).days)
        length = max(1, (end_date - start_date).days + 1)

        if status == "完了":
            block = "■"
        elif status == "進行中":
            block = "="
        elif status == "保留":
            block = "-"
        else:
            block = "□"

        bar = " " * offset + block * length
        if delayed == "遅延":
            bar = "!" + bar
        return bar

    chart_df["疑似ガント"] = chart_df.apply(
        lambda row: make_bar(row["表示開始日"], row["表示期限"], row["ステータス"], row["遅延_再計算"]),
        axis=1,
    )

    return chart_df[
        [
            "機能",
            "タスク名",
            "担当",
            "表示開始日",
            "表示期限",
            "進捗率",
            "ステータス",
            "遅延_再計算",
            "疑似ガント",
        ]
    ].rename(
        columns={
            "表示開始日": "開始日",
            "表示期限": "期限",
            "遅延_再計算": "遅延",
        }
    )


# ------------------------------------------------------------
# 設定ファイル読み込み (config.toml)
# ------------------------------------------------------------
_CONFIG_PATH = Path(__file__).parent / "config.toml"

if not _CONFIG_PATH.exists():
    st.error(
        f"設定ファイルが見つかりません: `{_CONFIG_PATH}`\n\n"
        "以下の形式で `config.toml` を作成してください:\n\n"
        "```toml\n"
        "[excel]\n"
        'file_path = "C:/path/to/progress.xlsx"\n'
        'sheet_name = "進捗一覧"\n'
        "```"
    )
    st.stop()

try:
    with open(_CONFIG_PATH, "rb") as _f:
        _config = tomllib.load(_f)
except tomllib.TOMLDecodeError as e:
    st.error(
        f"config.toml の解析に失敗しました: {e}\n\n"
        "**Windows パスのバックスラッシュ (`\\`) は TOML のエスケープ文字として扱われます。**\n\n"
        "パスの書き方:\n"
        "- スラッシュを使う（推奨）: `\"C:/Users/you/file.xlsx\"`\n"
        "- シングルクォートで囲む: `'C:\\\\Users\\\\you\\\\file.xlsx'`"
    )
    st.stop()

_excel_cfg = _config.get("excel", {})
file_path: str = _excel_cfg.get("file_path", "")
sheet_name: str = _excel_cfg.get("sheet_name", "進捗一覧")

if not file_path or not os.path.exists(file_path):
    st.error(
        f"config.toml に指定された Excel ファイルが見つかりません: `{file_path}`\n\n"
        "パスを確認して `config.toml` を修正してください。"
    )
    st.stop()

# ------------------------------------------------------------
# サイドバー: 設定情報の表示（読み取り専用）
# ------------------------------------------------------------
st.sidebar.header("入力")
st.sidebar.text("Excel ファイル (config.toml より)")
st.sidebar.code(file_path, language=None)
st.sidebar.text(f"シート名: {sheet_name}")

try:
    _mtime = os.path.getmtime(file_path)
    detected_header_row = detect_header_row(file_path, sheet_name)
    raw_df = load_progress_excel(file_path, sheet_name, _mtime)
except Exception as e:
    st.error(f"Excel の読み込みに失敗しました: {e}")
    st.stop()

st.caption(f"検出したヘッダ行: {detected_header_row + 1} 行目")

missing_columns = validate_columns(raw_df)
if missing_columns:
    st.error("必須列が不足しています。")
    st.write("不足列:", missing_columns)
    st.stop()

progress_df = prepare_dataframe(raw_df)

# ------------------------------------------------------------
# サイドバー: フィルタ
# ------------------------------------------------------------
st.sidebar.header("フィルタ")

function_options = sorted([x for x in progress_df["機能"].dropna().astype(str).tolist() if x and x != "nan" and x != "<NA>"])
function_options = list(dict.fromkeys(function_options))
selected_functions = st.sidebar.multiselect("機能", function_options, default=function_options)

member_options = sorted([x for x in progress_df["担当"].dropna().astype(str).tolist() if x and x != "nan" and x != "<NA>"])
member_options = list(dict.fromkeys(member_options))
selected_members = st.sidebar.multiselect("担当", member_options, default=member_options)

status_values = [str(s).strip() for s in progress_df["ステータス"].dropna().tolist()]
status_values = [s for s in status_values if s and s.lower() != "nan" and s != "<NA>"]
unique_status_values = list(dict.fromkeys(status_values))
status_options = [s for s in STATUS_ORDER if s in unique_status_values] + [
    s for s in sorted(unique_status_values) if s not in STATUS_ORDER
]
selected_statuses = st.sidebar.multiselect("ステータス", status_options, default=status_options)

show_delayed_only = st.sidebar.checkbox("遅延タスクのみ表示", value=False)
exclude_done = st.sidebar.checkbox("完了を非表示", value=False)

filtered_df = progress_df.copy()
filtered_df = filtered_df[filtered_df["機能"].isin(selected_functions)]
filtered_df = filtered_df[filtered_df["担当"].isin(selected_members)]
filtered_df = filtered_df[filtered_df["ステータス"].isin(selected_statuses)]

if show_delayed_only:
    filtered_df = filtered_df[filtered_df["遅延_再計算"] == "遅延"]

if exclude_done:
    filtered_df = filtered_df[filtered_df["ステータス"] != "完了"]

# ------------------------------------------------------------
# サマリ表示
# ------------------------------------------------------------
col1, col2, col3, col4, col5 = st.columns(5)

all_count = len(filtered_df)
completed_count = int((filtered_df["ステータス"] == "完了").sum())
in_progress_count = int((filtered_df["ステータス"] == "進行中").sum())
delayed_count = int((filtered_df["遅延_再計算"] == "遅延").sum())
avg_progress = round(filtered_df["進捗率"].mean(), 1) if all_count > 0 else 0.0

col1.metric("表示タスク数", all_count)
col2.metric("平均進捗率", f"{avg_progress}%")
col3.metric("進行中", in_progress_count)
col4.metric("完了", completed_count)
col5.metric("遅延", delayed_count)

with st.expander("担当別サマリ", expanded=False):
    member_summary = (
        filtered_df.groupby("担当", dropna=False)
        .agg(
            タスク数=("ID", "count"),
            平均進捗率=("進捗率", "mean"),
            遅延件数=("遅延_再計算", lambda s: (s == "遅延").sum()),
        )
        .reset_index()
    )
    member_summary["平均進捗率"] = member_summary["平均進捗率"].round(1)
    st.dataframe(member_summary, width="stretch")

# ------------------------------------------------------------
# ガントチャート
# ------------------------------------------------------------
st.subheader("ガントチャート")

view_mode_options = ["自動判定", "altair", "text"]
selected_view_mode = st.radio("表示方式", view_mode_options, horizontal=True)

if filtered_df.empty:
    st.warning("条件に一致するタスクがありません。")
else:
    if selected_view_mode == "自動判定":
        actual_mode = "altair" if ALTAIR_AVAILABLE else "text"
    else:
        actual_mode = selected_view_mode

    if actual_mode == "altair":
        if ALTAIR_AVAILABLE:
            st.caption("表示モード: Altair")
            gantt_chart = build_altair_gantt(filtered_df)
            st.altair_chart(gantt_chart, width="stretch")
        else:
            st.warning("Altair が利用できないため、text 表示に切り替えます。")
            text_gantt_df = build_text_gantt_table(filtered_df)
            st.dataframe(text_gantt_df, width="stretch")
    else:
        st.caption("表示モード: text")
        text_gantt_df = build_text_gantt_table(filtered_df)
        st.dataframe(text_gantt_df, width="stretch")

# ------------------------------------------------------------
# 明細一覧
# ------------------------------------------------------------
st.subheader("明細一覧")
show_columns = [
    "ID",
    "機能",
    "タスク区分",
    "タスク名",
    "担当",
    "開始日",
    "期限",
    "進捗率",
    "ステータス",
    "優先度",
    "遅延_再計算",
    "課題/懸念",
    "次アクション",
    "更新日",
]

display_df = filtered_df[show_columns].copy()
display_df = display_df.rename(columns={"遅延_再計算": "遅延"})

st.dataframe(display_df, width="stretch")

# ------------------------------------------------------------
# ダウンロード
# ------------------------------------------------------------
st.subheader("CSV ダウンロード")
csv_bytes = display_df.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    label="現在の表示内容を CSV で保存",
    data=csv_bytes,
    file_name="progress_filtered.csv",
    mime="text/csv",
)
