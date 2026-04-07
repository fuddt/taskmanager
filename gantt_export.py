"""
gantt_export.py — Excel からガントチャートを HTML ファイルとして出力するスクリプト

使い方:
    uv run python gantt_export.py              # → gantt_output.html
    uv run python gantt_export.py output.html  # → 任意のパスに出力

出力した HTML はブラウザで開くとインタラクティブなガントチャートが表示されます。
「日数 × 15px」の明示的な幅で描画するため、ブラウザの表示幅に制限されず
1日単位の目盛りが確実に表示されます。

追加パッケージ不要（Altair は Streamlit の依存としてインストール済み）。
"""

import os
import sys
import tomllib
from pathlib import Path

import altair as alt
import pandas as pd

# ============================================================
# 定数（main.py と共通）
# ============================================================

REQUIRED_COLUMNS = [
    "ID", "機能", "タスク区分", "タスク名", "担当",
    "開始日", "期限", "進捗率", "ステータス", "優先度",
]

STATUS_COLOR_MAP = {
    "未着手": "#BDBDBD",
    "進行中": "#1F77B4",
    "完了":   "#2CA02C",
    "保留":   "#7F7F7F",
}

# ============================================================
# データ処理関数（main.py より流用 ／ Streamlit 依存なし）
# ============================================================

def normalize_progress(value) -> float:
    if pd.isna(value):
        return 0.0
    if isinstance(value, str):
        cleaned = value.replace("%", "").strip()
        if not cleaned:
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
    preview_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=preview_rows)
    best_row_index, best_match_count = 0, -1
    for row_idx in range(len(preview_df)):
        row_values = {str(v).strip() for v in preview_df.iloc[row_idx].tolist() if pd.notna(v)}
        match_count = sum(1 for col in REQUIRED_COLUMNS if col in row_values)
        if match_count > best_match_count:
            best_match_count = match_count
            best_row_index = row_idx
    return best_row_index


def load_progress_excel(file_path: str, sheet_name: str) -> pd.DataFrame:
    header_row = detect_header_row(file_path, sheet_name)
    return pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)


def validate_columns(df: pd.DataFrame) -> list[str]:
    return [col for col in REQUIRED_COLUMNS if col not in df.columns]


def prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()

    string_cols = ["ID", "機能", "タスク区分", "タスク名", "担当",
                   "ステータス", "優先度", "課題/懸念", "次アクション"]
    for col in string_cols:
        if col in work.columns:
            work[col] = work[col].astype("string").str.strip()

    work["開始日"] = pd.to_datetime(work["開始日"], errors="coerce")
    work["期限"]   = pd.to_datetime(work["期限"],   errors="coerce")
    if "更新日" in work.columns:
        work["更新日"] = pd.to_datetime(work["更新日"], errors="coerce")

    work["進捗率"] = work["進捗率"].apply(normalize_progress)

    work["ステータス"] = (
        work["ステータス"].astype("string").str.strip()
        .replace({pd.NA: "未着手", "": "未着手", "nan": "未着手", "None": "未着手", "<NA>": "未着手"})
        .fillna("未着手")
    )
    work["優先度"] = work["優先度"].replace(
        {pd.NA: "", "nan": "", "None": "", "<NA>": ""}
    ).fillna("")

    today = pd.Timestamp.today().normalize()
    work["遅延_再計算"] = work.apply(
        lambda row: "遅延"
        if pd.notna(row["期限"]) and row["期限"] < today and row["ステータス"] != "完了"
        else "",
        axis=1,
    )

    work["表示開始日"] = work["開始日"]
    work.loc[work["表示開始日"].isna(), "表示開始日"] = work.loc[work["表示開始日"].isna(), "期限"]
    work["表示開始日"] = work["表示開始日"].fillna(today)
    work["表示期限"]   = work["期限"].fillna(work["表示開始日"])

    mask = work["表示開始日"] > work["表示期限"]
    work.loc[mask, "表示期限"] = work.loc[mask, "表示開始日"]

    work["表示名"] = work.apply(
        lambda row: f"[{row['機能']}] {row['タスク名']}（{row['担当']}）", axis=1
    )

    # ID 順ソート（数値IDは数値比較、英数字IDは文字列比較にフォールバック）
    _id_num = pd.to_numeric(work["ID"], errors="coerce")
    work["_id_sort"] = _id_num
    work = work.sort_values(by=["_id_sort", "ID"], ascending=[True, True], na_position="last")
    work = work.drop(columns=["_id_sort"])
    return work


# ============================================================
# Altair ガントチャート（HTML 出力専用）
# ============================================================

def build_altair_gantt_for_export(df: pd.DataFrame) -> alt.Chart:
    """
    ファイル出力用 Altair ガントチャート。

    Streamlit 版と異なり、日数 × 15px の明示的な幅を .properties() に指定する。
    これにより Vega-Lite が「画面幅に合わせて目盛りを間引く」動作を回避し、
    tickCount={"interval": "day", "step": 1} が確実に1日単位で描画される。
    """
    chart_df = df.copy()
    chart_df["表示名_遅延反映"] = chart_df.apply(
        lambda row: f"[遅延] {row['表示名']}" if row["遅延_再計算"] == "遅延" else row["表示名"],
        axis=1,
    )
    y_order = chart_df["表示名_遅延反映"].tolist()

    date_min  = chart_df["表示開始日"].min()
    date_max  = chart_df["表示期限"].max()
    day_range = max(7, (date_max - date_min).days + 1)

    # 1日 15px で幅を固定する（ブラウザ横スクロールで全体を確認可能）
    chart_width = day_range * 15

    bars = (
        alt.Chart(chart_df)
        .mark_bar(size=22)
        .encode(
            x=alt.X(
                "表示開始日:T",
                title="日付",
                axis=alt.Axis(
                    format="%m/%d",
                    tickCount={"interval": "day", "step": 1},
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
                alt.Tooltip("ID:N",           title="ID"),
                alt.Tooltip("機能:N",          title="機能"),
                alt.Tooltip("タスク区分:N",    title="タスク区分"),
                alt.Tooltip("タスク名:N",      title="タスク名"),
                alt.Tooltip("担当:N",          title="担当"),
                alt.Tooltip("表示開始日:T",    title="開始日"),
                alt.Tooltip("表示期限:T",      title="期限"),
                alt.Tooltip("進捗率:Q",        title="進捗率", format=".1f"),
                alt.Tooltip("優先度:N",        title="優先度"),
                alt.Tooltip("ステータス:N",    title="ステータス"),
                alt.Tooltip("遅延_再計算:N",   title="遅延"),
            ],
        )
    )

    today_df   = pd.DataFrame({"today": [pd.Timestamp.today().normalize()]})
    today_rule = (
        alt.Chart(today_df)
        .mark_rule(strokeDash=[6, 4], size=2, color="red")
        .encode(x="today:T")
    )

    chart = (bars + today_rule).properties(
        width=chart_width,
        height=max(400, len(chart_df) * 32),
        title="進捗ガントチャート",
    )

    return chart


# ============================================================
# エントリポイント
# ============================================================

if __name__ == "__main__":
    output_path = (
        sys.argv[1]
        if len(sys.argv) > 1
        else str(Path(__file__).parent / "gantt_output.html")
    )

    # ---- config.toml 読み込み ----
    config_path = Path(__file__).parent / "config.toml"
    if not config_path.exists():
        print(f"[ERROR] config.toml が見つかりません: {config_path}", file=sys.stderr)
        sys.exit(1)

    try:
        with open(config_path, "rb") as f:
            config = tomllib.load(f)
    except tomllib.TOMLDecodeError as e:
        print(f"[ERROR] config.toml の解析に失敗しました: {e}", file=sys.stderr)
        print("Windows パスはスラッシュ(/) かシングルクォート('...')で記述してください。",
              file=sys.stderr)
        sys.exit(1)

    excel_cfg  = config.get("excel", {})
    file_path  = excel_cfg.get("file_path", "")
    sheet_name = excel_cfg.get("sheet_name", "進捗一覧")

    if not file_path or not os.path.exists(file_path):
        print(f"[ERROR] Excel ファイルが見つかりません: {file_path}", file=sys.stderr)
        sys.exit(1)

    # ---- データ読み込み ----
    print(f"読み込み中: {file_path}  シート: [{sheet_name}]")
    raw_df = load_progress_excel(file_path, sheet_name)

    missing = validate_columns(raw_df)
    if missing:
        print(f"[ERROR] 必須列が不足しています: {missing}", file=sys.stderr)
        sys.exit(1)

    df = prepare_dataframe(raw_df)

    date_min  = df["表示開始日"].min()
    date_max  = df["表示期限"].max()
    day_range = max(7, (date_max - date_min).days + 1)

    # ---- ガントチャート生成・保存 ----
    chart = build_altair_gantt_for_export(df)

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    chart.save(str(out))

    print(f"出力完了: {out.resolve()}")
    print(f"  タスク数: {len(df)}  /  期間: {day_range} 日  /  チャート幅: {day_range * 15}px")
    print(f"  ブラウザで開いてください。")
