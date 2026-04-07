# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Install dependencies (uses uv) — required before first run and after adding deps
uv sync

# Run the Streamlit app
uv run streamlit run main.py

# Export Gantt chart as a standalone HTML file (browser, interactive)
uv run python gantt_export.py                  # → gantt_output.html
uv run python gantt_export.py output.html      # → arbitrary output path
```

## Architecture

This is a single-file Streamlit app (`main.py`) that visualizes project progress data from an Excel file.

**Data flow:**
1. User uploads an Excel file via the sidebar — the SSOT is the "進捗一覧" sheet (configurable)
2. `detect_header_row()` scans the first 15 rows to find the row that best matches `REQUIRED_COLUMNS`, tolerating title/description rows at the top
3. `load_progress_excel()` reads the sheet with the detected header row (result is `@st.cache_data`-cached by file bytes)
4. `prepare_dataframe()` normalizes strings, coerces dates, normalizes progress rates (handles `40` / `40%` / `0.4` variants), computes delay flags, and adds display/sort helper columns
5. Sidebar filters narrow the DataFrame; summary metrics and a Gantt chart are rendered from the filtered result

**Gantt chart rendering:**
- Primary: Altair (`build_altair_gantt`) — preferred for Japanese text rendering
- Fallback: text pseudo-Gantt table (`build_text_gantt_table`) — used when Altair is unavailable or explicitly selected

**Required Excel columns:** `ID`, `機能`, `タスク区分`, `タスク名`, `担当`, `開始日`, `期限`, `進捗率`, `ステータス`, `優先度`

**Status values:** `未着手`, `進行中`, `完了`, `保留` — used for ordering and color mapping
