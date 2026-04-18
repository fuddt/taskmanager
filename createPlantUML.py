from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, date
from pathlib import Path
from typing import Dict, List, Optional

from openpyxl import load_workbook


INPUT_FILE = "progress_template_plantuml_ssot.xlsx"
OUTPUT_FILE = "progress_gantt.puml"
SHEET_NAME = "進捗一覧"
HEADER_ROW = 6
DATA_START_ROW = 7


@dataclass
class TaskRow:
    task_id: str
    function_name: str
    task_category: str
    task_name: str
    display_name: str
    owner: str
    start_date: date
    end_date: date
    progress: int
    status: str
    priority: str
    display_order: int
    depends_on: List[str]
    task_type: str
    delay: str
    concern: str
    next_action: str
    updated_at: Optional[date]


def normalize_str(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def parse_date(value) -> Optional[date]:
    if value is None or value == "":
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    text = str(value).strip()

    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass

    return None


def parse_int(value, default: int = 0) -> int:
    if value is None or value == "":
        return default

    try:
        return int(float(value))
    except Exception:
        return default


def parse_depends(depends_text: str) -> List[str]:
    if not depends_text:
        return []

    return [x.strip() for x in depends_text.split(",") if x.strip()]


def format_date_for_plantuml(d: date) -> str:
    return d.strftime("%Y-%m-%d")


def sanitize_label(text: str) -> str:
    s = normalize_str(text)
    s = s.replace("[", "").replace("]", "")
    s = s.replace(" ", "_")
    return s if s else "NO_NAME"


def build_unique_label(task: TaskRow) -> str:
    base = sanitize_label(task.display_name or task.task_name)
    return f"{base}_ID_{task.task_id}"


def status_to_color(status: str) -> str:
    if status == "完了":
        return "LightGreen"
    elif status == "進行中":
        return "LightBlue"
    else:
        return "LightGray"


def build_header_map(ws) -> Dict[str, int]:
    header_map: Dict[str, int] = {}

    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=HEADER_ROW, column=col).value
        key = normalize_str(value)

        if key:
            header_map[key] = col

    return header_map


def required_columns() -> List[str]:
    return [
        "ID",
        "機能",
        "タスク区分",
        "タスク名",
        "表示名",
        "担当",
        "開始日",
        "期限",
        "進捗率",
        "ステータス",
        "優先度",
        "表示順",
        "依存タスクID",
        "タスク種別",
        "遅延",
        "課題/懸念",
        "次アクション",
        "更新日",
    ]


def validate_required_columns(header_map):
    missing = [c for c in required_columns() if c not in header_map]

    if missing:
        raise ValueError("必須列が見つかりません: " + ", ".join(missing))


def is_empty_row(ws, row_idx, header_map):
    keys = ["ID", "タスク名", "開始日", "期限"]

    for k in keys:
        if ws.cell(row=row_idx, column=header_map[k]).value:
            return False

    return True


def read_tasks_from_excel(file_path: str) -> List[TaskRow]:
    wb = load_workbook(file_path, data_only=False)

    if SHEET_NAME not in wb.sheetnames:
        raise ValueError("進捗一覧シートが存在しません")

    ws = wb[SHEET_NAME]

    header_map = build_header_map(ws)

    validate_required_columns(header_map)

    tasks = []

    for row_idx in range(DATA_START_ROW, ws.max_row + 1):

        if is_empty_row(ws, row_idx, header_map):
            continue

        task = TaskRow(
            task_id=normalize_str(ws.cell(row=row_idx, column=header_map["ID"]).value),
            function_name=normalize_str(ws.cell(row=row_idx, column=header_map["機能"]).value),
            task_category=normalize_str(ws.cell(row=row_idx, column=header_map["タスク区分"]).value),
            task_name=normalize_str(ws.cell(row=row_idx, column=header_map["タスク名"]).value),
            display_name=normalize_str(ws.cell(row=row_idx, column=header_map["表示名"]).value),
            owner=normalize_str(ws.cell(row=row_idx, column=header_map["担当"]).value),
            start_date=parse_date(ws.cell(row=row_idx, column=header_map["開始日"]).value),
            end_date=parse_date(ws.cell(row=row_idx, column=header_map["期限"]).value),
            progress=parse_int(ws.cell(row=row_idx, column=header_map["進捗率"]).value),
            status=normalize_str(ws.cell(row=row_idx, column=header_map["ステータス"]).value),
            priority=normalize_str(ws.cell(row=row_idx, column=header_map["優先度"]).value),
            display_order=parse_int(ws.cell(row=row_idx, column=header_map["表示順"]).value),
            depends_on=parse_depends(normalize_str(ws.cell(row=row_idx, column=header_map["依存タスクID"]).value)),
            task_type=normalize_str(ws.cell(row=row_idx, column=header_map["タスク種別"]).value).lower(),
            delay=normalize_str(ws.cell(row=row_idx, column=header_map["遅延"]).value),
            concern=normalize_str(ws.cell(row=row_idx, column=header_map["課題/懸念"]).value),
            next_action=normalize_str(ws.cell(row=row_idx, column=header_map["次アクション"]).value),
            updated_at=parse_date(ws.cell(row=row_idx, column=header_map["更新日"]).value),
        )

        if not task.display_name:
            task.display_name = task.task_name

        tasks.append(task)

    return tasks


def find_project_start(tasks):
    return min(task.start_date for task in tasks)


def generate_plantuml(tasks: List[TaskRow]):

    tasks = sorted(tasks, key=lambda x: (x.display_order, x.start_date))

    project_start = find_project_start(tasks)

    label_map = {t.task_id: build_unique_label(t) for t in tasks}

    lines = []

    lines.append("@startgantt")
    lines.append("")

    # ===== UI改善 =====

    lines.append("printscale daily zoom 2")
    lines.append("today is colored in Red")
    lines.append("")

    lines.append("legend")
    lines.append("|= Color |= 状態 |")
    lines.append("|<#LightGreen>| 完了 |")
    lines.append("|<#LightBlue>| 進行中 |")
    lines.append("|<#LightGray>| 未着手 |")
    lines.append("end legend")
    lines.append("")

    lines.append("title Progress Gantt")

    lines.append(f"Project starts {format_date_for_plantuml(project_start)}")

    lines.append("saturday are closed")
    lines.append("sunday are closed")
    lines.append("")

    current_category = None

    for task in tasks:

        if task.task_category != current_category:
            current_category = task.task_category
            lines.append(f"-- {current_category} --")

        label = label_map[task.task_id]

        start_str = format_date_for_plantuml(task.start_date)

        duration = (task.end_date - task.start_date).days + 1

        lines.append(f"[{label}] starts {start_str}")
        lines.append(f"[{label}] requires {duration} days")

        lines.append(f"[{label}] is {task.progress}% completed")

        color = status_to_color(task.status)

        lines.append(f"[{label}] is colored in {color}")

        for dep in task.depends_on:

            if dep in label_map:

                dep_label = label_map[dep]

                lines.append(f"[{label}] starts after [{dep_label}]'s end")

        lines.append("")

    lines.append("@endgantt")

    return "\n".join(lines)


def main():

    tasks = read_tasks_from_excel(INPUT_FILE)

    puml = generate_plantuml(tasks)

    Path(OUTPUT_FILE).write_text(puml, encoding="utf-8")

    print("PlantUML生成完了:", OUTPUT_FILE)


if __name__ == "__main__":
    main()