"""
SSOT Excel Generator v1

目的:
進捗管理Excelを完全自動生成する

実行:
python create_ssot_excel.py
"""

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule


def create_master(ws):

    ws["A1"] = "ステータス"
    ws["B1"] = "優先度"
    ws["C1"] = "影響度"

    status = ["未着手", "進行中", "完了", "保留"]
    priority = ["高", "中", "低"]
    impact = ["高", "中", "低"]

    for i, v in enumerate(status, start=2):
        ws[f"A{i}"] = v

    for i, v in enumerate(priority, start=2):
        ws[f"B{i}"] = v

    for i, v in enumerate(impact, start=2):
        ws[f"C{i}"] = v


def dropdown(ws, col, formula):

    dv = DataValidation(type="list", formula1=formula)
    ws.add_data_validation(dv)

    dv.add(f"{col}7:{col}200")


def create_progress(ws):

    headers = [
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
        "遅延",
        "課題/懸念",
        "次アクション",
        "更新日"
    ]

    bold = Font(bold=True)

    for i, h in enumerate(headers, 1):

        cell = ws.cell(row=6, column=i)
        cell.value = h
        cell.font = bold

    dropdown(ws, "I", "=マスタ!$A$2:$A$5")
    dropdown(ws, "J", "=マスタ!$B$2:$B$4")

    dv = DataValidation(type="whole", operator="between",
                        formula1="0", formula2="100")

    ws.add_data_validation(dv)
    dv.add("H7:H200")

    for r in range(7, 200):

        ws[f"K{r}"] = f'=IF(AND(G{r}<TODAY(),I{r}<>"完了"),"遅延","")'

    ws.conditional_formatting.add(
        "I7:I200",
        FormulaRule(formula=['I7="完了"'],
                    fill=None)
    )


def create_weekly(ws):

    ws["A1"] = "週次サマリ"

    ws["A3"] = "完了件数"
    ws["B3"] = '=COUNTIF(進捗一覧!I:I,"完了")'

    ws["A4"] = "進行中件数"
    ws["B4"] = '=COUNTIF(進捗一覧!I:I,"進行中")'

    ws["A5"] = "未着手件数"
    ws["B5"] = '=COUNTIF(進捗一覧!I:I,"未着手")'


def create_risk(ws):

    headers = [
        "ID",
        "分類",
        "内容",
        "影響度",
        "優先度",
        "担当",
        "期限",
        "ステータス"
    ]

    bold = Font(bold=True)

    for i, h in enumerate(headers, 1):

        cell = ws.cell(row=4, column=i)
        cell.value = h
        cell.font = bold

    dropdown(ws, "D", "=マスタ!$C$2:$C$4")
    dropdown(ws, "E", "=マスタ!$B$2:$B$4")
    dropdown(ws, "H", "=マスタ!$A$2:$A$5")


def main():

    wb = Workbook()

    progress = wb.active
    progress.title = "進捗一覧"

    master = wb.create_sheet("マスタ")
    weekly = wb.create_sheet("週次サマリ")
    risk = wb.create_sheet("課題・リスク")

    create_master(master)
    create_progress(progress)
    create_weekly(weekly)
    create_risk(risk)

    wb.save("progress_template_ssot.xlsx")

    print("SSOT Excel generated")


if __name__ == "__main__":
    main()