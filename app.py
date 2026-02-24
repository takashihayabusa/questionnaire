from flask import Flask, render_template, request
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_PATH = os.path.join(BASE_DIR, "survey_results.xlsx")

# =========================
# Excel保存（列ズレ完全防止版）
# =========================
def save_to_excel(row_data):

    headers = [
        "日時",
        "LINE名",
        "手入力名",
        "Q1",
        "Q2",
        "Q3",
        "Q4",
        "Q5",
        "Q6",
        "Q7理由",
        "ハラスメント",
        "会社評価",
        "会社への要望",
        "組合への要望"
    ]

    # ファイルがなければ新規作成
    if not os.path.exists(FILE_PATH):
        wb = Workbook()
        ws = wb.active
        ws.append(headers)
        wb.save(FILE_PATH)

    wb = load_workbook(FILE_PATH)
    ws = wb.active

    # 列数を強制的に14に揃える
    fixed_row = row_data[:14]  # 多すぎる場合カット
    while len(fixed_row) < 14:  # 足りない場合空白追加
        fixed_row.append("")

    ws.append(fixed_row)
    wb.save(FILE_PATH)

# =========================
# 表紙
# =========================
@app.route("/")
def cover():
    return render_template("cover.html")

# =========================
# アンケート
# =========================
@app.route("/survey", methods=["GET", "POST"])
def survey():

    if request.method == "POST":

        # 名前
        line_name = request.form.get("line_name") or ""
        manual_name = request.form.get("manual_name") or ""

        # 設問
        q1 = request.form.get("q1") or ""
        q2 = request.form.get("q2") or request.form.get("q2b") or ""
        q3 = request.form.get("q3") or ""
        q4 = request.form.get("q4") or ""
        q5 = request.form.get("q5") or ""
        q6 = request.form.get("q6") or ""
        q7_reason = request.form.get("q7_reason") or ""

        harassment = request.form.get("harassment") or ""
        company_eval = request.form.get("company_eval") or ""

        company_request = request.form.get("company_request") or ""
        union_request = request.form.get("union_request") or ""

        save_to_excel([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            line_name,
            manual_name,
            q1,
            q2,
            q3,
            q4,
            q5,
            q6,
            q7_reason,
            harassment,
            company_eval,
            company_request,
            union_request
        ])

        return render_template("complete.html")

    return render_template("survey.html")


if __name__ == "__main__":
    app.run()