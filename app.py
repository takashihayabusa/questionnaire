from flask import Flask, render_template, request
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_PATH = os.path.join(BASE_DIR, "survey_results.xlsx")


# ==========================
# Excel保存
# ==========================
def save_to_excel(data):

    headers = [
        "日時",
        "店舗名",
        "部門",
        "名前",
        "Q1",
        "Q2",
        "Q3",
        "Q4",
        "Q5",
        "Q6",
        "Q7",
        "Q8",
        "Q9",
        "Q10",
        "ハラスメント詳細",
        "相談したか",
        "結果",
        "会社要望",
        "組合要望"
    ]

    # ファイルがなければ新規作成
    if not os.path.exists(FILE_PATH):
        wb = Workbook()
        ws = wb.active
        ws.append(headers)
        wb.save(FILE_PATH)

    wb = load_workbook(FILE_PATH)
    ws = wb.active

    ws.append(data)
    wb.save(FILE_PATH)


# ==========================
# 画面表示
# ==========================
@app.route("/survey")
def index():
    return render_template("survey.html")


# ==========================
# 送信処理
# ==========================
@app.route("/submit", methods=["POST"])
def submit():

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    store = request.form.get("store")
    department = request.form.get("department")
    name = request.form.get("name")

    q1 = request.form.get("q1")

    # Q2は固定残業あり・なしどちらかを採用
    q2_fixed = request.form.get("q2")
    q2_no_fixed = request.form.get("q2b")
    q2 = q2_fixed if q2_fixed else q2_no_fixed

    q3 = request.form.get("q3")
    q4 = request.form.get("q4")
    q5 = request.form.get("q5")
    q6 = request.form.get("q6")
    q7 = request.form.get("q7")

    q8 = request.form.get("q8")
    q9 = request.form.get("q9")

    q10 = request.form.get("q10")
    harassment_detail = request.form.get("harassment_detail")
    consult = request.form.get("consult")
    result = request.form.get("result")

    company_request = request.form.get("company_request")
    union_request = request.form.get("union_request")

    row_data = [
        now,
        store,
        department,
        name,
        q1,
        q2,
        q3,
        q4,
        q5,
        q6,
        q7,
        q8,
        q9,
        q10,
        harassment_detail,
        consult,
        result,
        company_request,
        union_request
    ]

    save_to_excel(row_data)

    return render_template("thanks.html")


# ==========================
# 起動
# ==========================
if __name__ == "__main__":
    app.run(debug=True)