from flask import Flask, render_template, request
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os

app = Flask(__name__)

# ==============================
# 保存フォルダ
# ==============================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SAVE_DIR = os.path.join(BASE_DIR, "questionnaire")
os.makedirs(SAVE_DIR, exist_ok=True)

FILE_PATH = os.path.join(SAVE_DIR, "survey_results.xlsx")


# ==============================
# Excel保存
# ==============================

def save_to_excel(data):

    headers = [
        "日時",
        "店",
        "部門",
        "名前",
        "週休",
        "固定残業",
        "固定残業時間",
        "サービス残業",
        "有給休暇",
        "有給理由",
        "ハラスメント",
        "ハラスメント内容",
        "相談",
        "解決",
        "目撃",
        "会社評価",
        "会社要望",
        "組合要望"
    ]

    if not os.path.exists(FILE_PATH):

        wb = Workbook()
        ws = wb.active
        ws.append(headers)
        wb.save(FILE_PATH)

    wb = load_workbook(FILE_PATH)
    ws = wb.active

    ws.append(data)

    wb.save(FILE_PATH)


# ==============================
# アンケート画面
# ==============================

@app.route("/")
@app.route("/survey")
def survey():
    return render_template("survey.html")


# ==============================
# 送信
# ==============================

@app.route("/submit", methods=["POST"])
def submit():

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    store = request.form.get("store")
    dept = request.form.get("dept")
    name = request.form.get("name")

    q1 = request.form.get("q1")

    overtime = request.form.get("overtime")

    q3 = request.form.get("q3")
    q4 = request.form.get("q4")

    service = request.form.get("service")

    vacation = request.form.get("vacation")
    reason = request.form.get("reason")

    harass = request.form.get("harass")
    harass_text = request.form.get("harass_text")

    consult = request.form.get("consult")
    solve = request.form.get("solve")

    seen = request.form.get("seen")

    company = request.form.get("company")

    company_request = request.form.get("company_request")
    union_request = request.form.get("union_request")

    save_to_excel([
        now,
        store,
        dept,
        name,
        q1,
        overtime,
        q3 or q4,
        service,
        vacation,
        reason,
        harass,
        harass_text,
        consult,
        solve,
        seen,
        company,
        company_request,
        union_request
    ])

    return "<h2 style='text-align:center;'>ご協力ありがとうございました</h2>"


# ==============================

if __name__ == "__main__":
    app.run()