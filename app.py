from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import pickle, os
import pandas as pd

app = Flask(__name__)
app.secret_key = "supersecretkey"
DATA_FILE = "applications.pkl"
os.makedirs(os.path.dirname(DATA_FILE) or ".", exist_ok=True)

if not os.path.exists(DATA_FILE):
    with open(DATA_FILE, 'wb') as f:
        pickle.dump({}, f)

ADMIN_USERNAME = "林毅"
ADMIN_PASSWORD = "Linlin520"

def load_data():
    with open(DATA_FILE, 'rb') as f:
        return pickle.load(f)

def save_data(data):
    with open(DATA_FILE, 'wb') as f:
        pickle.dump(data, f)

@app.route("/", methods=["GET", "POST"])
def submit_application():
    if request.method == "POST":
        name = request.form["name"]
        application_time = request.form["application_time"]
        remarks = request.form["remarks"]

        data = load_data()
        data[name] = {"time": application_time, "remarks": remarks}
        save_data(data)
        flash("您的申请已经提交成功！", "success")
        return redirect(url_for("submit_application"))
    return render_template("user_form.html")

@app.route("/admin", methods=["GET", "POST"])
def admin_interface():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]
        if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
            data = load_data()
            return render_template("admin_dashboard.html", data=data)
        else:
            flash("用户名或密码错误", "danger")
            return redirect(url_for("admin_interface"))
    return render_template("admin_login.html")

@app.route("/export")
def export_excel():
    data = load_data()
    df = pd.DataFrame.from_dict(data, orient="index")
    df.index.name = "姓名"
    excel_path = "applications.xlsx"
    df.to_excel(excel_path)
    return send_file(excel_path, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
