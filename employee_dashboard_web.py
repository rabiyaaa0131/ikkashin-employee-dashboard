from datetime import date
from pathlib import Path

import pandas as pd
from flask import Flask, request, render_template_string

app = Flask(__name__)
BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "Daily Progress Tracker@RABIYA.xlsx"

COLUMNS = [
    "Employee ID", "Employee Name", "Date", "Login Time", "Logout Time", "Task",
    "Priority", "Task Status", "Completion %", "Today's Work", "Pending Task",
    "Pending Reason", "Next Step", "Expected Date", "Dependency / Help Needed",
    "Time Spent on Task", "Remarks", "Estimated Time", "Actual Time", "Waiting Time",
    "Reason for Delay", "Task Source", "Work Type", "Carry Forward",
    "Priority Changed?", "Tomorrow's Priority"
]

HTML = r'''<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Employee Tracking Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
*{box-sizing:border-box}body{margin:0;background:#f5f7fb;color:#172033;font-family:Arial,sans-serif}header{background:#172554;color:white;padding:22px 26px}header h1{margin:0 0 6px;font-size:24px}header p{margin:0;color:#dbeafe;font-size:13px}main{max-width:1500px;margin:20px auto;padding:0 16px}.panel,.metric,.filters{background:white;border:1px solid #e2e8f0;border-radius:10px}.filters{display:flex;flex-wrap:wrap;gap:12px;align-items:end;padding:15px;margin-bottom:15px}.field{display:flex;flex-direction:column;gap:6px;min-width:145px;flex:1}.field label{font-size:12px;color:#64748b;font-weight:bold}input,select,button{font:inherit;padding:10px;border:1px solid #cbd5e1;border-radius:6px;background:white;color:#172033}button{background:#2563eb;color:white;border-color:#2563eb;cursor:pointer;font-weight:bold}button.secondary{background:white;color:#2563eb}.metrics{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-bottom:15px}.metric{padding:16px}.metric .label{font-size:12px;color:#64748b;margin-bottom:8px}.metric .value{font-size:25px;font-weight:bold;color:#172554}.two{display:grid;grid-template-columns:1fr 1fr;gap:15px}.panel{padding:16px;margin-bottom:15px;min-width:0}.panel h2{font-size:16px;margin:0 0 12px}.chart{height:240px;position:relative}.table-wrap{overflow:auto;max-height:600px;border:1px solid #e2e8f0;border-radius:7px}table{border-collapse:collapse;width:100%;font-size:12px;white-space:nowrap}th,td{padding:9px 10px;border-bottom:1px solid #e2e8f0;text-align:left;max-width:320px;overflow:hidden;text-overflow:ellipsis}th{position:sticky;top:0;background:#eef2ff;color:#1e3a8a}tr:hover td{background:#f8fafc}.note{font-size:12px;line-height:1.5;color:#64748b}.empty{padding:24px;text-align:center;color:#64748b}@media(max-width:800px){.metrics{grid-template-columns:repeat(2,1fr)}.two{grid-template-columns:1fr}}@media print{.filters,button{display:none!important}body{background:white}header{background:white;color:black;border-bottom:1px solid #aaa}header p{color:#333}main{max-width:none;margin:0}.panel,.metric{break-inside:avoid}.table-wrap{max-height:none;overflow:visible}th{position:static}}
</style></head><body>
<header><h1>Employee Tracking Dashboard</h1><p>Employee tracking records · Daily, weekly and monthly analysis</p></header>
<main>
<form class="filters" method="get">
<div class="field"><label for="period">Statistics period</label><select id="period" name="period"><option value="daily" {% if period=='daily' %}selected{% endif %}>Daily</option><option value="weekly" {% if period=='weekly' %}selected{% endif %}>Weekly</option><option value="monthly" {% if period=='monthly' %}selected{% endif %}>Monthly</option></select></div>
<div class="field"><label for="selected_date">Date in period</label><input id="selected_date" name="selected_date" type="date" value="{{ selected_date }}"></div>
<div class="field"><label for="employee">Employee</label><select id="employee" name="employee"><option value="all">All employees</option>{% for e in employees %}<option value="{{ e }}" {% if employee==e %}selected{% endif %}>{{ e }}</option>{% endfor %}</select></div>
<div class="field"><label for="status">Task status</label><select id="status" name="status"><option value="all">All statuses</option>{% for s in statuses %}<option value="{{ s }}" {% if status==s %}selected{% endif %}>{{ s }}</option>{% endfor %}</select></div>
<button type="submit">Apply filters</button><button type="button" class="secondary" onclick="window.print()">Print view</button>
</form>
<div class="metrics"><div class="metric"><div class="label">Task records</div><div class="value">{{ total }}</div></div><div class="metric"><div class="label">Completed tasks</div><div class="value">{{ completed }}</div></div><div class="metric"><div class="label">Not completed</div><div class="value">{{ not_completed }}</div></div><div class="metric"><div class="label">Average completion</div><div class="value">{{ avg_completion }}%</div></div></div>
<div class="two"><section class="panel"><h2>Task status breakdown</h2><div class="chart"><canvas id="statusChart"></canvas></div></section><section class="panel"><h2>Tasks by priority</h2><div class="chart"><canvas id="priorityChart"></canvas></div></section></div>
<section class="panel"><h2>Employee tracking records</h2><p class="note">All available columns from the employee tracker are shown below for the selected period. Blank/no-task/leave/Sunday entries are not counted as tasks.</p>
{% if rows %}<div class="table-wrap"><table><thead><tr>{% for c in columns %}<th title="{{ c }}">{{ c }}</th>{% endfor %}</tr></thead><tbody>{% for row in rows %}<tr>{% for cell in row %}<td title="{{ cell }}">{{ cell }}</td>{% endfor %}</tr>{% endfor %}</tbody></table></div>{% else %}<div class="empty">No task records match the selected filters.</div>{% endif %}</section>
<section class="panel"><h2>Data note</h2><p class="note">This dashboard analyses the fields currently present in the daily tracking workbook. The reference document defines separate KPI/KPA calculations that require additional source logs (such as attendance, session, admin-task, sprint-task, communication and conduct logs). Scores are not fabricated from fields that are not available in this tracker.</p></section>
</main><script>
const statusLabels={{ status_labels|tojson }};const statusValues={{ status_values|tojson }};const priorityLabels={{ priority_labels|tojson }};const priorityValues={{ priority_values|tojson }};
if(window.Chart){new Chart(document.getElementById('statusChart'),{type:'bar',data:{labels:statusLabels,datasets:[{label:'Task records',data:statusValues}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{precision:0}}}}});new Chart(document.getElementById('priorityChart'),{type:'doughnut',data:{labels:priorityLabels,datasets:[{data:priorityValues}]},options:{responsive:true,maintainAspectRatio:false}})}
</script></body></html>'''


def parse_date(value):
    """Parse Excel dates and mixed text dates, including 02-0CT-2026 typos."""
    if pd.isna(value) or str(value).strip() == "":
        return pd.NaT
    if isinstance(value, (pd.Timestamp, date)):
        return pd.to_datetime(value, errors="coerce")
    text = str(value).strip().upper().replace("0CT", "OCT").replace("OCT", "OCT")
    return pd.to_datetime(text, errors="coerce", dayfirst=True)


def load_data():
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE.name}. Keep the Excel tracker in the same folder as this Python file."
        )
    data = pd.read_excel(EXCEL_FILE)
    data.columns = [str(c).strip() for c in data.columns]
    for column in COLUMNS:
        if column not in data.columns:
            data[column] = ""
    data = data[COLUMNS].copy()
    data["Date"] = data["Date"].apply(parse_date)
    data["Expected Date"] = data["Expected Date"].apply(parse_date)
    data = data.dropna(subset=["Date"]).copy()
    data["Employee Name"] = data["Employee Name"].fillna("").astype(str).str.strip()
    data["Task"] = data["Task"].fillna("").astype(str).str.strip()
    data["Task Status"] = data["Task Status"].fillna("Not specified").astype(str).str.strip()
    data["Priority"] = data["Priority"].fillna("Not specified").astype(str).str.strip()
    return data


def pct_value(value):
    if pd.isna(value):
        return None
    raw = str(value).strip().replace("%", "").replace("~", "")
    try:
        number = float(raw)
        if number <= 1 and "%" not in str(value):
            number *= 100
        return max(0.0, min(100.0, number))
    except (TypeError, ValueError):
        return None


def display_value(value):
    if pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%d-%b-%Y")
    return str(value)


@app.route("/")
def dashboard():
    try:
        df = load_data()
    except Exception as exc:
        return ("<h2>Dashboard could not load the tracker</h2><p>" + str(exc) +
                "</p><p>Keep the Excel file beside employee_dashboard_web.py and refresh.</p>"), 500

    period = request.args.get("period", "daily").lower()
    if period not in {"daily", "weekly", "monthly"}:
        period = "daily"
    try:
        anchor = pd.Timestamp(request.args.get("selected_date", date.today().isoformat()))
    except (TypeError, ValueError):
        anchor = pd.Timestamp(date.today())

    employee = request.args.get("employee", "all")
    status = request.args.get("status", "all")
    view = df.copy()
    if period == "daily":
        view = view[view["Date"].dt.date == anchor.date()]
    elif period == "weekly":
        start = anchor.normalize() - pd.Timedelta(days=anchor.weekday())
        end = start + pd.Timedelta(days=7)
        view = view[(view["Date"] >= start) & (view["Date"] < end)]
    else:
        view = view[(view["Date"].dt.year == anchor.year) & (view["Date"].dt.month == anchor.month)]

    employees = sorted(name for name in df["Employee Name"].unique().tolist() if name)
    statuses = sorted(status_name for status_name in df["Task Status"].unique().tolist() if status_name)
    if employee != "all":
        view = view[view["Employee Name"] == employee]
    if status != "all":
        view = view[view["Task Status"] == status]

    actual_task = ~view["Task"].str.lower().isin(["", "nan", "no task", "leave", "sunday", "-", "not applicable"])
    tasks = view[actual_task].copy()
    completion = tasks["Completion %"].apply(pct_value).dropna()
    completed = int(tasks["Task Status"].str.lower().str.contains("complete", na=False).sum())
    not_completed = max(0, len(tasks) - completed)
    status_counts = tasks["Task Status"].replace("", "Not specified").value_counts()
    priority_counts = tasks["Priority"].replace("", "Not specified").value_counts()
    rows = [[display_value(v) for v in row] for row in view[COLUMNS].itertuples(index=False, name=None)]

    return render_template_string(
        HTML, period=period, selected_date=anchor.strftime("%Y-%m-%d"), employee=employee, status=status,
        employees=employees, statuses=statuses, total=len(tasks), completed=completed,
        not_completed=not_completed, avg_completion=round(float(completion.mean()), 1) if len(completion) else 0,
        columns=COLUMNS, rows=rows, status_labels=status_counts.index.tolist(),
        status_values=status_counts.astype(int).tolist(), priority_labels=priority_counts.index.tolist(),
        priority_values=priority_counts.astype(int).tolist()
    )


if __name__ == "__main__":
    print("Employee Tracking Dashboard: http://127.0.0.1:5000")
    app.run(host="127.0.0.1", port=5000, debug=False)
