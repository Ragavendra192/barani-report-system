from flask import Flask, render_template, request, send_file
import pyodbc
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)

# -------------------- DB SETTINGS --------------------
SERVER = r"RAGAVENDRA\SQLEXPRESS"   # your SQL Server instance
DATABASE = "IIT300"                 # your database name

def get_conn():
    """Return a new SQL Server connection."""
    return pyodbc.connect(
        "DRIVER={ODBC Driver 17 for SQL Server};"
        f"SERVER={SERVER};"
        f"DATABASE={DATABASE};"
        "Trusted_Connection=yes;"
    )

# -------------------- HOME --------------------
@app.route("/")
def home():
    return render_template("home.html")
# -------------------- SHIFT REPORT --------------------
@app.route("/shift-report", methods=["GET", "POST"])
def shift_report():
    # Shift options
    shifts = ["Shift-1", "Shift-2", "Shift-3", "All Shift"]

    from_date = ""
    to_date = ""
    selected_shift = ""
    action = ""
    rows = []
    columns = []
    searched = False

    if request.method == "POST":
        searched = True
        from_date = request.form.get("from_date", "")
        to_date = request.form.get("to_date", "")
        selected_shift = request.form.get("shift", "")
        action = request.form.get("action", "search")

        base_query = """
            SELECT TOP 500
                ID,
                DATE1,
                TIME1,
                BATCHNO,
                RECEIPENAME,
                OPERATORNAME,
                ACKKW,
                ACKKWH
            FROM dbo.ActualLog
            WHERE 1 = 1
        """

        filters = []
        params = []

        # Date filter
        if from_date:
            filters.append(" AND DATE1 >= ?")
            params.append(from_date)
        if to_date:
            filters.append(" AND DATE1 <= ?")
            params.append(to_date)

        # Shift filter based on TIME1
        if selected_shift and selected_shift != "All Shift":
            if selected_shift == "Shift-1":
                filters.append(" AND TIME1 >= '06:00:00' AND TIME1 < '14:00:00'")
            elif selected_shift == "Shift-2":
                filters.append(" AND TIME1 >= '14:00:00' AND TIME1 < '22:00:00'")
            elif selected_shift == "Shift-3":
                filters.append(" AND (TIME1 >= '22:00:00' OR TIME1 < '06:00:00')")

        query = base_query + "".join(filters) + " ORDER BY DATE1, TIME1"

        # Excel export
        if action == "excel":
            with get_conn() as conn:
                df = pd.read_sql(query, conn, params=params)

            filename = f"ShiftReport_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            filepath = os.path.join(os.getcwd(), filename)
            df.to_excel(filepath, index=False)

            return send_file(filepath, as_attachment=True, download_name=filename)

        # Normal HTML table
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute(query, params)
            columns = [c[0] for c in cur.description]
            rows = [dict(zip(columns, r)) for r in cur.fetchall()]

    return render_template(
        "shift_report.html",
        shifts=shifts,
        from_date=from_date,
        to_date=to_date,
        selected_shift=selected_shift,
        rows=rows,
        columns=columns,
        searched=searched,
    )


# -------------------- OPERATOR REPORT --------------------
@app.route("/operator-report", methods=["GET", "POST"])
def operator_report():
    operators = []
    rows = []
    columns = []
    searched = False

    # Load operator names
    with get_conn() as conn:
        df_op = pd.read_sql(
            "SELECT DISTINCT OPERATORNAME FROM dbo.ActualLog "
            "WHERE OPERATORNAME IS NOT NULL AND OPERATORNAME <> '' "
            "ORDER BY OPERATORNAME",
            conn
        )
        operators = df_op["OPERATORNAME"].tolist()

    from_date = request.form.get("from_date", "")
    to_date = request.form.get("to_date", "")
    selected_operator = request.form.get("operator", "")
    action = request.form.get("action", "search")

    if request.method == "POST":
        searched = True

        base_query = """
            SELECT TOP 500
                ID,
                DATE1,
                TIME1,
                BATCHNO,
                RECEIPENAME,
                OPERATORNAME,
                ACKKW,
                ACKKWH
            FROM dbo.ActualLog
            WHERE 1 = 1
        """

        filters = []
        params = []

        if from_date:
            filters.append(" AND DATE1 >= ?")
            params.append(from_date)

        if to_date:
            filters.append(" AND DATE1 <= ?")
            params.append(to_date)

        if selected_operator:
            filters.append(" AND OPERATORNAME = ?")
            params.append(selected_operator)

        query = base_query + "".join(filters) + " ORDER BY DATE1, TIME1"

        # Excel Export
        if action == "excel":
            with get_conn() as conn:
                df = pd.read_sql(query, conn, params=params)
            filename = f"OperatorReport_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            filepath = os.path.join(os.getcwd(), filename)
            df.to_excel(filepath, index=False)
            return send_file(filepath, as_attachment=True)

        # HTML Table Display
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute(query, params)
            columns = [c[0] for c in cur.description]
            rows = [dict(zip(columns, r)) for r in cur.fetchall()]

    return render_template(
        "operator_report.html",
        operators=operators,
        from_date=from_date,
        to_date=to_date,
        selected_operator=selected_operator,
        rows=rows,
        columns=columns,
        searched=searched,
    )


# -------------------- PRODUCT REPORT --------------------
@app.route("/product-report", methods=["GET", "POST"])
def product_report():
    products = []
    rows = []
    columns = []
    searched = False

    # Load product/recipe names from RECEIPENAME
    with get_conn() as conn:
        df_prod = pd.read_sql(
            "SELECT DISTINCT RECEIPENAME FROM dbo.ActualLog "
            "WHERE RECEIPENAME IS NOT NULL AND RECEIPENAME <> '' "
            "ORDER BY RECEIPENAME",
            conn
        )
        products = df_prod["RECEIPENAME"].tolist()

    from_date = request.form.get("from_date", "")
    to_date = request.form.get("to_date", "")
    selected_product = request.form.get("product", "")
    action = request.form.get("action", "search")

    if request.method == "POST":
        searched = True

        base_query = """
            SELECT TOP 500
                ID,
                DATE1,
                TIME1,
                BATCHNO,
                RECEIPENAME,
                OPERATORNAME,
                ACKKW,
                ACKKWH
            FROM dbo.ActualLog
            WHERE 1 = 1
        """

        filters = []
        params = []

        # Date filters
        if from_date:
            filters.append(" AND DATE1 >= ?")
            params.append(from_date)

        if to_date:
            filters.append(" AND DATE1 <= ?")
            params.append(to_date)

        # Product/recipe filter using RECEIPENAME
        if selected_product:
            filters.append(" AND RECEIPENAME = ?")
            params.append(selected_product)

        query = base_query + "".join(filters) + " ORDER BY DATE1, TIME1"

        # Excel export
        if action == "excel":
            with get_conn() as conn:
                df = pd.read_sql(query, conn, params=params)
            filename = f"ProductReport_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            filepath = os.path.join(os.getcwd(), filename)
            df.to_excel(filepath, index=False)
            return send_file(filepath, as_attachment=True)

        # HTML table
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute(query, params)
            columns = [c[0] for c in cur.description]
            rows = [dict(zip(columns, r)) for r in cur.fetchall()]

    return render_template(
        "product_report.html",
        products=products,
        from_date=from_date,
        to_date=to_date,
        selected_product=selected_product,
        rows=rows,
        columns=columns,
        searched=searched,
    )


# START SERVER
if __name__ == "__main__":
    print("Starting Flask server...")
    app.run(debug=True)
