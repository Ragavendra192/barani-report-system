from flask import Flask, render_template, request, send_file
import pyodbc
import pandas as pd
from datetime import datetime
import os
import traceback
import tempfile

app = Flask(__name__)

# -------------------- DB SETTINGS --------------------
# Use environment variables if available (for future server/IT setup),
# otherwise fall back to your local SQL Server instance.
SERVER = os.getenv("SQL_SERVER", r"RAGAVENDRA\SQLEXPRESS")
DATABASE = os.getenv("SQL_DATABASE", "IIT300")


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
    shifts = ["Shift-1", "Shift-2", "Shift-3", "All Shift"]

    from_date = ""
    to_date = ""
    selected_shift = ""
    action = ""
    rows = []
    columns = []
    searched = False
    error_message = None

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

        # Shift filter
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
            try:
                with get_conn() as conn:
                    df = pd.read_sql(query, conn, params=params)

                filename = f"ShiftReport_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                # Use a temp folder that works on both Windows & Linux
                filepath = os.path.join(tempfile.gettempdir(), filename)
                df.to_excel(filepath, index=False)

                return send_file(filepath, as_attachment=True, download_name=filename)
            except Exception as e:
                error_message = "Could not generate Excel report."
                print("Error generating shift Excel:", e)
                traceback.print_exc()

        # Normal HTML table
        try:
            with get_conn() as conn:
                cur = conn.cursor()
                cur.execute(query, params)
                columns = [c[0] for c in cur.description]
                rows = [dict(zip(columns, r)) for r in cur.fetchall()]
        except Exception as e:
            error_message = "Could not load shift data."
            print("Error loading shift data:", e)
            traceback.print_exc()

    return render_template(
        "shift_report.html",
        shifts=shifts,
        from_date=from_date,
        to_date=to_date,
        selected_shift=selected_shift,
        rows=rows,
        columns=columns,
        searched=searched,
        error_message=error_message,
    )


# -------------------- OPERATOR REPORT --------------------
@app.route("/operator-report", methods=["GET", "POST"])
def operator_report():
    operators = []
    rows = []
    columns = []
    searched = False
    error_message = None

    # Try to load operator names.
    # On Render this will fail (no DB), but we catch it so page still loads.
    try:
        with get_conn() as conn:
            df_op = pd.read_sql(
                "SELECT DISTINCT OPERATORNAME FROM dbo.ActualLog "
                "WHERE OPERATORNAME IS NOT NULL AND OPERATORNAME <> '' "
                "ORDER BY OPERATORNAME",
                conn,
            )
            operators = df_op["OPERATORNAME"].tolist()
    except Exception as e:
        error_message = "Database not reachable. Showing empty operator list."
        print("Error loading operator names:", e)
        traceback.print_exc()

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

        try:
            # Excel Export
            if action == "excel":
                with get_conn() as conn:
                    df = pd.read_sql(query, conn, params=params)

                filename = f"OperatorReport_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                filepath = os.path.join(tempfile.gettempdir(), filename)
                df.to_excel(filepath, index=False)
                return send_file(filepath, as_attachment=True, download_name=filename)

            # HTML Table Display
            with get_conn() as conn:
                cur = conn.cursor()
                cur.execute(query, params)
                columns = [c[0] for c in cur.description]
                rows = [dict(zip(columns, r)) for r in cur.fetchall()]
        except Exception as e:
            error_message = "Could not load operator report data."
            print("Error in operator-report query:", e)
            traceback.print_exc()

    return render_template(
        "operator_report.html",
        operators=operators,
        from_date=from_date,
        to_date=to_date,
        selected_operator=selected_operator,
        rows=rows,
        columns=columns,
        searched=searched,
        error_message=error_message,
    )


# -------------------- PRODUCT REPORT --------------------
@app.route("/product-report", methods=["GET", "POST"])
def product_report():
    products = []
    rows = []
    columns = []
    searched = False
    error_message = None

    # Try to load product/recipe names
    try:
        with get_conn() as conn:
            df_prod = pd.read_sql(
                "SELECT DISTINCT RECEIPENAME FROM dbo.ActualLog "
                "WHERE RECEIPENAME IS NOT NULL AND RECEIPENAME <> '' "
                "ORDER BY RECEIPENAME",
                conn,
            )
            products = df_prod["RECEIPENAME"].tolist()
    except Exception as e:
        error_message = "Database not reachable. Showing empty product list."
        print("Error loading product names:", e)
        traceback.print_exc()

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

        if from_date:
            filters.append(" AND DATE1 >= ?")
            params.append(from_date)

        if to_date:
            filters.append(" AND DATE1 <= ?")
            params.append(to_date)

        if selected_product:
            filters.append(" AND RECEIPENAME = ?")
            params.append(selected_product)

        query = base_query + "".join(filters) + " ORDER BY DATE1, TIME1"

        try:
            # Excel export
            if action == "excel":
                with get_conn() as conn:
                    df = pd.read_sql(query, conn, params=params)

                filename = f"ProductReport_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                filepath = os.path.join(tempfile.gettempdir(), filename)
                df.to_excel(filepath, index=False)
                return send_file(filepath, as_attachment=True, download_name=filename)

            # HTML table
            with get_conn() as conn:
                cur = conn.cursor()
                cur.execute(query, params)
                columns = [c[0] for c in cur.description]
                rows = [dict(zip(columns, r)) for r in cur.fetchall()]
        except Exception as e:
            error_message = "Could not load product report data."
            print("Error in product-report query:", e)
            traceback.print_exc()

    return render_template(
        "product_report.html",
        products=products,
        from_date=from_date,
        to_date=to_date,
        selected_product=selected_product,
        rows=rows,
        columns=columns,
        searched=searched,
        error_message=error_message,
    )


# -------------------- START SERVER --------------------
if __name__ == "__main__":
    print("Starting Flask server...")
    # No debug=True in production
    app.run()
