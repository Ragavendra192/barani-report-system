import pyodbc

SERVER = r"RAGAVENDRA\SQLEXPRESS"
DATABASE = "IIT300"
TABLE_NAME = "ActualLog"   # change if your shift table is different

conn = pyodbc.connect(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    f"SERVER={SERVER};"
    f"DATABASE={DATABASE};"
    "Trusted_Connection=yes;"
)

cur = conn.cursor()
cur.execute(f"SELECT TOP 1 * FROM dbo.{TABLE_NAME}")
print("Columns in table", TABLE_NAME, ":")
for col in cur.description:
    print(col[0])

conn.close()
