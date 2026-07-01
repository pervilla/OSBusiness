import pyodbc, sys
from collections import defaultdict

conn_str = (
    "DRIVER={SQL Server};"
    "SERVER=.;"
    "DATABASE=OSBusiness;"
    "UID=sa;"
    "PWD=tu_password;"
)

conn = pyodbc.connect(conn_str)
cur = conn.cursor()

# Obtener todas las columnas nullable de todas las tablas
cur.execute("""
    SELECT TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME
    FROM INFORMATION_SCHEMA.COLUMNS
    WHERE IS_NULLABLE = 'YES'
      AND TABLE_CATALOG = DB_NAME()
      AND TABLE_SCHEMA = 'dbo'
      AND TABLE_NAME NOT LIKE 'sys%'
    ORDER BY TABLE_NAME, COLUMN_NAME
""")

cols = cur.fetchall()
if not cols:
    print("No hay columnas nullable.")
    sys.exit(0)

por_tabla = defaultdict(list)
for schema, table, col in cols:
    por_tabla[f"{schema}.{table}"].append(col)

total_null_cols = 0
for tabla, columnas in por_tabla.items():
    casos = []
    for col in columnas:
        casos.append(f"COUNT(CASE WHEN [{col}] IS NULL THEN 1 END) AS [{col}_nulls]")
    sql = f"SELECT {', '.join(casos)} FROM {tabla} WITH (NOLOCK)"
    try:
        cur.execute(sql)
        row = cur.fetchone()
        for i, col in enumerate(columnas):
            nulos = row[i]
            if nulos > 0:
                print(f"{tabla}.{col} -> {nulos} NULLs")
                total_null_cols += 1
    except Exception as e:
        print(f"ERROR en {tabla}: {e}", file=sys.stderr)

if total_null_cols == 0:
    print("NO HAY NULLs en ninguna tabla.")
else:
    print(f"\nTotal: {total_null_cols} columnas con NULLs.")
    print("\nPara generar los UPDATE correctivos, ejecuta este script con --fix")
