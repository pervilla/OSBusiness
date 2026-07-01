import pyodbc
import sys

conn_str = (
    "DRIVER={SQL Server};"
    "SERVER=.;"  # Cambia a tu servidor
    "DATABASE=OSBusiness;"  # Cambia a tu BD
    "UID=sa;"
    "PWD=tu_password;"
)

conn = pyodbc.connect(conn_str)
cur = conn.cursor()

# Obtener todas las columnas nullable de todas las tablas (excluye sys)
cur.execute("""
    SELECT TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME
    FROM INFORMATION_SCHEMA.COLUMNS
    WHERE IS_NULLABLE = 'YES'
      AND TABLE_CATALOG = DB_NAME()
      AND TABLE_SCHEMA = 'dbo'
    ORDER BY TABLE_NAME, COLUMN_NAME
""")

cols = cur.fetchall()
if not cols:
    print("No se encontraron columnas nullable.")
    sys.exit(0)

# Construir una sola consulta por tabla (más eficiente que una por columna)
from collections import defaultdict
por_tabla = defaultdict(list)
for schema, table, col in cols:
    por_tabla[f"{schema}.{table}"].append(col)

print("Buscando NULLs en la base de datos...\n")
found_any = False

for tabla, columnas in por_tabla.items():
    casos = []
    for col in columnas:
        casos.append(f"COUNT(CASE WHEN [{col}] IS NULL THEN 1 END) AS [{col}_nulls]")
    sql = f"SELECT {', '.join(casos)} FROM {tabla}"
    try:
        cur.execute(sql)
        row = cur.fetchone()
        for i, col in enumerate(columnas):
            nulos = row[i]
            if nulos > 0:
                print(f"  {tabla}.{col} -> {nulos} NULLs")
                found_any = True
    except Exception as e:
        print(f"  Error en {tabla}: {e}", file=sys.stderr)

if not found_any:
    print("No hay NULLs en la base de datos.")
else:
    print("\nPara generar los UPDATE correctivos, ejecuta este script con --fix")
