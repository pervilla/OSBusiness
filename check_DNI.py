import pyodbc

server = '192.168.101.200'
uid = 'sa'
pwd = '159357852456'
database = 'BDATOS'

conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={uid};PWD={pwd}'
conn = pyodbc.connect(conn_str, timeout=10)
cursor = conn.cursor()

print('=== BUSCANDO CAMPOS DE DNI/RUC EN CLIENTES ===')

# Check CLIENTES table for DNI/RUC fields
cursor.execute("""
    SELECT COLUMN_NAME, DATA_TYPE 
    FROM INFORMATION_SCHEMA.COLUMNS 
    WHERE TABLE_NAME = 'CLIENTES' 
    AND (COLUMN_NAME LIKE '%DNI%' OR COLUMN_NAME LIKE '%RUC%' OR COLUMN_NAME LIKE '%DOC%' OR COLUMN_NAME LIKE '%IDENT%')
    ORDER BY ORDINAL_POSITION
""")
columns = cursor.fetchall()
for col in columns:
    print(f'  {col[0]}: {col[1]}')

# Check LK_DIG_RUC value
print()
print('=== VERIFICANDO VARIABLES GLOBALES ===')
print('(Estas variables están en Módulo1.bas)')

# Check a sample client to see the data
print()
print('=== MUESTRA DE DATOS DE CLIENTE ===')
cursor.execute("SELECT TOP 3 CLI_CODCLIE, CLI_NOMBRE, CLI_RUC_ESPOSO, CLI_CP FROM CLIENTES WHERE CLI_CODCIA = '00'")
clients = cursor.fetchall()
for c in clients:
    print(f'  Codigo: {c[0]}, Nombre: {c[1][:30]}, RUC: {c[2]}, CP: {c[3]}')

conn.close()
