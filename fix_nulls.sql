-- PASO 1: Buscar NULLs en TODAS las tablas (una tabla a la vez, eficiente)
DECLARE @tname NVARCHAR(200), @colname NVARCHAR(200), @sql NVARCHAR(MAX), @dtype NVARCHAR(50)

DECLARE cur CURSOR FOR
SELECT TABLE_SCHEMA + '.' + TABLE_NAME, COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE IS_NULLABLE = 'YES' AND TABLE_CATALOG = DB_NAME() AND TABLE_SCHEMA = 'dbo'

OPEN cur
FETCH NEXT FROM cur INTO @tname, @colname, @dtype

WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql = 'IF EXISTS(SELECT 1 FROM ' + @tname + ' WHERE [' + @colname + '] IS NULL) '
    SET @sql = @sql + 'PRINT ''' + @tname + '.' + @colname + ' TIENE NULLS'''
    EXEC sp_executesql @sql
    FETCH NEXT FROM cur INTO @tname, @colname, @dtype
END

CLOSE cur
DEALLOCATE cur
GO

-- PASO 2: Después de revisar el paso 1, ejecuta esto para corregir
-- (descomenta EXEC al final si quieres aplicar automaticamente)
DECLARE @tname2 NVARCHAR(200), @colname2 NVARCHAR(200), @sql2 NVARCHAR(MAX), @dtype2 NVARCHAR(50)

DECLARE cur2 CURSOR FOR
SELECT TABLE_SCHEMA + '.' + TABLE_NAME, COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE IS_NULLABLE = 'YES' AND TABLE_CATALOG = DB_NAME() AND TABLE_SCHEMA = 'dbo'

OPEN cur2
FETCH NEXT FROM cur2 INTO @tname2, @colname2, @dtype2

WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql2 = 'UPDATE ' + @tname2 + ' SET [' + @colname2 + '] = ' +
        CASE 
            WHEN @dtype2 IN ('int','numeric','decimal','float','real','smallint','tinyint','bigint','money') THEN '0'
            WHEN @dtype2 IN ('char','varchar','nchar','nvarchar') THEN ''''''
            WHEN @dtype2 IN ('datetime','smalldatetime') THEN '''1900-01-01'''
            WHEN @dtype2 = 'bit' THEN '0'
            ELSE '0'
        END +
        ' WHERE [' + @colname2 + '] IS NULL;'
    PRINT @sql2
    -- EXEC sp_executesql @sql2  -- Descomenta para ejecutar
    FETCH NEXT FROM cur2 INTO @tname2, @colname2, @dtype2
END

CLOSE cur2
DEALLOCATE cur2
GO
