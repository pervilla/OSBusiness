-- =============================================
-- Author:     OSBusiness
-- Description: Consulta productos en transito
-- Basado en: PHP get_productos_transito()
-- Uso desde VB6: EXEC sp_productos_transito 'paracetamol 500'
-- =============================================
IF EXISTS (SELECT 1 FROM sysobjects WHERE name = 'sp_productos_transito' AND type = 'P')
    DROP PROCEDURE dbo.sp_productos_transito;
GO

CREATE PROCEDURE dbo.sp_productos_transito
    @busqueda VARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    -- Dividir la búsqueda en palabras
    DECLARE @palabras TABLE (palabra VARCHAR(50));
    DECLARE @pos INT, @palabra VARCHAR(50);
    DECLARE @busqueda_upper VARCHAR(200);
    
    SET @busqueda_upper = UPPER(LTRIM(RTRIM(@busqueda)));
    
    -- Insertar cada palabra
    WHILE LEN(@busqueda_upper) > 0
    BEGIN
        SET @pos = CHARINDEX(' ', @busqueda_upper);
        IF @pos > 0
        BEGIN
            SET @palabra = LTRIM(RTRIM(LEFT(@busqueda_upper, @pos - 1)));
            SET @busqueda_upper = LTRIM(RTRIM(SUBSTRING(@busqueda_upper, @pos + 1, LEN(@busqueda_upper))));
        END
        ELSE
        BEGIN
            SET @palabra = LTRIM(RTRIM(@busqueda_upper));
            SET @busqueda_upper = '';
        END
        
        IF LEN(@palabra) > 0
            INSERT INTO @palabras VALUES (@palabra);
    END

    -- Construir consulta dinámica
    DECLARE @sql NVARCHAR(MAX);
    DECLARE @condiciones NVARCHAR(MAX) = '';
    
    SELECT @condiciones = @condiciones + 
        CASE WHEN @condiciones = '' THEN '' ELSE ' AND ' END +
        'd.DES_PROD LIKE ''%' + palabra + '%'''
    FROM @palabras;
    
    IF @condiciones = ''
    BEGIN
        SELECT TOP 0
            '' AS ART_NOMBRE,
            0 AS CANTIDAD,
            0 AS PRECIO,
            '' AS NRO_FACTURA,
            '' AS FECHA_DOC,
            '' AS PROVEEDOR;
        RETURN;
    END

    SET @sql = N'
        SELECT TOP 50
            RTRIM(d.DES_PROD) AS ART_NOMBRE, 
            d.CANTIDAD, 
            d.PRECIO, 
            RTRIM(f.NRO_FACTURA) AS NRO_FACTURA, 
            CONVERT(VARCHAR, f.FECHA, 103) AS FECHA_DOC,
            RTRIM(c.CLI_NOMBRE) AS PROVEEDOR
        FROM IMPORT_FACT_DET d
        INNER JOIN IMPORT_FACT f ON d.IDFACT = f.ID
        LEFT JOIN CLIENTES c ON f.RUC = c.CLI_RUC_ESPOSO AND c.cli_cp = ''P''
        WHERE f.ESTADO = 0 
        AND ' + @condiciones + '
        ORDER BY f.FECHA ASC';

    EXEC sp_executesql @sql;
END
GO
