-- =============================================
-- Author:     OSBusiness
-- Description: Consulta stock de un articulo en locales anexos
-- Uso desde VB6: EXEC sp_stock_local ART_KEY, 'BDATOSPM'
-- =============================================
IF EXISTS (SELECT 1 FROM sysobjects WHERE name = 'sp_stock_local' AND type = 'P')
    DROP PROCEDURE dbo.sp_stock_local;
GO

CREATE PROCEDURE dbo.sp_stock_local
    @art_key INT,
    @local VARCHAR(20)
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @sql NVARCHAR(500);

    IF @local = 'BDATOSPM'
        SET @sql = 'SELECT ARM_CODART, ARM_STOCK FROM BDATOSPM.dbo.ARTICULO WHERE ARM_CODART = ' + CAST(@art_key AS VARCHAR);
    ELSE IF @local = 'SERVER02'
        SET @sql = 'SELECT ARM_CODART, ARM_STOCK FROM SERVER02.BDATOS.dbo.ARTICULO WHERE ARM_CODART = ' + CAST(@art_key AS VARCHAR);
    ELSE
    BEGIN
        SELECT -1 AS ARM_CODART, 0 AS ARM_STOCK;
        RETURN;
    END

    EXEC sp_executesql @sql;
END
GO

-- Ejemplo de uso:
-- EXEC sp_stock_local 12345, 'BDATOSPM'
-- EXEC sp_stock_local 12345, 'SERVER02'
