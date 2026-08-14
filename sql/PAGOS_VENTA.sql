/* =====================================================================
   PAGOS_VENTA : Medios de pago y vuelto por venta (LK_CODTRA = 2401)
   OSBusiness FE
   Ejecutar una sola vez en BDATOS.
   La aplicacion inserta 1 fila por cada venta grabada.
   ===================================================================== */
IF OBJECT_ID('dbo.PAGOS_VENTA', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.PAGOS_VENTA (
        PAG_CODCIA    CHAR(2)       NOT NULL,
        PAG_NUMOPER   NUMERIC(12,0) NOT NULL,   -- mismo numero que ALLOG.ALL_NUMOPER
        PAG_FECHA     DATETIME      NOT NULL,
        PAG_MEDIO     CHAR(1)       NOT NULL,   -- E=efectivo, Q=QR/Yape/Plin, T=tarjeta, M=mixto
        PAG_EFECTIVO  MONEY         NULL DEFAULT 0,
        PAG_QR        MONEY         NULL DEFAULT 0,
        PAG_TARJETA   MONEY         NULL DEFAULT 0,
        PAG_TOTAL     MONEY         NULL DEFAULT 0,
        PAG_VUELTO    MONEY         NULL DEFAULT 0,
        PAG_USUARIO   VARCHAR(10)   NULL,
        PAG_FBG       CHAR(1)       NULL,
        PAG_NUMSER    INT           NULL,
        PAG_NUMFAC    INT           NULL,
        CONSTRAINT PK_PAGOS_VENTA PRIMARY KEY (PAG_CODCIA, PAG_NUMOPER)
    )
END
