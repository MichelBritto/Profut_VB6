IF(SELECT COUNT(*) FROM SYSOBJECTS WHERE xtype = 'P' AND name = 'USP_SELECIONARESTADOS') >0
BEGIN
	DROP PROCEDURE USP_SELECIONARESTADOS
END

GO

CREATE PROCEDURE USP_SELECIONARESTADOS ( @Codigo_IN INT = NULL )
AS 
    SELECT  UF_CH ,
            ID_IN
    FROM    dbo.ESTADOS_T
    WHERE   ID_IN = ISNULL(@Codigo_IN, ID_IN)
    
GO

GRANT EXECUTE ON USP_SELECIONARESTADOS TO PUBLIC