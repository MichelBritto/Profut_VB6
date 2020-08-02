ALTER PROCEDURE dbo.USP_SELECIONARCARTEGORIA
    (
      @Cartegoria_in INT = NULL
    )
AS 
    SELECT  DESCRICAO_VC ,
            ID_IN
    FROM    dbo.CARTEGORIA_T
    WHERE   ID_IN = ISNULL(@Cartegoria_in, ID_IN)
    
GO
