ALTER PROCEDURE usp_SelecionarCargos
    (
      @Cargo_IN INT = NULL ,
      @Ativo_BT BIT = NULL
    )
AS 
    SELECT  ID_IN AS Cargo_IN ,
            Descricao_VC AS Cargo_VC ,
            Ativo_BT
    FROM    dbo.Cargo_T
    WHERE   ID_IN = ISNULL(@Cargo_IN, ID_IN)
            AND ISNULL(Ativo_BT, 0) = ( CASE WHEN ISNULL(@Ativo_BT, 0) = 0 THEN ISNULL(Ativo_BT, 0)
                                             ELSE 1
                                        END )
    
GO

GRANT EXECUTE ON usp_SelecionarCargos TO PUBLIC