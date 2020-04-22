CREATE PROCEDURE usp_SelecionarCargos ( @Cargo_IN INT = NULL )
AS 
    SELECT  ID_IN ,
            Descricao_VC
    FROM    dbo.Cargo_T
    WHERE   ID_IN = ISNULL(@Cargo_IN, ID_IN)
    
GO

GRANT EXECUTE ON usp_SelecionarCargos TO PUBLIC