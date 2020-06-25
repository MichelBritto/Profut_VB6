ALTER PROCEDURE usp_SelecionarCidades
    (
      @Cidade_IN INT = NULL ,
      @UF_IN INT = NULL
    )
AS 
    SELECT  ID_IN,
			Cidade_IN,
			Nome_VC,
			Estado_IN
    FROM    dbo.cidade_T
    --WHERE   Cidade_IN = ISNULL(@Cidade_IN,Cidade_IN)
    --        AND Estado_IN = ISNULL(@Cidade_IN, @UF_IN)
            
            
GO

GRANT EXECUTE ON usp_SelecionarCidades TO PUBLIC