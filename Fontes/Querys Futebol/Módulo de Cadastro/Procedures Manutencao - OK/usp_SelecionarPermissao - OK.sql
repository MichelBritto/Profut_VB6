ALTER PROCEDURE dbo.usp_SelecionarPermissao
    (
      @Permissao_IN INT = NULL
    )
AS 
    SELECT  DISTINCT
            PER.ID_IN ,
            PER.Permissao_IN ,
            PER.Descricao_VC ,
            PER.Status_BT
    FROM    dbo.Permissao_T (NOLOCK) PER
    WHERE   PER.Permissao_IN = ISNULL(@Permissao_IN, PER.Permissao_IN)
   

GO
