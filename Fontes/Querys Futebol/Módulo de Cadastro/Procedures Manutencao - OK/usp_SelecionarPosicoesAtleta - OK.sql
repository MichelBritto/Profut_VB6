CREATE PROCEDURE dbo.usp_SelecionarPosicoesAtleta
    (
      @Posicao_IN INT = NULL ,
      @Ativo_BT BIT = 1
    )
AS 
    SELECT  ID_IN AS Posicao_IN ,
            Descricao_VC AS Posicao_VC ,
            Ativo_BT
    FROM    dbo.PosicaoAtleta_T
    WHERE   ID_IN = ISNULL(@Posicao_IN, ID_IN)
            AND ISNULL(Ativo_BT, 0) = ( CASE WHEN ISNULL(@Ativo_BT, 0) = 0 THEN ISNULL(Ativo_BT, 0)
                                             ELSE 1
                                        END )
    
GO
