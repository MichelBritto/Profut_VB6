CREATE PROCEDURE dbo.USP_ADICIONARALTERARFOTOJOGADOR
    (
      @Jogador_IN INT ,
      @Caminhofoto_VC VARBINARY(max)
    )
AS 
    UPDATE  dbo.JOGADOR_T
    SET     ENDERECOIMAGEM_VC = @Caminhofoto_VC
    WHERE   ID_JOGADOR_IN = @Jogador_IN
	
GO
