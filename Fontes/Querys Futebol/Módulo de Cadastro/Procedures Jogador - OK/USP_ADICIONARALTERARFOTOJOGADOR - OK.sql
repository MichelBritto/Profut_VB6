ALTER PROCEDURE dbo.USP_ADICIONARALTERARFOTOJOGADOR
    (
      @Jogador_IN INT ,
      @Caminhofoto_VC VARCHAR(1024)
    )
AS 
    UPDATE  dbo.JOGADOR_T
    SET     ENDERECOIMAGEM_VC = @Caminhofoto_VC
    WHERE   ID_JOGADOR_IN = @Jogador_IN
	
GO
