CREATE PROCEDURE dbo.USP_APAGARJOGADORPORCODIGO
    (
      @Jogador_IN INT ,
      @Operacao_IN INT 
    )
AS --OPERAÇÃO
	--1 INATIVAR
	--2 EXCLUIR
	--3 REATIVAR

    IF @Operacao_IN = 1 
        BEGIN --inativar
            UPDATE  dbo.JOGADOR_T
            SET     EXCLUIDO_BT = 1
            WHERE   ID_JOGADOR_IN = @Jogador_IN	
        END
    ELSE 
        BEGIN -- excluir
            IF @Operacao_IN = 2 
                BEGIN
                    DELETE  dbo.JOGADOR_T
                    WHERE   ID_JOGADOR_IN = @Jogador_IN
                END
            ELSE 
                BEGIN -- reativar
                    UPDATE  dbo.JOGADOR_T
                    SET     EXCLUIDO_BT = 0
                    WHERE   ID_JOGADOR_IN = @Jogador_IN	
                END

        END
GO
