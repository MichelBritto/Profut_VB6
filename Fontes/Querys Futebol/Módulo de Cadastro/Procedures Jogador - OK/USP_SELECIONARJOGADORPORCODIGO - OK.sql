ALTER PROCEDURE dbo.USP_SELECIONARJOGADORPORCODIGO
    (
      @Jogador INT = NULL ,
      @N�oExcluidos BIT = 0
    )
AS 
    IF @Jogador IS NULL 
        BEGIN
            SELECT  *
            FROM    dbo.JOGADOR_T
            --WHERE   EXCLUIDO_BT = @N�oExcluidos
        END
    ELSE 
        BEGIN
            SELECT  *
            FROM    dbo.JOGADOR_T
            WHERE   ID_JOGADOR_IN = @Jogador
                    --AND EXCLUIDO_BT = @N�oExcluidos
        END

GO
