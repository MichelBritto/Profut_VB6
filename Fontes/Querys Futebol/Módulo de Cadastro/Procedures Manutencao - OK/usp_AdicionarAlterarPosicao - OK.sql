CREATE PROCEDURE dbo.usp_AdicionarAlterarPosicao
    (
      @Descricao_VC VARCHAR(1024) ,
      @Ativo_BT BIT ,
      @Posicao_IN INT = NULL
    )
AS 
    IF ISNULL(@Posicao_IN, 0) = 0 
        BEGIN
            IF ( SELECT COUNT(*)
                 FROM   dbo.PosicaoAtleta_T
                 WHERE  Descricao_VC = @Descricao_VC
               ) = 0 
                BEGIN
					--DESSA FORMA GARANTO QUE NAO SERÁ ADICIONADO NENHUM DUPLICADO
                    INSERT  INTO dbo.PosicaoAtleta_T
                            ( Descricao_VC, Ativo_BT )
                    VALUES  ( @Descricao_VC, -- Descricao_VC - varchar(128)
                              @Ativo_BT  -- Ativo_BT - bit
                              )
                END
        END
    ELSE 
        BEGIN
            UPDATE  dbo.PosicaoAtleta_T
            SET     Ativo_BT = @Ativo_BT ,
                    Descricao_VC = @Descricao_VC
            WHERE   ID_IN = @Posicao_IN
        END
GO
