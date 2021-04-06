CREATE PROCEDURE dbo.usp_AdicionarAlterarPermissaoPorUsuario
    (
      @Usuario_IN INT ,
      @Permissao_IN INT ,
      @Status_BT BIT
    )
AS --PRIMEIRO VERIFICO SE ESSE USUÁRIO TEM UM REGISTRO NA TABELA UsuarioXPermissao_T 
    IF ( SELECT COUNT(*)
         FROM   dbo.UsuarioXPermissao_T
         WHERE  Permissao_IN = @Permissao_IN
                AND Usuario_IN = @Usuario_IN
       ) > 0 
        BEGIN
	--SE EXISTE, APENAS ALTERO O STATUS
            UPDATE  dbo.UsuarioXPermissao_T
            SET     Status_BT = @Status_BT
            WHERE   Permissao_IN = @Permissao_IN
                    AND Usuario_IN = @Usuario_IN
        END
    ELSE 
        BEGIN
	--SE NÃO EXISTE, INSIRO UM REGISTRO NA TABELA CORRESPONDENTE AO USUARIO E A PERMISSÃO COM O STATUS CORRESPONDENTE
            INSERT  INTO dbo.UsuarioXPermissao_T
                    ( Usuario_IN ,
                      Permissao_IN ,
                      Status_BT
                    )
            VALUES  ( @Usuario_IN , -- Usuario_IN - int
                      @Permissao_IN , -- Permissao_IN - int
                      @Status_BT  -- Status_BT - bit
                    )
        END

GO
