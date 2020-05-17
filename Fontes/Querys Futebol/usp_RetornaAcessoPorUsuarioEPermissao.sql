CREATE PROCEDURE usp_RetornaAcessoPorUsuarioEPermissao
    (
      @Permissao_IN INT ,
      @Usuario_IN INT ,
      @Acesso_BT BIT OUTPUT
    )
AS 
    SET @Acesso_BT = 0
	
    IF ( SELECT COUNT(*)
         FROM   dbo.UsuarioXPermissao_T
         WHERE  Permissao_IN = @Permissao_IN
                AND Usuario_IN = @Usuario_IN
       ) > 0 
        BEGIN
            SET @Acesso_BT = 1
        END
        
GO

GRANT EXECUTE ON usp_RetornaAcessoPorUsuarioEPermissao TO PUBLIC


