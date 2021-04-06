CREATE PROCEDURE dbo.usp_RetornaAcessoPorUsuarioEPermissao
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
                AND Status_BT = 1
       ) > 0 
        BEGIN
            SET @Acesso_BT = 1
        END
        
GO
