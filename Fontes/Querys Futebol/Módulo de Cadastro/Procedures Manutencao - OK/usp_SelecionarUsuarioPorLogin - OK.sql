ALTER PROCEDURE dbo.usp_SelecionarUsuarioPorLogin
    (
      @LoginUsuario_VC VARCHAR(1024) ,
      @CodigoUsuario_IN INT OUTPUT
    )
AS 
    SET @CodigoUsuario_IN = ( SELECT TOP 1
                                        ID_IN
                              FROM      dbo.Usuario_T
                              WHERE     Login_VC = @LoginUsuario_VC
                            )
    
GO
