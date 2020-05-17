CREATE PROCEDURE usp_SelecionarPermissaoPorUsuario ( @Usuario_IN INT )
AS 
    SELECT  ISNULL(UXP.Status_BT, 0) AS Status_BT ,
            PER.Permissao_IN ,
            PER.Descricao_VC
    FROM    dbo.Permissao_T PER
            LEFT JOIN dbo.UsuarioXPermissao_T UXP ON PER.Permissao_IN = UXP.Permissao_IN
                                                     AND UXP.Usuario_IN = @Usuario_IN
                                                     
GO

GRANT EXECUTE ON usp_SelecionarPermissaoPorUsuario TO PUBLIC
