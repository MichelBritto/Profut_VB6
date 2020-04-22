CREATE PROCEDURE usp_SelecionarUsuarios ( @Usuario_IN INT = NULL )
AS 
    SELECT  USU.ID_IN ,
            USU.Login_VC ,
            USU.Nome_VC ,
            USU.Cargo_IN ,
            CAR.Descricao_VC ,
            USU.Telefone_VC ,
            USU.Email_VC
    FROM    dbo.Usuario_T USU
            INNER JOIN dbo.Cargo_T (NOLOCK) CAR ON CAR.ID_IN = USU.Cargo_IN
            
GO

GRANT EXECUTE ON usp_SelecionarUsuarios TO PUBLIC