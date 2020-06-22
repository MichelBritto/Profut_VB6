ALTER PROCEDURE usp_SelecionarUsuarios ( @Usuario_IN INT = NULL )
AS 
    SELECT  USU.ID_IN ,
            USU.Login_VC ,
            USU.Nome_VC ,
            USU.Cargo_IN ,
            CAR.Descricao_VC ,
            USU.Telefone_VC ,
            USU.Email_VC ,
            USU.clube_IN ,
            EQU.NOME_VC AS NomeEquipe_VC
    FROM    dbo.Usuario_T USU
            INNER JOIN dbo.Cargo_T (NOLOCK) CAR ON CAR.ID_IN = USU.Cargo_IN
            LEFT JOIN dbo.EQUIPE_T (NOLOCK) EQU ON EQU.ID_IN = USU.clube_IN
    WHERE   USU.ID_IN = ISNULL(@Usuario_IN, USU.ID_IN)
            
GO

GRANT EXECUTE ON usp_SelecionarUsuarios TO PUBLIC