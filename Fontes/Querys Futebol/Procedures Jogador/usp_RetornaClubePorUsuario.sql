CREATE PROCEDURE usp_RetornaClubePorUsuario
    (
      @CodigoUsuario_IN INT ,
      @Clube_IN INT OUTPUT
    )
AS 
    SET @Clube_IN = ( SELECT    ISNULL(clube_IN, 0)
                      FROM      dbo.Usuario_T
                      WHERE     ID_IN = @CodigoUsuario_IN
                    )
GO

GRANT EXECUTE ON usp_RetornaClubePorUsuario TO PUBLIC