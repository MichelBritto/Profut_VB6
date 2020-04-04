CREATE PROCEDURE dbo.usp_AltetarLoginUsuario
    (
      @SenhaAntiga_VC SYSNAME ,
      @SenhaNova_VC SYSNAME
    )
AS 
    BEGIN

        DECLARE @ERROR_IN INT
	
        EXEC @ERROR_IN = sp_password @SenhaAntiga_VC, @SenhaNova_VC

        IF @ERROR_IN <> 0 
            RETURN @ERROR_IN

        UPDATE  usuario_T
        SET     usu_trocarsenha_BT = 0
        WHERE   usu_login_VC = SYSTEM_USER

        RETURN @@ERROR
    END
