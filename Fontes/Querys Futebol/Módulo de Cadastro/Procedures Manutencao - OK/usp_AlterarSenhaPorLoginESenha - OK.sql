ALTER PROCEDURE dbo.usp_AlterarSenhaPorLoginESenha
    (
      @Login_VC VARCHAR(1024) ,
      @SenhaAntiga_VC VARCHAR(1024) ,
      @SenhaNova_VC VARCHAR(1024) ,
      @Resultado_BT BIT OUTPUT
    )
AS 
    DECLARE @Hash_VB AS VARBINARY(MAX)
    DECLARE @Comando_VC VARCHAR(1024)
	
    SET @Hash_VB = ( SELECT password_hash
                     FROM   master.sys.sql_logins
                     WHERE  name = @Login_VC
                   )
                   
    IF ISNULL(@Hash_VB, 0) = 0 
        RETURN 0
    
    IF ( pwdcompare(@SenhaAntiga_VC, @Hash_VB) ) = 1 
        BEGIN
            SET @Comando_VC = 'ALTER LOGIN [' + @login_VC + ']WITH PASSWORD=''@SenhaNova_VC'''
            SET @Comando_VC = REPLACE(@Comando_VC, '@SenhaNova_VC', @SenhaNova_VC)	
            EXEC (@Comando_VC)   
            SET @Resultado_BT = 1
        END
    ELSE 
        BEGIN
            SET @Resultado_BT = 0
        END
        
GO
