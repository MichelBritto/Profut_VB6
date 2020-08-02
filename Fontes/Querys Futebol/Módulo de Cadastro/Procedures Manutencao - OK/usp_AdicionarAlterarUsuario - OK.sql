ALTER PROCEDURE dbo.usp_AdicionarAlterarUsuario
    (
      @login_VC VARCHAR(20) ,
      @nomecompleto_VC VARCHAR(50) ,
      @email_VC VARCHAR(100) = NULL ,
      @telefone_VC VARCHAR(100) = NULL ,
      @usuario_IN INT OUTPUT ,
      @setorinterno_IN INT ,
      @clube_IN INT = NULL
    )
AS 
    DECLARE @comando_VC VARCHAR(500)
    DECLARE @NomeBase_VC VARCHAR(20)

    IF @usuario_IN > 0 
        BEGIN
            UPDATE  usuario_T
            SET     Login_VC = @login_VC ,
                    Nome_VC = @nomecompleto_VC ,
                    Telefone_VC = @telefone_VC ,
                    Cargo_IN = CASE WHEN @setorinterno_IN = 0 THEN 'NULL'
                                    ELSE @setorinterno_IN
                               END ,
                    email_VC = @email_VC ,
                    clube_IN = @clube_IN
            WHERE   ID_IN = @usuario_IN
        END
    ELSE 
        BEGIN
            INSERT  INTO usuario_T
                    ( login_VC ,
                      nome_VC ,
                      cargo_IN ,
                      email_VC ,
                      Telefone_VC ,
                      clube_IN 
                    )
            VALUES  ( @login_VC ,
                      @nomecompleto_VC ,
                      @setorinterno_IN ,
                      @email_VC ,
                      @telefone_VC ,
                      @clube_IN
                    )

            SET @usuario_IN = @@IDENTITY
            SET @NomeBase_VC = DB_NAME() 
		-- Criando usuário no SQL Server

            SET @comando_VC = 'CREATE LOGIN [' + @login_VC + '] WITH PASSWORD=''123'''		
            EXEC (@comando_VC)
            SET @comando_VC = 'CREATE USER [' + @nomecompleto_VC + '] FOR LOGIN [' + @login_VC + ']'
            EXEC (@comando_VC)			
        END


    RETURN @@ERROR
GO
