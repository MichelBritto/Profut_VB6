IF ( SELECT COUNT(*)
     FROM   sysobjects
     WHERE  xtype = 'P'
            AND name = 'USP_ADICIONARJOGADOR'
   ) > 0 
    BEGIN
        DROP PROCEDURE USP_ADICIONARJOGADOR
    END

GO

CREATE PROCEDURE USP_ADICIONARJOGADOR
    (
      @APELIDO_VC VARCHAR(128) ,
      @CARTEGORIA_IN INT ,
      @EQUIPE_IN INT ,
      @NOMEATLETA_VC VARCHAR(128) ,
      @DATANASCIMENTO_DT DATETIME ,
      @LOCALNASCIMENTO_VC VARCHAR(1024) ,
      @CERTIDAONASCIMENTO_VC VARCHAR(128) ,
      @CARTORIO_VC VARCHAR(128) ,
      @IDENTIDADE_VC VARCHAR(128) ,
      @ORGAOIDENTIDADE_VC VARCHAR(128) ,
      @NOMEPAI_VC VARCHAR(128) ,
      @STRNOMEMAE_VC VARCHAR(128) ,
      @ESTADO_IN INT ,
      @CIDADE_VC VARCHAR(128) ,
      @BAIRRO_VC VARCHAR(128) ,
      @ENDERECO_VC VARCHAR(128) ,
      @TELCEL1_VC VARCHAR(128) ,
      @TELCEL2_VC VARCHAR(128) ,
      @WPP1_BT BIT ,
      @WPP2_BT BIT ,
      @EMAIL_VC VARCHAR(128) ,
      @FACEBOOK_VC VARCHAR(128) ,
      @ESCOLA_VC VARCHAR(128) ,
      @ESTADOESCOLA_IN INT ,
      @CIDADEESCOLA_VC VARCHAR(128) ,
      @BAIRROESCOLA_VC VARCHAR(128) ,
      @ENDERECOESCOLA_VC VARCHAR(128) ,
      @REDESOCIALESCOLA_VC VARCHAR(128) ,
      @INSTAGRAM_VC VARCHAR(128) ,
      @ENDERECOIMAGEM_VC VARCHAR(1024) ,
      @SEXO_IN INT = 1 ,
      @NUMEROCAMISA_IN INT ,
      @CodigoOutput INT OUTPUT
    )
AS 
    INSERT  INTO dbo.JOGADOR_T
            ( APELIDO_VC ,
              CARTEGORIA_IN ,
              EQUIPE_IN ,
              NOMEATLETA_VC ,
              DATANASCIMENTO_DT ,
              LOCALNASCIMENTO_VC ,
              CERTIDAONASCIMENTO_VC ,
              CARTORIO_VC ,
              IDENTIDADE_VC ,
              ORGAOIDENTIDADE_VC ,
              NOMEPAI_VC ,
              STRNOMEMAE_VC ,
              ESTADO_IN ,
              CIDADE_VC ,
              BAIRRO_VC ,
              ENDERECO_VC ,
              TELCEL1_VC ,
              TELCEL2_VC ,
              WPP1_BT ,
              WPP2_BT ,
              EMAIL_VC ,
              FACEBOOK_VC ,
              ESCOLA_VC ,
              ESTADOESCOLA_IN ,
              CIDADEESCOLA_VC ,
              BAIRROESCOLA_VC ,
              ENDERECOESCOLA_VC ,
              REDESOCIALESCOLA_VC ,
              INSTAGRAM_VC ,
              ENDERECOIMAGEM_VC ,
              EXCLUIDO_BT ,
              USUARIOCADASTRO_VC ,
              DATACADASTRO_DT ,
              USUARIOULTIMAALTERACAO_VC ,
              DATAULTIMAALTERACAO_DT ,
              SEXO_IN ,
              NUMEROCAMISA_IN
              	        
            )
    VALUES  ( @APELIDO_VC ,
              @CARTEGORIA_IN ,
              @EQUIPE_IN ,
              @NOMEATLETA_VC ,
              @DATANASCIMENTO_DT ,
              @LOCALNASCIMENTO_VC ,
              @CERTIDAONASCIMENTO_VC ,
              @CARTORIO_VC ,
              @IDENTIDADE_VC ,
              @ORGAOIDENTIDADE_VC ,
              @NOMEPAI_VC ,
              @STRNOMEMAE_VC ,
              @ESTADO_IN ,
              @CIDADE_VC ,
              @BAIRRO_VC ,
              @ENDERECO_VC ,
              @TELCEL1_VC ,
              @TELCEL2_VC ,
              @WPP1_BT ,
              @WPP2_BT ,
              @EMAIL_VC ,
              @FACEBOOK_VC ,
              @ESCOLA_VC ,
              @ESTADOESCOLA_IN ,
              @CIDADEESCOLA_VC ,
              @BAIRROESCOLA_VC ,
              @ENDERECOESCOLA_VC ,
              @REDESOCIALESCOLA_VC ,
              @INSTAGRAM_VC ,
              @ENDERECOIMAGEM_VC ,
              0 ,
              SYSTEM_USER ,
              GETDATE() ,
              SYSTEM_USER ,
              GETDATE() ,
              @SEXO_IN ,
              @NUMEROCAMISA_IN
	        
            )
	        
    SET @CodigoOutput = SCOPE_IDENTITY()
	        
GO

GRANT EXECUTE ON USP_ADICIONARJOGADOR TO PUBLIC

GO

IF ( SELECT COUNT(*)
     FROM   sysobjects
     WHERE  xtype = 'P'
            AND name = 'USP_ALTERARJOGADOR'
   ) > 0 
    BEGIN
        DROP PROCEDURE USP_ALTERARJOGADOR
    END

GO

CREATE PROCEDURE USP_ALTERARJOGADOR
    (
      @ID_IN INT ,
      @APELIDO_VC VARCHAR(128) ,
      @CARTEGORIA_IN INT ,
      @EQUIPE_IN INT ,
      @NOMEATLETA_VC VARCHAR(128) ,
      @DATANASCIMENTO_DT DATETIME ,
      @LOCALNASCIMENTO_VC VARCHAR(1024) ,
      @CERTIDAONASCIMENTO_VC VARCHAR(128) ,
      @CARTORIO_VC VARCHAR(128) ,
      @IDENTIDADE_VC VARCHAR(128) ,
      @ORGAOIDENTIDADE_VC VARCHAR(128) ,
      @NOMEPAI_VC VARCHAR(128) ,
      @STRNOMEMAE_VC VARCHAR(128) ,
      @ESTADO_IN INT ,
      @CIDADE_VC VARCHAR(128) ,
      @BAIRRO_VC VARCHAR(128) ,
      @ENDERECO_VC VARCHAR(128) ,
      @TELCEL1_VC VARCHAR(128) ,
      @TELCEL2_VC VARCHAR(128) ,
      @WPP1_BT BIT ,
      @WPP2_BT BIT ,
      @EMAIL_VC VARCHAR(128) ,
      @FACEBOOK_VC VARCHAR(128) ,
      @ESCOLA_VC VARCHAR(128) ,
      @ESTADOESCOLA_IN INT ,
      @CIDADEESCOLA_VC VARCHAR(128) ,
      @BAIRROESCOLA_VC VARCHAR(128) ,
      @ENDERECOESCOLA_VC VARCHAR(128) ,
      @REDESOCIALESCOLA_VC VARCHAR(128) ,
      @INSTAGRAM_VC VARCHAR(128) ,
      @ENDERECOIMAGEM_VC VARCHAR(1024) ,
      @SEXO_IN INT = 1 ,
      @NUMEROCAMISA_IN INT ,
      @CodigoOutput INT OUTPUT
    )
AS 
    UPDATE  dbo.JOGADOR_T
    SET     APELIDO_VC = @APELIDO_VC ,
            CARTEGORIA_IN = @CARTEGORIA_IN ,
            EQUIPE_IN = @EQUIPE_IN ,
            NOMEATLETA_VC = @NOMEATLETA_VC ,
            DATANASCIMENTO_DT = @DATANASCIMENTO_DT ,
            LOCALNASCIMENTO_VC = @LOCALNASCIMENTO_VC ,
            CERTIDAONASCIMENTO_VC = @CERTIDAONASCIMENTO_VC ,
            CARTORIO_VC = @CARTORIO_VC ,
            IDENTIDADE_VC = @IDENTIDADE_VC ,
            ORGAOIDENTIDADE_VC = @ORGAOIDENTIDADE_VC ,
            NOMEPAI_VC = @NOMEPAI_VC ,
            STRNOMEMAE_VC = @STRNOMEMAE_VC ,
            ESTADO_IN = @ESTADO_IN ,
            CIDADE_VC = @CIDADE_VC ,
            BAIRRO_VC = @BAIRRO_VC ,
            ENDERECO_VC = @ENDERECO_VC ,
            TELCEL1_VC = @TELCEL1_VC ,
            TELCEL2_VC = @TELCEL2_VC ,
            WPP1_BT = @WPP1_BT ,
            WPP2_BT = @WPP2_BT ,
            EMAIL_VC = @EMAIL_VC ,
            FACEBOOK_VC = @FACEBOOK_VC ,
            ESTADOESCOLA_IN = @ESTADOESCOLA_IN ,
            ESCOLA_VC = @ESCOLA_VC ,
            CIDADEESCOLA_VC = @CIDADEESCOLA_VC ,
            BAIRROESCOLA_VC = @BAIRROESCOLA_VC ,
            ENDERECOESCOLA_VC = @ENDERECOESCOLA_VC ,
            REDESOCIALESCOLA_VC = @REDESOCIALESCOLA_VC ,
            INSTAGRAM_VC = @INSTAGRAM_VC ,
            ENDERECOIMAGEM_VC = @ENDERECOIMAGEM_VC ,
            DATAULTIMAALTERACAO_DT = GETDATE() ,
            USUARIOULTIMAALTERACAO_VC = SYSTEM_USER ,
            SEXO_IN = @SEXO_IN ,
            NUMEROCAMISA_IN = @NUMEROCAMISA_IN
    WHERE   ID_JOGADOR_IN = @ID_IN
	
    SET @CodigoOutput = @ID_IN

GO

GRANT EXECUTE ON USP_ALTERARJOGADOR TO PUBLIC

GO
IF ( SELECT COUNT(*)
     FROM   sysobjects
     WHERE  xtype = 'P'
            AND name = 'USP_SELECIONARJOGADORPORCODIGO'
   ) > 0 
    BEGIN
        DROP PROCEDURE USP_SELECIONARJOGADORPORCODIGO
    END

GO

CREATE PROCEDURE USP_SELECIONARJOGADORPORCODIGO
    (
      @Jogador INT = NULL ,
      @NãoExcluidos BIT = 0
    )
AS 
    IF @Jogador IS NULL 
        BEGIN
            SELECT  *
            FROM    dbo.JOGADOR_T
            WHERE   EXCLUIDO_BT = @NãoExcluidos
        END
    ELSE 
        BEGIN
            SELECT  *
            FROM    dbo.JOGADOR_T
            WHERE   ID_JOGADOR_IN = @Jogador
                    AND EXCLUIDO_BT = @NãoExcluidos
        END

GO

GRANT EXECUTE ON USP_SELECIONARJOGADORPORCODIGO TO PUBLIC

GO

IF ( SELECT COUNT(*)
     FROM   sysobjects
     WHERE  xtype = 'P'
            AND name = 'USP_APAGARJOGADORPORCODIGO'
   ) > 0 
    BEGIN
        DROP PROCEDURE USP_APAGARJOGADORPORCODIGO
    END

GO

CREATE PROCEDURE USP_APAGARJOGADORPORCODIGO ( @Jogador INT )
AS 
    UPDATE  dbo.JOGADOR_T
    SET     EXCLUIDO_BT = 1
    WHERE   ID_JOGADOR_IN = @Jogador
	
GO

GRANT EXECUTE ON dbo.USP_APAGARJOGADORPORCODIGO TO PUBLIC

