ALTER PROCEDURE dbo.USP_ADICIONARJOGADOR
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
      @ENDERECOIMAGEM_VC VARBINARY(MAX) = NULL ,
      @SEXO_IN INT = 1 ,
      @NUMEROCAMISA_IN INT ,
      @POSICAO_IN INT ,
      @CodigoOutput INT OUTPUT
    )
AS 
    DECLARE @CodigoUnico_VC VARCHAR(8)
	
    SET @CodigoUnico_VC = CAST(RAND() * 100000000 AS BIGINT)
	
    IF ( SELECT COUNT(*)
         FROM   dbo.JOGADOR_T
         WHERE  CODIGOUNICO_VC = @CodigoUnico_VC
       ) > 0 
        BEGIN
            SET @CodigoUnico_VC = CAST(RAND() * 100000000 AS BIGINT)
        END
	
	
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
              NUMEROCAMISA_IN ,
              POSICAO_IN ,
              CODIGOUNICO_VC
              	        
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
              @NUMEROCAMISA_IN ,
              @POSICAO_IN ,
              @CodigoUnico_VC
	        
            )
	        
    SET @CodigoOutput = SCOPE_IDENTITY()
	        
GO
