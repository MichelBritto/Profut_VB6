CREATE PROCEDURE dbo.usp_AdicionarEquipe
    (
      @NOME_VC VARCHAR(128) ,
      @SIGLA_VC VARCHAR(128) ,
      @RESPONSAVEL_VC VARCHAR(128) ,
      @CONTATO_VC1 VARCHAR(128) ,
      @WHATSAPP1_BT BIT ,
      @CONTATO2_VC VARCHAR(128) ,
      @WHATSAP2_BT BIT ,
      @EMAILCONTATO_VC VARCHAR(128) ,
      @ENDERECOIMAGEM_VC VARBINARY(MAX) = NULL,
      @CODIGO_IN INT OUTPUT
    )
AS 
    INSERT  INTO dbo.EQUIPE_T
            ( NOME_VC ,
              SIGLA_VC ,
              RESPONSAVEL_VC ,
              CONTATO1_VC ,
              WHATSAPP1_BT ,
              CONTATO2_VC ,
              WHATSAP2_BT ,
              EMAILCONTATO_VC ,
              USUARIOCADASTRO_VC ,
              USUARIOULTIMAALTERACAO_VC ,
              DATACADASTRO_DT ,
              DATAULTIMAALTERACAO_DT ,
              EXCLUIDO_BT ,
              ENDERECOIMAGEM_VC
	        
            )
    VALUES  ( @NOME_VC ,
              @SIGLA_VC ,
              @RESPONSAVEL_VC ,
              @CONTATO_VC1 ,
              @WHATSAPP1_BT ,
              @CONTATO2_VC ,
              @WHATSAP2_BT ,
              @EMAILCONTATO_VC ,
              CURRENT_USER , -- USUARIOCADASTRO_VC - varchar(128)
              CURRENT_USER , -- USUARIOULTIMAALTERACAO_VC - varchar(128)
              GETDATE() , -- DATACADASTRO_DT - datetime
              GETDATE() , -- DATAULTIMAALTERACAO_DT - datetime
              0 ,  -- EXCLUIDO_BT - bit
              @ENDERECOIMAGEM_VC
            )

    SET @CODIGO_IN = SCOPE_IDENTITY()


GO


