CREATE TABLE dbo.JOGADOR_T
    (
      ID_JOGADOR_IN INT IDENTITY(1, 1)
                        NOT NULL ,
      APELIDO_VC VARCHAR(128) NULL ,
      CARTEGORIA_IN INT NULL ,
      EQUIPE_IN INT NULL ,
      NOMEATLETA_VC VARCHAR(128) NULL ,
      DATANASCIMENTO_DT DATETIME NULL ,
      LOCALNASCIMENTO_VC VARCHAR(1024) NULL ,
      CERTIDAONASCIMENTO_VC VARCHAR(128) NULL ,
      CARTORIO_VC VARCHAR(128) NULL ,
      IDENTIDADE_VC VARCHAR(128) NULL ,
      ORGAOIDENTIDADE_VC VARCHAR(128) NULL ,
      NOMEPAI_VC VARCHAR(128) NULL ,
      STRNOMEMAE_VC VARCHAR(128) NULL ,
      ESTADO_IN INT NULL ,
      CIDADE_VC VARCHAR(128) NULL ,
      BAIRRO_VC VARCHAR(128) NULL ,
      ENDERECO_VC VARCHAR(128) NULL ,
      TELCEL1_VC VARCHAR(128) NULL ,
      TELCEL2_VC VARCHAR(128) NULL ,
      WPP1_BT BIT NULL ,
      WPP2_BT BIT NULL ,
      EMAIL_VC VARCHAR(128) NULL ,
      FACEBOOK_VC VARCHAR(128) NULL ,
      ESCOLA_VC VARCHAR(128) NULL ,
      ESTADOESCOLA_IN INT NULL ,
      CIDADEESCOLA_VC VARCHAR(128) NULL ,
      BAIRROESCOLA_VC VARCHAR(128) NULL ,
      ENDERECOESCOLA_VC VARCHAR(128) NULL ,
      REDESOCIALESCOLA_VC VARCHAR(128) NULL ,
      INSTAGRAM_VC VARCHAR(128) NULL ,
      ENDERECOIMAGEM_VC VARBINARY(MAX) NULL ,
      EXCLUIDO_BT BIT NULL ,
      USUARIOCADASTRO_VC VARCHAR(128) NULL ,
      DATACADASTRO_DT DATETIME NULL ,
      USUARIOULTIMAALTERACAO_VC VARCHAR(128) NULL ,
      DATAULTIMAALTERACAO_DT DATETIME NULL ,
      SEXO_IN INT NULL ,
      NUMEROCAMISA_IN INT NULL ,
      POSICAO_IN INT,
      CONSTRAINT pk_jogador_IN PRIMARY KEY CLUSTERED ( ID_JOGADOR_IN ASC )
    )
