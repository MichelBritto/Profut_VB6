CREATE TABLE JOGADOR_T
(
	ID_JOGADOR_IN INT IDENTITY(1,1),
	APELIDO_VC VARCHAR(128),
	CARTEGORIA_IN INT,
	EQUIPE_IN INT,
	NOMEATLETA_VC VARCHAR(128),
	DATANASCIMENTO_DT DATETIME,
	LOCALNASCIMENTO_VC VARCHAR(1024),
	CERTIDAONASCIMENTO_VC VARCHAR(128),
	CARTORIO_VC VARCHAR(128),
	IDENTIDADE_VC VARCHAR(128),
	ORGAOIDENTIDADE_VC VARCHAR(128),
	NOMEPAI_VC VARCHAR(128),
	STRNOMEMAE_VC VARCHAR(128),
	ESTADO_IN INT,
	CIDADE_VC VARCHAR(128),
	BAIRRO_VC VARCHAR(128),
	ENDERECO_VC VARCHAR(128),
	TELCEL1_VC VARCHAR(128),
	TELCEL2_VC VARCHAR(128),
	WPP1_BT BIT,
	WPP2_BT BIT,
	EMAIL_VC VARCHAR(128),
	FACEBOOK_VC VARCHAR(128),
	ESCOLA_VC VARCHAR(128),
	ESTADOESCOLA_IN INT,
	CIDADEESCOLA_VC VARCHAR(128),
	BAIRROESCOLA_VC VARCHAR(128),
	ENDERECOESCOLA_VC VARCHAR(128),
	REDESOCIALESCOLA_VC VARCHAR(128),
	INSTAGRAM_VC VARCHAR(128),
	ENDERECOIMAGEM_VC VARCHAR(1024),
	EXCLUIDO_BT BIT,
	USUARIOCADASTRO_VC VARCHAR(128),
	DATACADASTRO_DT DATETIME,
	USUARIOULTIMAALTERACAO_VC VARCHAR(128),
	DATAULTIMAALTERACAO_DT DATETIME,
	SEXO_IN INT, --1 MASCULINO 2 FEMININO,
	NUMEROCAMISA_IN INT
	CONSTRAINT pk_jogador_IN PRIMARY KEY(ID_JOGADOR_IN)
)
