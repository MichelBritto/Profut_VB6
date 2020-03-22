ALTER PROCEDURE dbo.usp_SelecionarDadosParaRelatorioDeJogador
    (
      @Nome_VC VARCHAR(1024) = NULL ,
      @Apelido_VC VARCHAR(1024) = NULL ,
      @Cartegoria_IN INT = NULL ,
      @DataNascimentoDE_DT DATETIME = NULL ,
      @DataNascimentoATE_DT DATETIME = NULL ,
      @UF_IN INT = NULL ,
      @Cidade_VC VARCHAR(1024) = NULL ,
      @Bairro_VC VARCHAR(1024) = NULL ,
      @Equipes_VC VARCHAR(1024) = NULL,
      @Sexo_IN INT = NULL
    )
AS 
	SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED
	SET NOCOUNT ON
	
    IF ISNULL(@Nome_VC, '') <> '' 
        BEGIN
            SET @Nome_VC = '%' + @Nome_VC + '%'
        END
    IF ISNULL(@Apelido_VC, '') <> '' 
        BEGIN
            SET @Apelido_VC = '%' + @Apelido_VC + '%'
        END

    IF ISNULL(@Cidade_VC, '') <> '' 
        BEGIN
            SET @Cidade_VC = '%' + @Cidade_VC + '%'
        END

    IF ISNULL(@Bairro_VC, '') <> '' 
        BEGIN
            SET @Bairro_VC = '%' + @Bairro_VC + '%'
        END
        
    IF ISNULL(@DataNascimentoDE_DT, 0) = 0
        AND ISNULL(@DataNascimentoATE_DT, 0) <> 0 
        BEGIN
            SET @DataNascimentoDE_DT = '1900/1/1'
        END
    IF ISNULL(@DataNascimentoDE_DT, 0) <> 0
        AND ISNULL(@DataNascimentoATE_DT, 0) = 0 
        BEGIN
            SET @DataNascimentoATE_DT = '2199/12/31'
        END
    IF ISNULL(@DataNascimentoDE_DT, 0) = 0
        AND ISNULL(@DataNascimentoATE_DT, 0) = 0 
        BEGIN
			SET @DataNascimentoDE_DT = '1900/1/1'
			SET @DataNascimentoATE_DT = '2199/12/31'
        END
        
    SELECT	jog.ID_JOGADOR_IN AS Codigo,
			jog.SEXO_IN AS Sexo,
			CAT.DESCRICAO_VC AS Cartegoria,
			EQU.NOME_VC AS Equipe,
			est.UF_CH,
			jog.ID_JOGADOR_IN ,
			jog.APELIDO_VC AS Apelido,
			jog.CARTEGORIA_IN	,
			jog.EQUIPE_IN	,
			jog.NOMEATLETA_VC	AS Nome,
			jog.DATANASCIMENTO_DT AS DataNascimento	,
			jog.LOCALNASCIMENTO_VC	,
			jog.CERTIDAONASCIMENTO_VC	,
			jog.CARTORIO_VC	,
			jog.IDENTIDADE_VC	,
			jog.ORGAOIDENTIDADE_VC	,
			jog.NOMEPAI_VC	,
			jog.STRNOMEMAE_VC	,
			jog.ESTADO_IN	,
			jog.CIDADE_VC	,
			jog.BAIRRO_VC	,
			jog.ENDERECO_VC	,
			jog.TELCEL1_VC	,
			jog.TELCEL2_VC	,
			jog.WPP1_BT	,
			jog.WPP2_BT	,
			jog.EMAIL_VC	,
			jog.FACEBOOK_VC	,
			jog.ESCOLA_VC	,
			jog.ESTADOESCOLA_IN	,
			jog.CIDADEESCOLA_VC	,
			jog.BAIRROESCOLA_VC	,
			jog.ENDERECOESCOLA_VC	,
			jog.REDESOCIALESCOLA_VC	,
			jog.INSTAGRAM_VC	,
			jog.ENDERECOIMAGEM_VC	,
			jog.EXCLUIDO_BT	,
			jog.USUARIOCADASTRO_VC	,
			jog.DATACADASTRO_DT	,
			jog.USUARIOULTIMAALTERACAO_VC	,
			jog.DATAULTIMAALTERACAO_DT

			
    FROM    dbo.JOGADOR_T (NOLOCK) JOG
            INNER JOIN dbo.EQUIPE_T (NOLOCK) EQU ON EQU.ID_IN = JOG.EQUIPE_IN
            INNER JOIN dbo.CARTEGORIA_T (NOLOCK) CAT ON CAT.ID_IN = JOG.CARTEGORIA_IN
            INNER JOIN dbo.ESTADOS_T (NOLOCK) EST ON EST.ID_IN = JOG.ESTADO_IN
            
	WHERE	JOG.NOMEATLETA_VC LIKE ISNULL(@Nome_VC,JOG.NOMEATLETA_VC) AND
			JOG.APELIDO_VC LIKE ISNULL(@Apelido_VC,JOG.APELIDO_VC) AND
			JOG.CARTEGORIA_IN = ISNULL(@Cartegoria_IN,JOG.CARTEGORIA_IN) AND
			JOG.DATANASCIMENTO_DT BETWEEN @DataNascimentoDE_DT AND @DataNascimentoATE_DT AND
			JOG.ESTADO_IN = ISNULL(@UF_IN,JOG.ESTADO_IN) AND
			JOG.CIDADE_VC LIKE ISNULL(@Cidade_VC,JOG.CIDADE_VC) AND
			JOG.BAIRRO_VC LIKE ISNULL(@Bairro_VC,JOG.BAIRRO_VC) AND
			JOG.EQUIPE_IN IN (SELECT EQUIPE_IN  FROM dbo.fnc_MontaEquipes(@Equipes_VC)) AND
			JOG.SEXO_IN = (CASE WHEN ISNULL(@Sexo_IN,0) = 0 THEN JOG.SEXO_IN WHEN ISNULL(@Sexo_IN,0) = 1 THEN 1 ELSE 2 END)
			
	
      --@Apelido_VC VARCHAR(1024) = NULL ,
      --@Cartegoria_IN INT = NULL ,
      --@DataNascimentoDE_DT DATETIME = NULL ,
      --@DataNascimentoATE_DT DATETIME = NULL ,
      --@UF_IN INT = NULL ,
      --@Cidade_VC VARCHAR(1024) = NULL ,
      --@Bairro_VC VARCHAR(1024) = NULL ,
      --@Equipes_VC VARCHAR(1024) = NULL
GO
