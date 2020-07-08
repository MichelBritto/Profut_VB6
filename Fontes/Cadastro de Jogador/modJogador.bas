Attribute VB_Name = "modJogador"
Option Explicit


Type TypJogador

    lngCodigo As Long
    strApelido As String
    lngCartegoria  As Long
    lngEquipe As Long
    strNomeAtleta  As String
    datDataNascimento  As Date
    datDataNascimentoDE  As Date
    datDataNascimentoATE  As Date
    strLocalNascimento As String
    strCertidaoNascimento  As String
    strCartorio As String
    strIdentidade As String
    strOrgaoIdentidade As String
    strObservacao As String
    strNomePai  As String
    strNomeMae  As String
    lngEstado   As Long
    strCidade As String
    strBairro As String
    strEndereco  As String
    strTelefone1 As String
    strTelefone2 As String
    blnWpp1 As Boolean
    blnWpp2 As Boolean
    strEmailContato As String
    strFacebook As String
    strEscola As String
    lngEstadoEscola As Long
    strCidadeEscola As String
    strBairroEscola As String
    strEnderecoEscola  As String
    strEmailContatoEscola As String
    strFacebookEscola As String
    strInstagram As String
    strEnderecoImagem() As Byte
    strEquipes As String
    lngSexo As Long
    lngNumeroCamisa As Long
    blnTemImagem As Boolean
    lngPosicao As Long

End Type

Public Sub modJogador_AdicionarJogador(ByRef udtJogador As TypJogador)

On Error GoTo Erro

    Set gobjCmd.ActiveConnection = gSMConexao.Conexao
    gobjCmd.CommandText = "dbo.usp_AdicionarJogador"
    gobjCmd.CommandType = adCmdStoredProc
    gobjCmd.CommandTimeout = 1000

    With gobjCmd
        .Parameters("@APELIDO_VC").Value = udtJogador.strApelido
        .Parameters("@CARTEGORIA_IN").Value = udtJogador.lngCartegoria
        .Parameters("@EQUIPE_IN").Value = udtJogador.lngEquipe
        .Parameters("@NOMEATLETA_VC").Value = udtJogador.strNomeAtleta
        .Parameters("@DATANASCIMENTO_DT").Value = udtJogador.datDataNascimento
        .Parameters("@LOCALNASCIMENTO_VC").Value = udtJogador.strLocalNascimento
        .Parameters("@CERTIDAONASCIMENTO_VC").Value = udtJogador.strCertidaoNascimento
        .Parameters("@CARTORIO_VC").Value = udtJogador.strCartorio
        .Parameters("@IDENTIDADE_VC").Value = udtJogador.strIdentidade
        .Parameters("@ORGAOIDENTIDADE_VC").Value = udtJogador.strOrgaoIdentidade
        .Parameters("@NOMEPAI_VC").Value = udtJogador.strNomePai
        .Parameters("@STRNOMEMAE_VC").Value = udtJogador.strNomeMae
        .Parameters("@ESTADO_IN").Value = udtJogador.lngEstado
        .Parameters("@CIDADE_VC").Value = udtJogador.strCidade
        .Parameters("@BAIRRO_VC").Value = udtJogador.strBairro
        .Parameters("@ENDERECO_VC").Value = udtJogador.strEndereco
        .Parameters("@TELCEL1_VC").Value = udtJogador.strTelefone1
        .Parameters("@TELCEL2_VC").Value = udtJogador.strTelefone2
        .Parameters("@WPP1_BT").Value = udtJogador.blnWpp1
        .Parameters("@WPP2_BT").Value = udtJogador.blnWpp2
        .Parameters("@EMAIL_VC").Value = udtJogador.strEmailContato
        .Parameters("@FACEBOOK_VC").Value = udtJogador.strFacebook
        .Parameters("@ESCOLA_VC").Value = udtJogador.strEscola
        .Parameters("@ESTADOESCOLA_IN").Value = udtJogador.lngEstadoEscola
        .Parameters("@CIDADEESCOLA_VC").Value = udtJogador.strCidadeEscola
        .Parameters("@BAIRROESCOLA_VC").Value = udtJogador.strBairroEscola
        .Parameters("@ENDERECOESCOLA_VC").Value = udtJogador.strEnderecoEscola
        .Parameters("@REDESOCIALESCOLA_VC").Value = udtJogador.strFacebookEscola
        .Parameters("@INSTAGRAM_VC").Value = udtJogador.strInstagram
        .Parameters("@ENDERECOIMAGEM_VC").Value = IIf(udtJogador.blnTemImagem = True, udtJogador.strEnderecoImagem(), Null)
        .Parameters("@SEXO_IN").Value = udtJogador.lngSexo
        .Parameters("@NUMEROCAMISA_IN").Value = udtJogador.lngNumeroCamisa
        .Parameters("@POSICAO_IN").Value = udtJogador.lngPosicao
    End With
    gobjCmd.Execute , , adExecuteNoRecords
    
    udtJogador.lngCodigo = NZ(gobjCmd.Parameters("@CodigoOutput").Value)

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "modJogador" & vbCrLf & "No Procedimento: " & "modJogador_AdicionarJogador" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Public Sub modJogador_AlterarJogador(ByRef udtJogador As TypJogador)
    On Error GoTo Erro

    Set gobjCmd.ActiveConnection = gSMConexao.Conexao
    gobjCmd.CommandText = "dbo.USP_ALTERARJOGADOR"
    gobjCmd.CommandType = adCmdStoredProc
    gobjCmd.CommandTimeout = 1000

    With gobjCmd
        .Parameters("@ID_IN").Value = udtJogador.lngCodigo
        .Parameters("@APELIDO_VC").Value = udtJogador.strApelido
        .Parameters("@CARTEGORIA_IN").Value = udtJogador.lngCartegoria
        .Parameters("@EQUIPE_IN").Value = udtJogador.lngEquipe
        .Parameters("@NOMEATLETA_VC").Value = udtJogador.strNomeAtleta
        .Parameters("@DATANASCIMENTO_DT").Value = udtJogador.datDataNascimento
        .Parameters("@LOCALNASCIMENTO_VC").Value = udtJogador.strLocalNascimento
        .Parameters("@CERTIDAONASCIMENTO_VC").Value = udtJogador.strCertidaoNascimento
        .Parameters("@CARTORIO_VC").Value = udtJogador.strCartorio
        .Parameters("@IDENTIDADE_VC").Value = udtJogador.strIdentidade
        .Parameters("@ORGAOIDENTIDADE_VC").Value = udtJogador.strOrgaoIdentidade
        .Parameters("@NOMEPAI_VC").Value = udtJogador.strNomePai
        .Parameters("@STRNOMEMAE_VC").Value = udtJogador.strNomeMae
        .Parameters("@ESTADO_IN").Value = udtJogador.lngEstado
        .Parameters("@CIDADE_VC").Value = udtJogador.strCidade
        .Parameters("@BAIRRO_VC").Value = udtJogador.strBairro
        .Parameters("@ENDERECO_VC").Value = udtJogador.strEndereco
        .Parameters("@TELCEL1_VC").Value = udtJogador.strTelefone1
        .Parameters("@TELCEL2_VC").Value = udtJogador.strTelefone2
        .Parameters("@WPP1_BT").Value = udtJogador.blnWpp1
        .Parameters("@WPP2_BT").Value = udtJogador.blnWpp2
        .Parameters("@EMAIL_VC").Value = udtJogador.strEmailContato
        .Parameters("@FACEBOOK_VC").Value = udtJogador.strFacebook
        .Parameters("@ESCOLA_VC").Value = udtJogador.strEscola
        .Parameters("@ESTADOESCOLA_IN").Value = udtJogador.lngEstadoEscola
        .Parameters("@CIDADEESCOLA_VC").Value = udtJogador.strCidadeEscola
        .Parameters("@BAIRROESCOLA_VC").Value = udtJogador.strBairroEscola
        .Parameters("@ENDERECOESCOLA_VC").Value = udtJogador.strEnderecoEscola
        .Parameters("@REDESOCIALESCOLA_VC").Value = udtJogador.strFacebookEscola
        .Parameters("@INSTAGRAM_VC").Value = udtJogador.strInstagram
        .Parameters("@ENDERECOIMAGEM_VC").Value = IIf(udtJogador.blnTemImagem = True, udtJogador.strEnderecoImagem(), Null)
        .Parameters("@SEXO_IN").Value = udtJogador.lngSexo
        .Parameters("@NUMEROCAMISA_IN").Value = udtJogador.lngNumeroCamisa
        .Parameters("@POSICAO_IN").Value = udtJogador.lngPosicao
    End With
    gobjCmd.Execute , , adExecuteNoRecords
    
    udtJogador.lngCodigo = NZ(gobjCmd.Parameters("@CodigoOutput").Value)
    Exit Sub
Erro:
       Call MsgBox("Erro no módulo: " & "modJogador" & vbCrLf & "No Procedimento: " & "modJogador_AlterarJogador" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Public Sub modJogador_SelecionarJogadorPorCodigo(ByVal lngJogador As Long, ByRef objrs As Recordset)
On Error GoTo Erro
      
    Set gobjCmd.ActiveConnection = gSMConexao.Conexao
    gobjCmd.CommandText = "dbo.usp_SelecionarJogadorPorCodigo"
    gobjCmd.CommandType = adCmdStoredProc
    gobjCmd.CommandTimeout = 1000
    
    With gobjCmd
        .Parameters("@Jogador").Value = lngJogador
    End With
    
    Set objrs = gobjCmd.Execute

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "modJogador" & vbCrLf & "No Procedimento: " & "modJogador_SelecionarJogadorPorCodigo" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Public Sub modJogador_AdicionarAlterarFotoJogador(ByVal lngJogador As Long)
On Error GoTo Erro
      
    Set gobjCmd.ActiveConnection = gSMConexao.Conexao
    gobjCmd.CommandText = "dbo.USP_ADICIONARALTERARFOTOJOGADOR"
    gobjCmd.CommandType = adCmdStoredProc
    gobjCmd.CommandTimeout = 1000
    
    With gobjCmd
        .Parameters("@Jogador_IN").Value = lngJogador
    End With
    
    gobjCmd.Execute , , adExecuteNoRecords

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "modJogador" & vbCrLf & "No Procedimento: " & "modJogador_SelecionarJogadorPorCodigo" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Public Sub modJogador_SelecionarDadosParaRelatorioDeJogador(ByRef udtJogador As TypJogador, ByRef objrs As Recordset)
On Error GoTo Erro
      
    Set gobjCmd.ActiveConnection = gSMConexao.Conexao
    gobjCmd.CommandText = "dbo.usp_SelecionarDadosParaRelatorioDeJogador"
    gobjCmd.CommandType = adCmdStoredProc
    gobjCmd.CommandTimeout = 1000
    
    With gobjCmd
        .Parameters("@Nome_VC").Value = IIf(udtJogador.strNomeAtleta = "", Null, udtJogador.strNomeAtleta)
        .Parameters("@Apelido_VC").Value = IIf(udtJogador.strApelido = "", Null, udtJogador.strApelido)
        .Parameters("@Cartegoria_IN").Value = IIf(udtJogador.lngCartegoria = 0, Null, udtJogador.lngCartegoria)
        .Parameters("@DataNascimentoDE_DT").Value = IIf(udtJogador.datDataNascimentoDE = 0, Null, udtJogador.datDataNascimentoDE)
        .Parameters("@DataNascimentoATE_DT").Value = IIf(udtJogador.datDataNascimentoATE = 0, Null, udtJogador.datDataNascimentoATE)
        .Parameters("@UF_IN").Value = IIf(udtJogador.lngEstado = 0, Null, udtJogador.lngEstado)
        .Parameters("@Cidade_VC").Value = IIf(udtJogador.strCidade = "", Null, udtJogador.strCidade)
        .Parameters("@Bairro_VC").Value = IIf(udtJogador.strBairro = "", Null, udtJogador.strBairro)
        .Parameters("@Equipes_VC").Value = IIf(udtJogador.strEquipes = "", Null, udtJogador.strEquipes)
        .Parameters("@Sexo_IN").Value = IIf(udtJogador.lngSexo = 0, Null, udtJogador.lngSexo)
    End With
    
    Set objrs = gobjCmd.Execute

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "modJogador" & vbCrLf & "No Procedimento: " & "modJogador_SelecionarJogadorPorCodigo" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Public Sub modJogador_ApagarJogadorPorCodigo(ByVal lngJogador As Long, lngOperacao As Long)
On Error GoTo Erro
'1 - INATIVAR
'2 - EXCLUIR
      
    Set gobjCmd.ActiveConnection = gSMConexao.Conexao
    gobjCmd.CommandText = "dbo.USP_APAGARJOGADORPORCODIGO"
    gobjCmd.CommandType = adCmdStoredProc
    gobjCmd.CommandTimeout = 1000
    
    With gobjCmd
        .Parameters("@Jogador_IN").Value = lngJogador
        .Parameters("@Operacao_IN").Value = lngOperacao
    End With
    
    gobjCmd.Execute , , adExecuteNoRecords
    
Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "modJogador" & vbCrLf & "modJogador_ApagarJogadorPorCodigo" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

