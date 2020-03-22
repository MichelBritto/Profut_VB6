Attribute VB_Name = "modEquipe"
Option Explicit


Type TypEquipe
    lngCodigo As Long
    strNome As String
    strSigla As String
    strResponsavel As String
    strContato1 As String
    blnWpp1 As Boolean
    strContato2 As String
    blnWpp2 As Boolean
    strEmailContato As String
    strEnderecoImagem() As Byte
End Type

Public Sub modEquipe_AdicionarEquipe(ByRef udtEquipe As TypEquipe)

On Error GoTo Erro

    Set gobjCmd.ActiveConnection = gSMConexao.Conexao
    gobjCmd.CommandText = "dbo.USP_ADICIONAREQUIPE"
    gobjCmd.CommandType = adCmdStoredProc
    gobjCmd.CommandTimeout = 1000

    With gobjCmd
       .Parameters("@NOME_VC").Value = udtEquipe.strNome
       .Parameters("@SIGLA_VC").Value = udtEquipe.strSigla
       .Parameters("@RESPONSAVEL_VC").Value = udtEquipe.strResponsavel
       .Parameters("@CONTATO_VC1").Value = udtEquipe.strContato1
       .Parameters("@WHATSAPP1_BT").Value = udtEquipe.blnWpp1
       .Parameters("@CONTATO2_VC").Value = udtEquipe.strContato2
       .Parameters("@WHATSAP2_BT").Value = udtEquipe.blnWpp2
       .Parameters("@ENDERECOIMAGEM_VC").Value = udtEquipe.blnWpp2
       .Parameters("@EMAILCONTATO_VC").Value = udtEquipe.strEmailContato
       .Parameters("@ENDERECOIMAGEM_VC").Value = udtEquipe.strEnderecoImagem
    End With
    gobjCmd.Execute , , adExecuteNoRecords
    
    udtEquipe.lngCodigo = NZ(gobjCmd("@CODIGO_IN").Value)

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "modEquipe" & vbCrLf & "No Procedimento: " & "modEquipe_AdicionarEquipe" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Public Sub modEquipe_AlterarEquipe(ByRef udtEquipe As TypEquipe)
On Error GoTo Erro

    Set gobjCmd.ActiveConnection = gSMConexao.Conexao
    gobjCmd.CommandText = "dbo.USP_ALTERAREQUIPE"
    gobjCmd.CommandType = adCmdStoredProc
    gobjCmd.CommandTimeout = 1000
    
    With gobjCmd
        .Parameters("@EQUIPE_IN").Value = udtEquipe.lngCodigo
        .Parameters("@NOME_VC").Value = udtEquipe.strNome
        .Parameters("@SIGLA_VC").Value = udtEquipe.strSigla
        .Parameters("@RESPONSAVEL_VC").Value = udtEquipe.strResponsavel
        .Parameters("@CONTATO_VC1").Value = udtEquipe.strContato1
        .Parameters("@WHATSAPP1_BT").Value = udtEquipe.blnWpp1
        .Parameters("@CONTATO2_VC").Value = udtEquipe.strContato2
        .Parameters("@WHATSAP2_BT").Value = udtEquipe.blnWpp2
        .Parameters("@ENDERECOIMAGEM_VC").Value = udtEquipe.strEnderecoImagem()
        .Parameters("@EMAILCONTATO_VC").Value = udtEquipe.strEmailContato
    End With
    gobjCmd.Execute , , adExecuteNoRecords
    
    'udtEquipe.lngCodigo = NZ(gobjCmd("@CODIGO_IN").Value)

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "modEquipe" & vbCrLf & "No Procedimento: " & "modEquipe_AlterarEquipe" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Public Sub modEquipe_SelecionarEquipePorCodigo(ByVal lngCodigo As Long, ByRef objrs As Recordset)
On Error GoTo Erro
    
    Set gobjCmd.ActiveConnection = gSMConexao.Conexao
    gobjCmd.CommandText = "dbo.USP_SELECIONAREQUIPEPORCODIGO"
    gobjCmd.CommandType = adCmdStoredProc
    gobjCmd.CommandTimeout = 1000
    
    With gobjCmd
       .Parameters("@EQUIPE_IN").Value = lngCodigo
    End With
    
    Set objrs = gobjCmd.Execute
    
    Exit Sub
Erro:
       Call MsgBox("Erro no módulo: " & "modEquipe" & vbCrLf & "No Procedimento: " & "modEquipe_SelecionarEquipePorCodigo" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub





