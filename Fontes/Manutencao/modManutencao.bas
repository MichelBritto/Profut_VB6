Attribute VB_Name = "modManutencao"
Option Explicit

Public Sub modManutencao_AdicionarAlterarUsuario(ByVal strLogin As String, ByVal strNome As String, ByVal lngCargo As Long, Optional ByVal lngUsuario As Long, _
                                     Optional ByVal strTelefone As String, Optional ByVal strEmail As String)

10    On Error GoTo Erro

20        Set gobjCmd.ActiveConnection = gSMConexao.Conexao
30        gobjCmd.CommandText = "dbo.usp_AdicionarAlterarUsuario"
40        gobjCmd.CommandType = adCmdStoredProc
50        gobjCmd.CommandTimeout = 1000

60        With gobjCmd
70           .Parameters("@login_VC").Value = strLogin
80           .Parameters("@nomecompleto_VC").Value = strNome
90           .Parameters("@email_VC").Value = IIf(strEmail = "", Null, strEmail)
100          .Parameters("@telefone_VC").Value = IIf(strTelefone = "", Null, strTelefone)
110          .Parameters("@usuario_IN").Value = IIf(lngUsuario = 0, Null, lngUsuario)
120          .Parameters("@setorinterno_IN").Value = IIf(lngCargo = 0, Null, lngCargo)
130       End With
140       gobjCmd.Execute , , adExecuteNoRecords

150   Exit Sub
Erro:
160      Call MsgBox("Erro no módulo: " & "modManutencao" & vbCrLf & "modEquipe_AdicionarEquipe" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Public Sub modManutencao_SelecionarUsuario(ByRef objRs As Recordset, Optional ByVal lngUsuario As Long)
10    On Error GoTo Erro
            
20        Set gobjCmd.ActiveConnection = gSMConexao.Conexao
30        gobjCmd.CommandText = "usp_SelecionarUsuarios"
40        gobjCmd.CommandType = adCmdStoredProc
50        gobjCmd.CommandTimeout = 1000
          
60        With gobjCmd
70            .Parameters("@Usuario_IN").Value = IIf(lngUsuario = 0, Null, lngUsuario)
80        End With
          
90        Set objRs = gobjCmd.Execute

100   Exit Sub
Erro:
110      Call MsgBox("Erro no módulo: " & "modManutencao" & vbCrLf & "modManutencao_SelecionarUsuario" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")


End Sub
