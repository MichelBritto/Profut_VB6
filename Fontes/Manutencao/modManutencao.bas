Attribute VB_Name = "modManutencao"
Option Explicit

Public Sub modManutencao_AdicionarAlterarUsuario(ByVal strLogin As String, ByVal strNome As String, ByVal lngCargo As Long, Optional ByVal lngUsuario As Long, _
                                     Optional ByVal strTelefone As String, Optional ByVal strEmail As String, Optional ByVal lngClube As Long)

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
130          .Parameters("@clube_IN").Value = IIf(lngClube = 0, Null, lngClube)
140       End With
150       gobjCmd.Execute , , adExecuteNoRecords

160   Exit Sub
Erro:
170      Call MsgBox("Erro no módulo: " & "modManutencao" & vbCrLf & "modEquipe_AdicionarEquipe" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Public Sub modManutencao_SelecionarUsuario(ByRef objrs As Recordset, Optional ByVal lngUsuario As Long)
10    On Error GoTo Erro
      
20        Set gobjCmd.ActiveConnection = gSMConexao.Conexao
30        gobjCmd.CommandText = "usp_SelecionarUsuarios"
40        gobjCmd.CommandType = adCmdStoredProc
50        gobjCmd.CommandTimeout = 1000
    
60        With gobjCmd
70              .Parameters("@Usuario_IN").Value = IIf(lngUsuario = 0, Null, lngUsuario)
80        End With
    
90        Set objrs = gobjCmd.Execute

100   Exit Sub
Erro:
110      Call MsgBox("Erro no módulo: " & "modManutencao" & vbCrLf & "modManutencao_SelecionarUsuario" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")


End Sub

Public Sub modManutencao_SelecionarCargos(ByRef objrs As Recordset, Optional ByVal lngCargo As Long, Optional ByVal blnSomenteAtivo As Boolean = True)
10    On Error GoTo Erro
            
20        Set gobjCmd.ActiveConnection = gSMConexao.Conexao
30        gobjCmd.CommandText = "usp_SelecionarCargos"
40        gobjCmd.CommandType = adCmdStoredProc
50        gobjCmd.CommandTimeout = 1000
          
60        With gobjCmd
70            .Parameters("@Cargo_IN").Value = IIf(lngCargo = 0, Null, lngCargo)
80            .Parameters("@Ativo_BT").Value = blnSomenteAtivo
90        End With
          
100       Set objrs = gobjCmd.Execute


110   Exit Sub
Erro:
120      Call MsgBox("Erro no módulo: " & "modManutencao" & vbCrLf & "modManutencao_SelecionarCargos" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Public Sub modManutencao_SelecionarPermissao(ByRef objrs As Recordset, Optional ByVal lngPermissao As Long)
10    On Error GoTo Erro
            
20        Set gobjCmd.ActiveConnection = gSMConexao.Conexao
30        gobjCmd.CommandText = "usp_SelecionarPermissao"
40        gobjCmd.CommandType = adCmdStoredProc
50        gobjCmd.CommandTimeout = 1000
          
60        With gobjCmd
70            .Parameters("@Permissao_IN").Value = IIf(lngPermissao = 0, Null, lngPermissao)
80        End With
90        Set objrs = gobjCmd.Execute


100   Exit Sub
Erro:
110      Call MsgBox("Erro no módulo: " & "modManutencao" & vbCrLf & "modManutencao_SelecionarCargos" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Public Sub modManutencao_SelecionarPermissaoPorUsuario(ByVal lngUsuario As Long, ByRef objrs As Recordset)
10    On Error GoTo Erro
            
20        Set gobjCmd.ActiveConnection = gSMConexao.Conexao
30        gobjCmd.CommandText = "usp_SelecionarPermissaoPorUsuario"
40        gobjCmd.CommandType = adCmdStoredProc
50        gobjCmd.CommandTimeout = 1000
          
60        With gobjCmd
70            .Parameters("@Usuario_IN").Value = lngUsuario
80        End With
          
90        Set objrs = gobjCmd.Execute
100   Exit Sub

Erro:
110      Call MsgBox("Erro no módulo: " & "modManutencao" & vbCrLf & "modManutencao_SelecionarPermissaoPorUsuario" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Public Sub modManutencao_AdicionarAlterarPermissaoPorUsuario(ByVal lngUsuario As Long, ByVal lngPermissao As Long, ByVal blnStatus As Boolean)
10    On Error GoTo Erro
            
20        Set gobjCmd.ActiveConnection = gSMConexao.Conexao
30        gobjCmd.CommandText = "usp_AdicionarAlterarPermissaoPorUsuario"
40        gobjCmd.CommandType = adCmdStoredProc
50        gobjCmd.CommandTimeout = 1000
          
60        With gobjCmd
70            .Parameters("@Usuario_IN").Value = lngUsuario
80            .Parameters("@Permissao_IN").Value = lngPermissao
90            .Parameters("@Status_BT").Value = IIf(blnStatus = True, 1, 0)
100       End With
          
110       gobjCmd.Execute , , adExecuteNoRecords

120   Exit Sub
Erro:
130      Call MsgBox("Erro no módulo: " & "modManutencao" & vbCrLf & "modManutencao_AdicionarAlterarPermissaoPorUsuario" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Public Sub modManutencao_AdicionarAlterarCargo(ByVal strDescricao As String, ByVal blnAtivo As Boolean, Optional ByVal lngCargo As Long)
10    On Error GoTo Erro
            
20        Set gobjCmd.ActiveConnection = gSMConexao.Conexao
30        gobjCmd.CommandText = "usp_AdicionarAlterarCargo"
40        gobjCmd.CommandType = adCmdStoredProc
50        gobjCmd.CommandTimeout = 1000
          
60        With gobjCmd
70            .Parameters("@Descricao_VC").Value = strDescricao
80            .Parameters("@Ativo_BT").Value = IIf(blnAtivo = True, 1, 0)
90            .Parameters("@Cargo_IN").Value = IIf(lngCargo = 0, Null, lngCargo)
100       End With
          
110       gobjCmd.Execute , , adExecuteNoRecords

120   Exit Sub
Erro:
130      Call MsgBox("Erro no módulo: " & "modManutencao" & vbCrLf & "modManutencao_AdicionarAlterarCargo" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")


End Sub

Public Sub modManutencao_AdicionarAlterarPosicao(ByVal strDescricao As String, ByVal blnAtivo As Boolean, Optional ByVal lngPosicao As Long)
10    On Error GoTo Erro
            
20        Set gobjCmd.ActiveConnection = gSMConexao.Conexao
30        gobjCmd.CommandText = "usp_AdicionarAlterarPosicao"
40        gobjCmd.CommandType = adCmdStoredProc
50        gobjCmd.CommandTimeout = 1000
          
60        With gobjCmd
70            .Parameters("@Descricao_VC").Value = strDescricao
80            .Parameters("@Ativo_BT").Value = IIf(blnAtivo = True, 1, 0)
90            .Parameters("@Posicao_IN").Value = IIf(lngPosicao = 0, Null, lngPosicao)
100       End With
          
110       gobjCmd.Execute , , adExecuteNoRecords

120   Exit Sub
Erro:
130      Call MsgBox("Erro no módulo: " & "modManutencao" & vbCrLf & "modManutencao_AdicionarAlterarPosicao" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Public Sub modManutencao_SelecionarPosicoesAtleta(ByRef objrs As Recordset, Optional ByVal lngPosicao As Long, Optional ByVal blnSomenteAtivo As Boolean = True)
10    On Error GoTo Erro
            
20        Set gobjCmd.ActiveConnection = gSMConexao.Conexao
30        gobjCmd.CommandText = "usp_SelecionarPosicoesAtleta"
40        gobjCmd.CommandType = adCmdStoredProc
50        gobjCmd.CommandTimeout = 1000
          
60        With gobjCmd
70            .Parameters("@Posicao_IN").Value = IIf(lngPosicao = 0, Null, lngPosicao)
80            .Parameters("@Ativo_BT").Value = blnSomenteAtivo
90        End With
          
100       Set objrs = gobjCmd.Execute


110   Exit Sub
Erro:
120      Call MsgBox("Erro no módulo: " & "modManutencao" & vbCrLf & "modManutencao_SelecionarPosicoesAtleta" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")


End Sub

Public Sub modManutencao_AlterarSenhaPorLoginESenha(ByVal strSenhaAtual As String, strSenhaNova As String, ByRef blnResultado As Boolean)
10    On Error GoTo Erro
            
            
20        Set gobjCmd.ActiveConnection = gSMConexao.Conexao
30        gobjCmd.CommandText = "usp_AlterarSenhaPorLoginESenha"
40        gobjCmd.CommandType = adCmdStoredProc
50        gobjCmd.CommandTimeout = 1000
          
60        With gobjCmd
70            .Parameters("@Login_VC").Value = gSMConexao.LoginUsuario
80            .Parameters("@SenhaAntiga_VC").Value = strSenhaAtual
90            .Parameters("@SenhaNova_VC").Value = strSenhaNova
100       End With
          
110       gobjCmd.Execute , , adExecuteNoRecords
          
120       blnResultado = NB(gobjCmd.Parameters("@Resultado_BT").Value)

130   Exit Sub
Erro:
140      Call MsgBox("Erro no módulo: " & "modManutencao" & vbCrLf & "modManutencao_AlterarSenhaPorLoginESenha" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")


End Sub
