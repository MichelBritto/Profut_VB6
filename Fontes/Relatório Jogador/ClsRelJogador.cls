VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsRelJogador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function Show(ByRef SMConexao As clsConexaoMC, _
                     Optional ByVal strCaption As String, _
                     Optional ByVal modal As FormShowConstants = vbModeless, _
                     Optional ByVal OwnerForm As Object = Nothing, _
                     Optional ByVal blnCarregarProcurar As Boolean = 0)
On Error GoTo Erro
        
    Set gSMConexao = SMConexao

    frmRelJogador.Caption = strCaption
    frmRelJogador.CarregadoViaProcurar = blnCarregarProcurar
    frmRelJogador.Show modal, OwnerForm

Exit Function
Erro:
    Select Case Err.Number
        Case 373
                frmRelJogador.Show vbModeless, OwnerForm
        Case 401
                frmRelJogador.Show vbModal, OwnerForm
        Case Else
            Call MsgBox("Erro no módulo: " & "ClsRelJogador" & vbCrLf & "No Procedimento: " & "Show" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
    End Select

End Function

Public Function ID() As Integer
On Error GoTo Erro
      
    ID = frmRelJogador.IDJogador

Exit Function
Erro:
   Call MsgBox("Erro no módulo: " & "ClsRelJogador" & vbCrLf & "No Procedimento: " & "ID" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Function



