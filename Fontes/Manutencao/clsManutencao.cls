VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsManutencao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function Show(ByRef SMConexao As clsConexaoMC, _
                     Optional ByVal strCaption As String, _
                     Optional ByVal modal As FormShowConstants = vbModeless, _
                     Optional ByVal OwnerForm As Object = Nothing)
On Error GoTo Erro
        
    Set gSMConexao = SMConexao

    frmManutencao.Caption = strCaption
    frmManutencao.Show modal, OwnerForm

Exit Function
Erro:
    Select Case Err.Number
        Case 373
                frmManutencao.Show vbModeless, OwnerForm
        Case 401
                frmManutencao.Show vbModal, OwnerForm
        Case Else
                Call MsgBox("Erro no módulo: " & "clsManutencao" & vbCrLf & "No Procedimento: " & "Show" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
    End Select
End Function

