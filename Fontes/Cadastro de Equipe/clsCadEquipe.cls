VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCadEquipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function Show(ByRef SMConexao As clsConexaoMC, _
                     ByVal strCaption As String, _
                     Optional ByVal modal As FormShowConstants = vbModeless, _
                     Optional ByVal OwnerForm As Object = Nothing)
On Error GoTo Erro
        
    Set gSMConexao = SMConexao

    'frmCadastroDeEquipe.Caption = strCaption
    'frmCadastroDeEquipe.Show modal, OwnerForm
    frmCadastroDeEquipeV2.Show modal, OwnerForm

Exit Function
Erro:
    Select Case Err.Number
        Case 373
                frmCadastroDeEquipeV2.Show vbModeless, OwnerForm
            Resume Next
            
        Case 401
                frmCadastroDeEquipeV2.Show vbModal, OwnerForm
            Resume Next
        Case Else
            Call MsgBox("clsCadEquipe", "Show", Err.Description, Err.Number, Erl)
    End Select

End Function



