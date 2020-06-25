VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCadJogador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Public Function Show(ByRef SMConexao As clsConexaoMC, _
                     Optional ByVal strCaption As String, _
                     Optional ByVal modal As FormShowConstants = vbModeless, _
                     Optional ByVal OwnerForm As Object = Nothing, _
                     Optional ByVal lngJogador As Long = 0, _
                     Optional ByVal blnCarregando As Boolean = False)
On Error GoTo Erro
        
    Set gSMConexao = SMConexao

    'frmCadastroDeJogador.Caption = strCaption
    'frmCadastroDeJogador.Show modal, OwnerForm
    frmCadastroDeJogadorV2.Jogador = lngJogador
    frmCadastroDeJogadorV2.Carregando = blnCarregando
    frmCadastroDeJogadorV2.Show modal, OwnerForm

Exit Function
Erro:
    Select Case Err.Number
        Case 373
                frmCadastroDeJogadorV2.Show vbModeless, OwnerForm
            Resume Next
            
        Case 401
                frmCadastroDeJogadorV2.Show vbModal, OwnerForm
            Resume Next
        Case Else
            Call MsgBox("clsCadJogador", "Show", Err.Description, Err.Number, Erl)
    End Select

End Function


