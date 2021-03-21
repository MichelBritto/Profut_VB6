VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Begin VB.Form frmAdicionarFotoJogador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adicionar Foto do Jogador"
   ClientHeight    =   7710
   ClientLeft      =   7455
   ClientTop       =   1800
   ClientWidth     =   5880
   ControlBox      =   0   'False
   Icon            =   "frmAdicionarFotoJogador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdProximaFoto 
      Caption         =   ">>"
      Height          =   285
      Left            =   1530
      TabIndex        =   1
      Top             =   2565
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmdFotoAnterior 
      Caption         =   "<<"
      Height          =   285
      Left            =   315
      TabIndex        =   0
      Top             =   2565
      Visible         =   0   'False
      Width           =   330
   End
   Begin Threed.SSCommand cmdSalvar 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   6750
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      _Version        =   196609
      PictureFrames   =   1
      Enabled         =   0   'False
      Picture         =   "frmAdicionarFotoJogador.frx":038A
      Caption         =   "Salvar"
      PictureAlignment=   0
      BevelWidth      =   1
   End
   Begin Threed.SSFrame SSFrame 
      Height          =   5835
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   10292
      _Version        =   196609
      Begin VB.Image picImagem 
         Height          =   5700
         Left            =   60
         Stretch         =   -1  'True
         Top             =   75
         Width           =   5700
      End
   End
   Begin Threed.SSFrame frame 
      Height          =   780
      Left            =   30
      TabIndex        =   4
      Top             =   5895
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   1376
      _Version        =   196609
      Begin VB.TextBox txtFoto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   90
         TabIndex        =   6
         Top             =   225
         Width           =   5190
      End
      Begin Threed.SSCommand cmdProcurar 
         Height          =   375
         Left            =   5310
         TabIndex        =   5
         Top             =   225
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         _Version        =   196609
         PictureFrames   =   1
         Enabled         =   0   'False
         Picture         =   "frmAdicionarFotoJogador.frx":0644
         PictureAlignment=   0
      End
   End
   Begin Threed.SSCommand cmdExcluirFoto 
      Height          =   375
      Left            =   180
      TabIndex        =   7
      Top             =   6750
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   661
      _Version        =   196609
      PictureFrames   =   1
      Picture         =   "frmAdicionarFotoJogador.frx":0BDE
      Caption         =   "Cancelar"
      PictureAlignment=   0
      BevelWidth      =   1
   End
   Begin MSComctlLib.StatusBar Sta 
      Align           =   2  'Align Bottom
      Height          =   210
      Left            =   0
      TabIndex        =   8
      Top             =   7500
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   370
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSCommand cmdWebcam 
      Height          =   375
      Left            =   4020
      TabIndex        =   11
      Top             =   6750
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      _Version        =   196609
      PictureFrames   =   1
      Picture         =   "frmAdicionarFotoJogador.frx":0F00
      Caption         =   "WebCam"
      PictureAlignment=   0
      BevelWidth      =   1
   End
   Begin VB.Label Label4 
      Caption         =   "Tamanho recomendado da imagem: 440 X 440"
      Height          =   225
      Left            =   1245
      TabIndex        =   10
      Top             =   7170
      Width           =   3375
   End
   Begin VB.Label lblFotos 
      Alignment       =   2  'Center
      Caption         =   "0 de 0"
      Height          =   240
      Left            =   765
      TabIndex        =   9
      Top             =   2610
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "frmAdicionarFotoJogador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mobjRSFotosProduto      As ADODB.Recordset
Dim mstrDiretorioFoto       As String
Dim mlngCodigo              As Long
Dim mstrDescricao           As String
Dim mblnCarregado           As Boolean
Dim mblnAlterouFoto         As Boolean
Public Property Let FotoWebcam(ByVal strEndereco As String)
    mstrDiretorioFoto = strEndereco
End Property

Public Property Let CodigoProduto(ByVal lngCodigo As Long)
    mlngCodigo = lngCodigo
End Property

Public Property Let DescricaoProduto(ByVal strDescricao As String)
    mstrDescricao = strDescricao
End Property

Public Property Get AlterouFoto() As Boolean
    AlterouFoto = mblnAlterouFoto
End Property

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdExcluirFoto_Click()

On Error GoTo Erro

    Unload Me

'    If Val(lblFotos.Tag) = 0 Or Dir(txtFoto.Text, vbArchive) = "" Then
'        MsgBox "Não foi possível identificar a foto a ser excluída.", vbOKOnly + vbCritical, "Aviso"
'        Exit Sub
'    End If
'
'    If MsgBox("Confirma a exclusão da foto do produto?", vbYesNo + vbQuestion, "Aviso") = vbNo Then
'        Exit Sub
'    End If
'
'    ''''''''''''''''modBDProduto_ApagarFotoProdutoPorSequencia NZ(lblCodigo.Caption), NZ(lblFotos.Tag)
'    'If genuStatusErro = SMTratarErroErro Then GoTo Erro
'
'    Kill mstrDiretorioFoto
'
'    'Call GravarPendenciaAtualizacaoSite(Val(lblCodigo.Caption), ApagarFoto, hwnd, lblCodigo.Caption & "_" & lblFotos.Tag & ".jpg")
'
'    mblnAlterouFoto = True
'    Set picImagem = Nothing
'    Call CarregarFotos


Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmAdicionarFotoJogador" & vbCrLf & "No Procedimento: " & "cmdExcluirFoto_Click" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Private Sub cmdNovo_Click()

End Sub

Private Sub cmdProcurar_Click()
10        On Error GoTo Erro
          
20        dlg.Filter = "Arquivo de Imagem (JPG) | *.jpg"
30        dlg.ShowOpen
              
40        If Dir(dlg.FileName, vbArchive) = "" Then
50            MsgBox "Arquivo Inválido", vbOKOnly + vbInformation, "Aviso"
60            Exit Sub
70        End If
          
80        txtFoto.Text = Empty
90        picImagem.Stretch = True
100       picImagem.Picture = LoadPicture(dlg.FileName)
110       txtFoto.Text = dlg.FileName
          If txtFoto.Enabled Then txtFoto.SetFocus

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmAdicionarFotoJogador" & vbCrLf & "No Procedimento: " & "cmdProcurar_Click" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Private Sub cmdSalvar_Click()
On Error GoTo Erro

    Dim strFoto                     As String
    
'    If Not Val(lblFotos.Tag) = -1 Then
'        MsgBox "Clique no botão Nova Foto para incluir outra", vbOKOnly + vbInformation, "Aviso"
'        Exit Sub
'    End If
'
    If Dir(txtFoto.Text) = "" Then
        MsgBox "Caminho da foto é inválido", vbOKOnly + vbInformation, "Aviso"
        Exit Sub
    Else
        picImagem.Stretch = True
        picImagem.Picture = LoadPicture(txtFoto.Text)
        frmCadastroDeJogadorV2.DiretorioFotoJogador = txtFoto.Text
    End If
    
    Call cmdCancelar_Click
    
    
'    gSMConexao.BeginTransaction
'    Call modBDProduto_AdicionarFotoProduto(Val(lblCodigo.Caption), FileLen(txtFoto.Text), strFoto)
'    If genuStatusErro = SMTratarErroErro Then GoTo Erro
'
'    FileCopy txtFoto.Text, strFoto
'
'    gSMConexao.CommitTransaction
'
'    mblnAlterouFoto = True
'    MsgBox "Foto do produto gravada com sucesso!", vbOKOnly + vbApplicationModal + vbApplicationModal, "Aviso"
'
'    Call CarregarFotos
'    cmdSalvar.Enabled = False
'    txtFoto.Text = Empty
'    txtFoto.Enabled = False
'    cmdProcurar.Enabled = False

Exit Sub
Erro:
    'gSMConexao.RollbackTransaction
    cmdSalvar.Enabled = False
End Sub

Private Sub cmdWebcam_Click()
    frmwebcam.Show vbModal, Me
    picImagem.Stretch = True
    picImagem.Picture = LoadPicture(mstrDiretorioFoto)
    txtFoto.Text = mstrDiretorioFoto
End Sub

Private Sub Form_Activate()
    If Not mblnCarregado Then
        mblnCarregado = True
        If mlngCodigo > 0 Then CarregarFotos
    End If
End Sub

Private Sub Form_Load()
    mblnCarregado = False
    mblnAlterouFoto = False
    
    Set picImagem.Picture = Nothing
    lblFotos.Caption = "Nova"
    lblFotos.Tag = -1
    txtFoto.Text = ""
    cmdSalvar.Enabled = True
    txtFoto.Enabled = True
    cmdProcurar.Enabled = True
    
    Sta.Panels(1).Text = gSMConexao.LoginUsuario
    Sta.Panels(1).Width = frmAdicionarFotoJogador.Width / 3
    Sta.Panels(2).Text = gSMConexao.NomeBaseDados
    Sta.Panels(2).Width = frmAdicionarFotoJogador.Width / 3
    Sta.Panels(3).Text = gSMConexao.NomeServidor
    Sta.Panels(3).Width = frmAdicionarFotoJogador.Width / 3
End Sub

Public Sub CarregarFotos()
10    On Error GoTo Erro

20        ''''''''modBDProduto_SelecionarFotosProdutoPorProduto NZ(lblCodigo.Caption), mobjRSFotosProduto
30        'If genuStatusErro = SMTratarErroErro Then GoTo Erro

40        If Not mobjRSFotosProduto.EOF Then
50            mstrDiretorioFoto = mobjRSFotosProduto!pro_foto_VC
60            Set picImagem.Picture = LoadPicture(mobjRSFotosProduto!pro_foto_VC)
70            lblFotos.Tag = mobjRSFotosProduto!pro_sequencia_IN
80            lblFotos.Caption = "1 de " & mobjRSFotosProduto.RecordCount
        
90        Else
100           Set picImagem.Picture = Nothing
110           lblFotos.Tag = 0
120           lblFotos.Caption = "0"
130           cmdExcluirFoto.Enabled = False
140       End If
          
150   Exit Sub
Erro:

160       If Err.Number = 53 Then
170           Set picImagem.Picture = Nothing
180           lblFotos.Tag = 0
190           lblFotos.Caption = "0"
200           cmdExcluirFoto.Enabled = False
210       Else
                Call MsgBox("Erro no módulo: " & "frmAdicionarFotoJogador" & vbCrLf & "No Procedimento: " & "CarregarFotos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
230       End If
End Sub

Private Sub cmdProximaFoto_Click()
On Error GoTo Erro
    
    If mobjRSFotosProduto Is Nothing Then Exit Sub
    If mobjRSFotosProduto.State = 0 Then Exit Sub
    If mobjRSFotosProduto.RecordCount = 0 Then Exit Sub
    
    If mobjRSFotosProduto.EOF Then Exit Sub
        
    mobjRSFotosProduto.MoveNext
    
    If mobjRSFotosProduto.EOF Then Exit Sub
    
    lblFotos.Caption = mobjRSFotosProduto.AbsolutePosition & " de " & mobjRSFotosProduto.RecordCount
    lblFotos.Tag = mobjRSFotosProduto!pro_sequencia_IN
    Set picImagem.Picture = LoadPicture(mobjRSFotosProduto!pro_foto_VC)
    txtFoto.Text = mobjRSFotosProduto!pro_foto_VC
    

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmAdicionarFotoJogador" & vbCrLf & "No Procedimento: " & "cmdProximaFoto_Click" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Private Sub cmdFotoAnterior_Click()
On Error GoTo Erro

    If mobjRSFotosProduto Is Nothing Then Exit Sub
    If mobjRSFotosProduto.State = 0 Then Exit Sub
    If mobjRSFotosProduto.RecordCount = 0 Then Exit Sub
    
    If mobjRSFotosProduto.BOF Then Exit Sub
  
    mobjRSFotosProduto.MovePrevious
    
    If mobjRSFotosProduto.BOF Then Exit Sub
    
    lblFotos.Caption = mobjRSFotosProduto.AbsolutePosition & " de " & mobjRSFotosProduto.RecordCount
    lblFotos.Tag = mobjRSFotosProduto!pro_sequencia_IN
    Set picImagem.Picture = LoadPicture(mobjRSFotosProduto!pro_foto_VC)
    txtFoto.Text = mobjRSFotosProduto!pro_foto_VC
    

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmAdicionarFotoJogador" & vbCrLf & "No Procedimento: " & "cmdFotoAnterior_Click" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Private Sub Form_Resize()
'    If WindowState = vbMinimized Then Exit Sub
'    If Height <> 3855 Then Height = 3855
'    If Width <> 8280 Then Width = 8280
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnCarregado = False
End Sub

Private Sub lblCodigo_Click()

End Sub

Private Sub lblProduto_Click()

End Sub

Private Sub picImagem_DblClick()
    Call CarregarFotos
End Sub

