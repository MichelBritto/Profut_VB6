VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form frmwebcam 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capturando Imagens da Web Cam"
   ClientHeight    =   7935
   ClientLeft      =   5415
   ClientTop       =   1785
   ClientWidth     =   7260
   ControlBox      =   0   'False
   Icon            =   "frmwebcam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4500
      Top             =   3000
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3630
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   2505
   End
   Begin VB.CommandButton cmdCapturaImagem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Capturar Imagem da Web Cam"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7200
      Width           =   2445
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   7000
      Left            =   120
      ScaleHeight     =   6945
      ScaleWidth      =   6945
      TabIndex        =   0
      Top             =   90
      Width           =   7000
   End
   Begin MSComctlLib.StatusBar Sta 
      Align           =   2  'Align Bottom
      Height          =   210
      Left            =   0
      TabIndex        =   3
      Top             =   7725
      Width           =   7260
      _ExtentX        =   12806
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
End
Attribute VB_Name = "frmwebcam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_CAP_DRIVER_CONNECT As Long = 1034
Private Const WM_CAP_DRIVER_DISCONNECT As Long = 1035
Private Const WM_CAP_GRAB_FRAME As Long = 1084
Private Const WM_CAP_EDIT_COPY As Long = 1054
Private Const WM_CAP_DLG_VIDEOFORMAT As Long = 1065
Private Const WM_CAP_DLG_VIDEOSOURCE As Long = 1066
Private Const WM_CLOSE = &H10
Private mCapHwnd As Long
Private mStrEnderecoFoto As String
Private Sub AtivaVideoContinuo()
On Error Resume Next
 Timer1.Enabled = True
 Timer1.Interval = 50
End Sub
Private Sub DesativaVideoContinuo()
On Error Resume Next
 Timer1.Enabled = False
 Timer1.Interval = 50
End Sub
Private Sub cmdCapturaImagem_Click()
On Error Resume Next
    'Captura a imagem atual
    Dim lngrnd1 As Long
    Dim lngrnd2 As Long
    Dim lngrnd3 As Long
    Dim lngrnd4 As Long
    Dim lngrnd5 As Long
    
    lngrnd1 = (Rnd() * 20)
    lngrnd2 = (Rnd() * 20)
    lngrnd3 = (Rnd() * 20)
    lngrnd4 = (Rnd() * 20)
    lngrnd5 = (Rnd() * 20)
    
    Clipboard.Clear
    SendMessage mCapHwnd, WM_CAP_GRAB_FRAME, 0, 0
    SendMessage mCapHwnd, WM_CAP_EDIT_COPY, 0, 0
    Picture1.Picture = Clipboard.GetData
    DesativaVideoContinuo
    mStrEnderecoFoto = "C:\ProFut\IMG\" & lngrnd1 & lngrnd2 & lngrnd3 & lngrnd4 & lngrnd5 & ".jpg"
retentar:
    On Error GoTo criardiretorio
    Call SavePicture(Picture1.Image, mStrEnderecoFoto)
    frmAdicionarFotoJogador.FotoWebcam = mStrEnderecoFoto
    Unload Me
    Exit Sub
criardiretorio:
    If Dir("C:\ProFut\IMG", vbDirectory) = "" Then
        MkDir "C:\ProFut\IMG"
    End If
    
    GoTo retentar
   
End Sub
Private Sub EncerraWebCam()
On Error Resume Next
 'Desliga a c�mera
   SendMessage mCapHwnd, WM_CAP_DRIVER_DISCONNECT, 0, 0
End Sub
Private Sub IniciaWebCam()
On Error Resume Next
'Inicia a c�mera
   mCapHwnd = capCreateCaptureWindow("captura Janela", 0, 0, 0, 320, 240, Me.hwnd, 0)
   SendMessage mCapHwnd, WM_CAP_DRIVER_CONNECT, 0, 0
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
On Error Resume Next
    IniciaWebCam
    AtivaVideoContinuo
End Sub

Private Sub Form_Load()
    Sta.Panels(1).Text = gSMConexao.LoginUsuario
    Sta.Panels(1).Width = frmwebcam.Width / 3
    Sta.Panels(2).Text = gSMConexao.NomeBaseDados
    Sta.Panels(2).Width = frmwebcam.Width / 3
    Sta.Panels(3).Text = gSMConexao.NomeServidor
    Sta.Panels(3).Width = frmwebcam.Width / 3
End Sub

Private Sub Form_Terminate()
On Error Resume Next
   'Desliga a c�mera
   SendMessage mCapHwnd, WM_CAP_DRIVER_DISCONNECT, 0, 0
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
'Exibe imagem continua no pictubox
   Clipboard.Clear
   SendMessage mCapHwnd, WM_CAP_GRAB_FRAME, 0, 0
   SendMessage mCapHwnd, WM_CAP_EDIT_COPY, 0, 0
   Picture1.Picture = Clipboard.GetData
   DoEvents
End Sub
