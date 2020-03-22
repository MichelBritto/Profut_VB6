VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmMenu 
   Caption         =   "ProFut - Menu Principal"
   ClientHeight    =   6375
   ClientLeft      =   4515
   ClientTop       =   2340
   ClientWidth     =   10860
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   10860
   Begin VB.Frame fraPrincipal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "{"
      Height          =   6375
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   10785
      Begin VB.PictureBox logo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4515
         Left            =   4890
         Picture         =   "frmMenu.frx":038A
         ScaleHeight     =   4515
         ScaleWidth      =   4515
         TabIndex        =   2
         Top             =   720
         Width           =   4515
      End
      Begin VB.Frame fraAcessos 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6195
         Left            =   30
         TabIndex        =   1
         Top             =   120
         Width           =   3435
         Begin Threed.SSCommand cmdCadastroDeJogador 
            Height          =   1065
            Left            =   30
            TabIndex        =   3
            Top             =   60
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1879
            _Version        =   196609
            ForeColor       =   14737632
            BackColor       =   0
            PictureMaskColor=   16777152
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   11.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmMenu.frx":2729
            Caption         =   "Cadastro de Jogador"
            Alignment       =   1
            PictureAlignment=   9
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdCadastroEquipe 
            Height          =   1065
            Left            =   30
            TabIndex        =   4
            Top             =   1290
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1879
            _Version        =   196609
            ForeColor       =   14737632
            BackColor       =   0
            PictureMaskColor=   16777152
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   11.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmMenu.frx":2ADB
            Caption         =   "Cadastro de Equipe"
            Alignment       =   1
            PictureAlignment=   9
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdRelatorioJogador 
            Height          =   1065
            Left            =   30
            TabIndex        =   5
            Top             =   2520
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1879
            _Version        =   196609
            ForeColor       =   14737632
            BackColor       =   0
            PictureMaskColor=   16777152
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   11.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmMenu.frx":2E8D
            Caption         =   "Relatório de Jogador"
            Alignment       =   1
            PictureAlignment=   9
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSair 
            Height          =   1065
            Left            =   30
            TabIndex        =   6
            Top             =   5010
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1879
            _Version        =   196609
            ForeColor       =   14737632
            BackColor       =   0
            PictureMaskColor=   16777152
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   11.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmMenu.frx":323F
            Caption         =   "Sair"
            Alignment       =   1
            PictureAlignment=   9
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdManutencao 
            Height          =   1065
            Left            =   30
            TabIndex        =   7
            Top             =   3780
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1879
            _Version        =   196609
            ForeColor       =   14737632
            BackColor       =   0
            PictureMaskColor=   16777152
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   11.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmMenu.frx":35F1
            Caption         =   "Manutenção"
            Alignment       =   1
            PictureAlignment=   9
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCadastroDeJogador_Click()
    Dim objJogador As clsCadJogador
    Set objJogador = New clsCadJogador
    
    If Not gSMConexao Is Nothing Then
        If gSMConexao.EstadoConexaoBD = adStateOpen Then
            
            objJogador.Show gSMConexao, "ProFut - Cadastro de Jogador"
            Exit Sub
        Else
           gSMConexao.conectar
        End If
    End If
End Sub


Private Sub cmdCadastroEquipe_Click()
    Dim objEquipe As clsCadEquipe
    Set objEquipe = New clsCadEquipe
    
    If Not gSMConexao Is Nothing Then
        If gSMConexao.EstadoConexaoBD = adStateOpen Then
            
            objEquipe.Show gSMConexao, "Profut - Cadastro de Equipe"
            Exit Sub
        Else
           gSMConexao.conectar
        End If
    End If
End Sub


Private Sub cmdCadastroEquipe_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H000080FF&"
    cmdCadastroEquipe.ForeColor = &H80FF&
End Sub

Private Sub cmdCadastroEquipe_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H00000000&"
    cmdCadastroEquipe.ForeColor = vbWhite
End Sub


Private Sub cmdManutencao_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H000080FF&"
    cmdManutencao.ForeColor = &H80FF&
End Sub

Private Sub cmdManutencao_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H00000000&"
    cmdManutencao.ForeColor = vbWhite
End Sub

Private Sub cmdRelatorioJogador_Click()
    Dim ObjRelatorioJogador As ClsRelJogador
    Set ObjRelatorioJogador = New ClsRelJogador
    
    If Not gSMConexao Is Nothing Then
        If gSMConexao.EstadoConexaoBD = adStateOpen Then
            
            ObjRelatorioJogador.Show gSMConexao, "ProFut - Relatório de Jogador"
            Exit Sub
        Else
            gSMConexao.conectar
        End If
    End If
End Sub


Private Sub cmdRelatorioJogador_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H000080FF&"
    cmdRelatorioJogador.ForeColor = &H80FF&
End Sub

Private Sub cmdRelatorioJogador_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H00000000&"
    cmdRelatorioJogador.ForeColor = vbWhite
End Sub

Private Sub cmdSair_Click()
    frmPrincipal.FecharPrograma = True
End Sub
Private Sub cmdSair_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H000080FF&"
    cmdSair.ForeColor = &H80FF&
End Sub

Private Sub cmdSair_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H00000000&"
    cmdSair.ForeColor = vbWhite
End Sub

Private Sub cmdCadastroDeJogador_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H000080FF&"
    cmdCadastroDeJogador.ForeColor = &H80FF&
End Sub

Private Sub cmdCadastroDeJogador_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H00000000&"
    cmdCadastroDeJogador.ForeColor = vbWhite
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmPrincipal.FecharPrograma = True
End Sub

