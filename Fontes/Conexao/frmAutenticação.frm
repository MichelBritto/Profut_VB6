VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmAutenticacao 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2505
   ClientLeft      =   11235
   ClientTop       =   7200
   ClientWidth     =   3855
   ControlBox      =   0   'False
   Icon            =   "frmAutenticação.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkAlterarSenha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Caption         =   "Desejo Alterar a senha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   90
      TabIndex        =   9
      Top             =   1830
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "C&ancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   75
      TabIndex        =   10
      Top             =   2130
      Width           =   1890
   End
   Begin VB.CommandButton cmdConfirmar 
      Appearance      =   0  'Flat
      Caption         =   "Prosseguir >>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2060
      TabIndex        =   11
      Top             =   2130
      Width           =   1740
   End
   Begin Threed.SSFrame fraUsuarioSenha 
      Height          =   1755
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   3096
      _Version        =   196609
      Begin VB.TextBox txtServidor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   855
         TabIndex        =   5
         Text            =   "FALCO-PC"
         Top             =   990
         Width           =   2820
      End
      Begin VB.TextBox txtBase 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   855
         TabIndex        =   7
         Text            =   "DBFUTEBOL"
         Top             =   1365
         Width           =   2820
      End
      Begin VB.TextBox txt_senha 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   855
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   "mi1401"
         Top             =   480
         Width           =   2820
      End
      Begin VB.TextBox txt_usuario 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   855
         TabIndex        =   1
         Text            =   "sa"
         Top             =   105
         Width           =   2820
      End
      Begin VB.Label lblServidor 
         Caption         =   "Servidor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   105
         TabIndex        =   6
         Top             =   1065
         Width           =   825
      End
      Begin VB.Label lblBase 
         Caption         =   "Base"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   8
         Top             =   1425
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Senha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   4
         Top             =   540
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Usuário"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   105
         TabIndex        =   2
         Top             =   180
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmAutenticacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub FecharPeloMenu()
    If gblnLoginRealizado = True Then Unload frmAutenticacao
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdConfirmar_Click()

Dim fso As New Scripting.FileSystemObject
Dim ts As Scripting.TextStream


On Error GoTo Erro
    
    gstrLoginUsuario = txt_usuario.Text
    gstrSenhaUsuario = txt_senha.Text
    gstrNomeBaseDados = txtBase.Text
    gstrNomeServidor = txtServidor.Text
    
    
    'Set ts = fso.OpenTextFile(AppEx.Path("C:\Users\FALCO\Desktop\trabson"), ForWriting, True)
    
'    ts.Write txt_usuario.Text
'    ts.Write txtBase.Text
'    ts.Write txtServidor.Text
'    ts.Close
'    Set ts = Nothing
    
    gSMConexao.conectar
    
    If gobjConn.State = adStateOpen Then
        Call cmdCancelar_Click
        Exit Sub
    Else
        GoTo Erro
    End If
Exit Sub
Erro:
    MsgBox "Falha ao realizar login. Verifique se as informações estão corretas.", vbOKOnly + vbInformation, "Atenção"
    gblnLoginRealizado = False
End Sub

Private Sub txt_senha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdConfirmar_Click
    End If
End Sub

Private Sub txt_usuario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdConfirmar_Click
    End If
End Sub


Private Sub txtBase_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdConfirmar_Click
    End If
End Sub


Private Sub txtServidor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdConfirmar_Click
    End If
End Sub
