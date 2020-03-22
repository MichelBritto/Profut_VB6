VERSION 5.00
Begin VB.Form frmOpcaoImpressao 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1830
   ClientLeft      =   8550
   ClientTop       =   4710
   ClientWidth     =   3345
   ControlBox      =   0   'False
   Icon            =   "frmOpcaoImpressao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3345
   Begin VB.CommandButton cmdCarteirinha 
      BackColor       =   &H80000007&
      Caption         =   "Carteitinha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1650
      MaskColor       =   &H0000C0C0&
      TabIndex        =   1
      Top             =   900
      Width           =   1365
   End
   Begin VB.CommandButton cmdFicha 
      BackColor       =   &H80000007&
      Caption         =   "Ficha de Jogador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1650
      MaskColor       =   &H0000C0C0&
      TabIndex        =   0
      Top             =   240
      Width           =   1365
   End
   Begin VB.Image imgCancelar 
      Height          =   195
      Left            =   3120
      Picture         =   "frmOpcaoImpressao.frx":038A
      Top             =   0
      Width           =   195
   End
   Begin VB.Image Image 
      Height          =   1110
      Left            =   270
      Picture         =   "frmOpcaoImpressao.frx":05D4
      Stretch         =   -1  'True
      Top             =   270
      Width           =   1140
   End
End
Attribute VB_Name = "frmOpcaoImpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCarteirinha_Click()
    frmCadastroDeJogadorV2.OpcaoImpressao = 2
    Unload Me
End Sub

Private Sub cmdFicha_Click()
    frmCadastroDeJogadorV2.OpcaoImpressao = 1
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'frmCadastroDeJogadorV2.OpcaoImpressao = 0
End Sub

Private Sub imgCancelar_Click()
    Unload Me
End Sub
