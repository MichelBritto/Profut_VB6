VERSION 5.00
Begin VB.Form frmPrincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProFut - Programa de Recursos Orientados ao Futebol"
   ClientHeight    =   7605
   ClientLeft      =   4485
   ClientTop       =   2340
   ClientWidth     =   10845
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   10845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPrincipal 
      BackColor       =   &H00FFFFFF&
      Height          =   7635
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   10785
      Begin VB.PictureBox logo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   2670
         Picture         =   "frmPrincipal.frx":038A
         ScaleHeight     =   5175
         ScaleWidth      =   4995
         TabIndex        =   1
         Top             =   720
         Width           =   4995
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mblnLoginRealizado  As Boolean

Public Property Let FecharPrograma(blnMensagem As Boolean)
    If blnMensagem Then
        If MsgBox("Deseja realmente sair do programa?", vbYesNo + vbInformation, "Atenção!") = vbYes Then End
    Else
        End
    End If
End Property


Private Sub Form_Activate()

    frmPrincipal.Visible = True
    Set gSMConexao = New clsConexaoMC
    gSMConexao.Login frmPrincipal
    frmPrincipal.Visible = True
    
    mblnLoginRealizado = gSMConexao.LoginRealizado
    
    If mblnLoginRealizado = True Then
        frmPrincipal.Visible = False
        Call chamamenu
    Else
        End
    End If
End Sub

Public Sub chamamenu()
    frmMenu.Show vbModeless, Me
End Sub

