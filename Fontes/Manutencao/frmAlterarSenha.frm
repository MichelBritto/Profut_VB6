VERSION 5.00
Begin VB.Form frmAlterarSenha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProFut - Alterar Senha"
   ClientHeight    =   2685
   ClientLeft      =   9240
   ClientTop       =   5235
   ClientWidth     =   3690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3690
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPrincipal 
      Height          =   2715
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   3615
      Begin VB.CheckBox chkMostrarSenha 
         Caption         =   "Ver Senha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1245
      End
      Begin VB.CommandButton cmdAlterar 
         Appearance      =   0  'Flat
         Caption         =   "Confirmar"
         Height          =   405
         Left            =   1500
         Picture         =   "frmAlterarSenha.frx":0000
         TabIndex        =   8
         Top             =   2130
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Appearance      =   0  'Flat
         Caption         =   "Cancelar"
         Height          =   405
         Left            =   2550
         TabIndex        =   7
         Top             =   2130
         Width           =   975
      End
      Begin VB.TextBox txtConfirmarSenha 
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   60
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1590
         Width           =   3465
      End
      Begin VB.TextBox txtSenhaAtual 
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   90
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   330
         Width           =   3435
      End
      Begin VB.TextBox txtNovaSenha 
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   60
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   960
         Width           =   3465
      End
      Begin VB.Label Label1 
         Caption         =   "Confirmar nova senha"
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Top             =   1380
         Width           =   2745
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label 
         Caption         =   "Senha atual"
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   2145
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Nova senha"
         Height          =   285
         Left            =   90
         TabIndex        =   3
         Top             =   750
         Width           =   2745
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmAlterarSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstrLogin As String

Public Property Let Login(strLogin As String)
10        mstrLogin = strLogin
End Property

Private Function VerificarCampos() As Boolean
10    On Error GoTo Erro
            
          Dim blnContinua As Boolean
          Dim strMensagem As String
          
          
20        blnContinua = True
30        strMensagem = "Não foi possível alterar a senha pois existem as seguintes pendências:" & vbCrLf
            
40        If Len(txtNovaSenha.Text) < 6 Then
50            strMensagem = strMensagem & vbCrLf & "-> A senha precisa conter pelo menos 6 caracteres."
60            blnContinua = False
70        End If
          
80        If SomenteOsNumerosDaString(txtNovaSenha.Text) = "" Then
90            strMensagem = strMensagem & vbCrLf & "-> A senha precisa conter pelo menos um carater numérico."
100           blnContinua = False
110       End If
          
120       If txtNovaSenha.Text <> txtConfirmarSenha.Text Then
130           strMensagem = strMensagem & vbCrLf & "-> A senha digitada não é a mesma que a senha confirmada."
140           blnContinua = False
150       End If
          
160       If Not blnContinua Then
170           MsgBox strMensagem, vbOKOnly + vbExclamation, "Atenção"
180            VerificarCampos = False
190       Else
200            VerificarCampos = True
210       End If
          

220   Exit Function
Erro:
230      Call MsgBox("Erro no módulo: " & "frmAlterarSenha" & vbCrLf & "VerificarCampos" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")


End Function

Private Sub chkMostrarSenha_Click()
10    On Error GoTo Erro
            
20        If chkMostrarSenha.Value = vbChecked Then
30            txtConfirmarSenha.PasswordChar = ""
40            txtSenhaAtual.PasswordChar = ""
50            txtNovaSenha.PasswordChar = ""
60        Else
70            txtConfirmarSenha.PasswordChar = "*"
80            txtSenhaAtual.PasswordChar = "*"
90            txtNovaSenha.PasswordChar = "*"
100       End If

110   Exit Sub
Erro:
120      Call MsgBox("Erro no módulo: " & "frmAlterarSenha" & vbCrLf & "chkMostrarSenha_Click" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")


End Sub

Private Sub cmdAlterar_Click()
10    On Error GoTo Erro
          Dim blnResultado As Boolean
            
20        If VerificarCampos Then
30            Call modManutencao_AlterarSenhaPorLoginESenha(mstrLogin, txtSenhaAtual.Text, txtNovaSenha.Text, blnResultado)
              
40            If blnResultado = True Then
50                MsgBox "Senha alterada com sucesso!", vbOKOnly + vbInformation, "Atenção!"
60                Unload Me
70            Else
80                MsgBox "Senha não alterada!" & vbCrLf & "A senha digitada não coincide com a senha atual!", vbOKOnly + vbExclamation, "Atenção!"
90                txtSenhaAtual.Text = ""
100               txtNovaSenha.Text = ""
110               txtConfirmarSenha.Text = ""
120           End If

130       End If

140   Exit Sub
Erro:
150      Call MsgBox("Erro no módulo: " & "frmAlterarSenha" & vbCrLf & "cmdAlterar_Click" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")


End Sub

Private Sub cmdCancelar_Click()
10        Unload Me
End Sub
