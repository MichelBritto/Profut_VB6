VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProFut - Menu Principal"
   ClientHeight    =   6330
   ClientLeft      =   4515
   ClientTop       =   2340
   ClientWidth     =   10860
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   10860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPrincipal 
      BackColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   10785
      Begin VB.PictureBox logo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   4560
         Picture         =   "frmMenu.frx":500A
         ScaleHeight     =   5175
         ScaleWidth      =   4995
         TabIndex        =   11
         Top             =   480
         Width           =   4995
      End
      Begin VB.PictureBox pcMM 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1920
         Picture         =   "frmMenu.frx":8F06
         ScaleHeight     =   405
         ScaleWidth      =   2295
         TabIndex        =   10
         Top             =   6900
         Width           =   2295
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
         Begin Threed.SSCommand cmdMenuCompeticao 
            Height          =   645
            Left            =   30
            TabIndex        =   13
            Top             =   780
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1138
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
            Picture         =   "frmMenu.frx":9938
            Caption         =   "Menu Competição"
            Alignment       =   1
            PictureAlignment=   9
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdCadastroEquipe 
            Height          =   645
            Left            =   30
            TabIndex        =   3
            Top             =   780
            Visible         =   0   'False
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1138
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
            Picture         =   "frmMenu.frx":9CEA
            Caption         =   "Clube/Equipe"
            Alignment       =   1
            PictureAlignment=   9
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdRelatorioJogador 
            Height          =   645
            Left            =   3270
            TabIndex        =   4
            Top             =   1530
            Visible         =   0   'False
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1138
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
            Picture         =   "frmMenu.frx":A09C
            Caption         =   "Lista de Jogadores"
            Alignment       =   1
            PictureAlignment=   9
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSair 
            Height          =   1035
            Left            =   30
            TabIndex        =   5
            Top             =   5130
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1826
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
            Picture         =   "frmMenu.frx":A44E
            Caption         =   "Sair"
            Alignment       =   1
            PictureAlignment=   9
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdUsuario 
            Height          =   645
            Left            =   3270
            TabIndex        =   6
            Top             =   2250
            Visible         =   0   'False
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1138
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
            Picture         =   "frmMenu.frx":A800
            Caption         =   "Usuário"
            Alignment       =   1
            PictureAlignment=   9
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdPermissao 
            Height          =   645
            Left            =   3270
            TabIndex        =   7
            Top             =   2970
            Visible         =   0   'False
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1138
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
            Picture         =   "frmMenu.frx":ABB2
            Caption         =   "Permissão"
            Alignment       =   1
            PictureAlignment=   9
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdCadCargo 
            Height          =   645
            Left            =   3270
            TabIndex        =   8
            Top             =   3690
            Visible         =   0   'False
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1138
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
            Picture         =   "frmMenu.frx":AF64
            Caption         =   "Cadastro de Cargo"
            Alignment       =   1
            PictureAlignment=   9
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdCadPosicao 
            Height          =   645
            Left            =   3270
            TabIndex        =   9
            Top             =   4410
            Visible         =   0   'False
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1138
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
            Picture         =   "frmMenu.frx":B316
            Caption         =   "Cadastro de Posição"
            Alignment       =   1
            PictureAlignment=   9
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdModuloCadastro 
            Height          =   645
            Left            =   30
            TabIndex        =   12
            Top             =   60
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1138
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
            Picture         =   "frmMenu.frx":B6C8
            Caption         =   "Menu Cadastral"
            Alignment       =   1
            PictureAlignment=   9
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdCadastroDeJogador 
            Height          =   645
            Left            =   30
            TabIndex        =   2
            Top             =   60
            Visible         =   0   'False
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1138
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
            Picture         =   "frmMenu.frx":BA7A
            Caption         =   "Ficha de Jogador"
            Alignment       =   1
            PictureAlignment=   9
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdVoltar 
            Height          =   1035
            Left            =   3270
            TabIndex        =   14
            Top             =   5130
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1826
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
            Picture         =   "frmMenu.frx":BE2C
            Caption         =   "Voltar ao menu"
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
            
            If RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 1) Then
                objJogador.Show gSMConexao, "ProFut - Cadastro de Jogador"
                Exit Sub
            Else
                MsgBox "Acesso negado!" & vbCrLf & "-> Usuário não possuí a permissão Nº 1", vbOKOnly + vbExclamation, "Atenção!"
            End If
        Else
           gSMConexao.conectar
        End If
    End If
End Sub


Private Sub cmdCadastroEquipe_Click()
    Dim objEquipe As clsCadEquipe
    Set objEquipe = New clsCadEquipe
    
On Error GoTo Erro
          
    If Not gSMConexao Is Nothing Then
        If gSMConexao.EstadoConexaoBD = adStateOpen Then
            If RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 5) Then
                objEquipe.Show gSMConexao, "Profut - Cadastro de Equipe"
                Exit Sub
            Else
                MsgBox "Acesso negado!" & vbCrLf & "->Usuário não tem a permissão Nº5", vbOKOnly + vbExclamation, "Atenção!"
            End If
        Else
           gSMConexao.conectar
        End If
    End If
    
Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmMenu" & vbCrLf & "cmdCadastroEquipe_Click" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
    
    
End Sub


Private Sub cmdCadastroEquipe_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H000080FF&"
    cmdCadastroEquipe.ForeColor = &H80FF&
End Sub

Private Sub cmdCadastroEquipe_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H00000000&"
    cmdCadastroEquipe.ForeColor = vbWhite
End Sub


Private Sub cmdManutencao_Click()
    If cmdUsuario.Visible = False Then
        cmdUsuario.Visible = True
    Else
        cmdUsuario.Visible = False
    End If
    
    If cmdPermissao.Visible = False Then
        cmdPermissao.Visible = True
    Else
        cmdPermissao.Visible = False
    End If
End Sub

Private Sub cmdManutencao_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H000080FF&"
    cmdManutencao.ForeColor = &H80FF&
End Sub

Private Sub cmdManutencao_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H00000000&"
    cmdManutencao.ForeColor = vbWhite
End Sub

Private Sub cmdCadCargo_Click()
On Error GoTo Erro
    Dim objMan As clsManutencao
    Set objMan = New clsManutencao
    
    If RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 11) Then
        objMan.ShowCargos gSMConexao, "ProFut - Cadastro de Cargos", vbModeless, Me
    Else
        MsgBox "Acesso negado!" & vbCrLf & "->Usuário não tem a permissão Nº11", vbOKOnly + vbExclamation, "Atenção!"
    End If
    
    Exit Sub
Erro:
       Call MsgBox("Erro no módulo: " & "frmMenu" & vbCrLf & "cmdCadCargo_Click" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub
Private Sub cmdCadCargo_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H000080FF&"
    cmdCadCargo.ForeColor = &H80FF&
End Sub

Private Sub cmdCadCargo_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H00000000&"
    cmdCadCargo.ForeColor = vbWhite
End Sub

Private Sub cmdCadPosicao_Click()
On Error GoTo Erro
   
    Dim objMan As clsManutencao
    Set objMan = New clsManutencao
    
    If Not gSMConexao Is Nothing Then
        If gSMConexao.EstadoConexaoBD = adStateOpen Then
    
            If RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 15) Then
                objMan.ShowPosicao gSMConexao, "ProFut - Cadastro de Posição", vbModeless, Me
            Else
                MsgBox "Acesso negado!" & vbCrLf & "->Usuário não tem a permissão Nº15", vbOKOnly + vbExclamation, "Atenção!"
            End If
        End If
    Else
        gSMConexao.conectar
    End If

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmMenu" & vbCrLf & "cmdCadPosicao_Click" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")


End Sub
Private Sub cmdCadPosicao_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H000080FF&"
    cmdCadPosicao.ForeColor = &H80FF&
End Sub

Private Sub cmdCadPosicao_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H00000000&"
    cmdCadPosicao.ForeColor = vbWhite
End Sub


Private Sub cmdMenuCompeticao_Click()
    MsgBox "Módulo ainda não implementado, aguarde novidades!", vbOKOnly + vbInformation, "Atenção!"
End Sub
Private Sub cmdMenuCompeticao_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H000080FF&"
    cmdMenuCompeticao.ForeColor = &H80FF&
End Sub

Private Sub cmdMenuCompeticao_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H00000000&"
    cmdMenuCompeticao.ForeColor = vbWhite
End Sub

Private Sub cmdModuloCadastro_Click()
On Error GoTo Erro
    'Oculto os menús
    cmdModuloCadastro.Visible = False
    cmdMenuCompeticao.Visible = False
    cmdSair.Visible = False
    'Exibo os cadastros
    cmdCadastroDeJogador.Visible = True
    cmdCadastroEquipe.Visible = True
    cmdRelatorioJogador.Visible = True
    cmdPermissao.Visible = True
    cmdUsuario.Visible = True
    cmdCadCargo.Visible = True
    cmdCadPosicao.Visible = True
    cmdVoltar.Visible = True
    'Acerto o left
    cmdCadastroDeJogador.Left = cmdModuloCadastro.Left
    cmdCadastroEquipe.Left = cmdModuloCadastro.Left
    cmdRelatorioJogador.Left = cmdModuloCadastro.Left
    cmdPermissao.Left = cmdModuloCadastro.Left
    cmdUsuario.Left = cmdModuloCadastro.Left
    cmdCadCargo.Left = cmdModuloCadastro.Left
    cmdCadPosicao.Left = cmdModuloCadastro.Left
    cmdVoltar.Left = cmdModuloCadastro.Left
    
Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmMenu" & vbCrLf & "cmdModuloCadastro_Click" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Private Sub cmdModuloCadastro_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H000080FF&"
    cmdModuloCadastro.ForeColor = &H80FF&
End Sub

Private Sub cmdModuloCadastro_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H00000000&"
    cmdModuloCadastro.ForeColor = vbWhite
End Sub

Private Sub cmdRelatorioJogador_Click()
    Dim ObjRelatorioJogador As ClsRelJogador
    Set ObjRelatorioJogador = New ClsRelJogador
    
On Error GoTo Erro
          
    
    If Not gSMConexao Is Nothing Then
        If gSMConexao.EstadoConexaoBD = adStateOpen Then
            If RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 7) Then
                ObjRelatorioJogador.Show gSMConexao, "ProFut - Relatório de Jogador"
                Exit Sub
            Else
                MsgBox "Acesso negado!" & vbCrLf & "->Usuário não tem a permissão Nº7", vbOKOnly + vbExclamation, "Atenção!"
            End If
        Else
            gSMConexao.conectar
        End If
    End If
    
Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmMenu" & vbCrLf & "cmdRelatorioJogador_Click" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
    
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

Private Sub cmdUsuario_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H000080FF&"
    cmdUsuario.ForeColor = &H80FF&
End Sub

Private Sub cmdUsuario_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H00000000&"
    cmdUsuario.ForeColor = vbWhite
End Sub
Private Sub cmdUsuario_Click()
    Dim ObjManutencao As clsManutencao
    Set ObjManutencao = New clsManutencao
On Error GoTo Erro
          
    
    If Not gSMConexao Is Nothing Then
    
        If gSMConexao.EstadoConexaoBD = adStateOpen Then
            If RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 8) Then
                ObjManutencao.ShowUsuarios gSMConexao, "ProFut - Cadastro de Usuário"
                Exit Sub
            Else
                MsgBox "Acesso negado!" & vbCrLf & "->Usuário não tem a permissão Nº8", vbOKOnly + vbExclamation, "Atenção!"
            End If
        Else
            gSMConexao.conectar
        End If
    End If

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmMenu" & vbCrLf & "cmdUsuario_Click" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")


End Sub

Private Sub cmdVoltar_Click()
On Error GoTo Erro
    
    'Exibo os menús
    cmdModuloCadastro.Visible = True
    cmdMenuCompeticao.Visible = True
    cmdSair.Visible = True
    'oculto os cadastros
    cmdCadastroDeJogador.Visible = False
    cmdCadastroEquipe.Visible = False
    cmdRelatorioJogador.Visible = False
    cmdPermissao.Visible = False
    cmdUsuario.Visible = False
    cmdCadCargo.Visible = False
    cmdCadPosicao.Visible = False
    cmdVoltar.Visible = False
        
    Exit Sub
Erro:
       Call MsgBox("Erro no módulo: " & "frmMenu" & vbCrLf & "cmdVoltar_Click" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub
Private Sub cmdVoltar_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H000080FF&"
    cmdVoltar.ForeColor = &H80FF&
End Sub

Private Sub cmdVoltar_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H00000000&"
    cmdVoltar.ForeColor = vbWhite
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmPrincipal.FecharPrograma = True
End Sub

Private Sub cmdPermissao_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H000080FF&"
    cmdPermissao.ForeColor = &H80FF&
End Sub

Private Sub cmdPermissao_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'cmdCadastroDeJogador.BackColor = "&H00000000&"
    cmdPermissao.ForeColor = vbWhite
End Sub
Private Sub cmdPermissao_Click()
    Dim ObjManutencao As clsManutencao
    Set ObjManutencao = New clsManutencao
    
On Error GoTo Erro
    
    If Not gSMConexao Is Nothing Then
    
        If gSMConexao.EstadoConexaoBD = adStateOpen Then
            If RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 10) Then
                ObjManutencao.ShowPermissao gSMConexao, "ProFut - Cadastro de Usuário"
                Exit Sub
            Else
                MsgBox "Acesso negado!" & vbCrLf & "->Usuário não tem a permissão Nº10", vbOKOnly + vbExclamation, "Atenção!"
            End If
        Else
            gSMConexao.conectar
        End If
    End If
    
Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmMenu" & vbCrLf & "cmdPermissao_Click" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")


End Sub
