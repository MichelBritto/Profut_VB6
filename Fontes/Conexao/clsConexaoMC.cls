VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConexaoMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Conexao() As ADODB.Connection
    If EstadoConexaoBD <= 0 Then
        Call conectar
    End If
    Set Conexao = gobjConn
End Property
Public Property Get TransacaoAberta() As Boolean
    TransacaoAberta = gblnTransacaoAberta
End Property

Public Property Get Empresa() As Long
    Empresa = glngEmpresa
End Property

Public Property Get LoginUsuario() As String
    LoginUsuario = gstrLoginUsuario
End Property

Public Property Get NomeUsuario() As String
    NomeUsuario = gstrNomeUsuario
End Property

Public Property Get EmailUsuario() As String
    EmailUsuario = gstrEmailUsuario
End Property

Public Property Get SenhaUsuario() As String
    SenhaUsuario = gstrSenhaUsuario
End Property

Public Property Get NomeBaseDados() As String
    NomeBaseDados = gstrNomeBaseDados
End Property

Public Property Get NomeServidor() As String
    NomeServidor = gstrNomeServidor
End Property

Public Property Get CodigoUsuario() As Integer
    CodigoUsuario = glngUsuario
End Property

'Permite conhecer o modo do estado da conexão do banco de dados
Public Property Get EstadoConexaoBD() As Long
On Error Resume Next
    EstadoConexaoBD = gobjConn.State
End Property
Public Property Get LoginRealizado() As Long
On Error Resume Next
    LoginRealizado = gblnLoginRealizado
End Property

Public Sub Login(Optional ByRef frmPai As Object = Nothing)
10    On Error GoTo erro

20        If frmPai Is Nothing Then
30            frmAutenticacao.Show vbModal
40        Else
50            frmAutenticacao.Move frmPai.Left + (frmPai.Width - frmAutenticacao.Width) - 300, frmPai.Top + (frmPai.Height - frmAutenticacao.Height) - 800
60            frmAutenticacao.Show vbModal, frmPai
70        End If

80    Exit Sub
erro:
90        Select Case Err.Number
              Case 373
100               frmAutenticacao.Show vbModeless
110               Resume Next
                  
120           Case 401
130               frmAutenticacao.Show vbModal
140               Resume Next
                  
150           Case Else
170       End Select
End Sub

Private Sub Class_Initialize()
   'Cria objeto connection
  If gobjConn Is Nothing Then
      Set gobjConn = New Connection
      Set gobjConn = CreateObject("ADODB.Connection")
  End If
End Sub

 
Public Sub BeginTransaction()
10    On Error GoTo erro
20        If Not gblnTransacaoAberta Then
30            gobjConn.BeginTrans
40            gblnTransacaoAberta = True
50        End If
60    Exit Sub
erro:
70        MsgBox "Ocorreu o erro no procedimento BeginTransaction." & Chr(13) & _
                        "Caso o erro persista, informe os seguintes dados ao administrador da rede:" & Chr(13) & _
                        "Nome do projeto: modConexao." & Chr(13) & _
                        "Tipo do módulo: Módulo de classe." & Chr(13) & _
                        "Nome do módulo: clsConexao." & Chr(13) & _
                        "Tipo de procedimento: Sub." & Chr(13) & _
                        "Nome do procedimento: BeginTransaction." & Chr(13) & _
                        "Numero do erro: " & Err.Number & "." & Chr(13) & _
                        "Descrição do erro: " & Err.Description & "." & Chr(13) & _
                        "Linha do erro: " & Erl & ".", vbOKOnly + vbCritical, "ERRO!"
End Sub

Public Sub CommitTransaction()
10    On Error GoTo erro
          
20        If gblnTransacaoAberta Then
30            gobjConn.CommitTrans
40            gblnTransacaoAberta = False
50        End If
60    Exit Sub
erro:
70        MsgBox "Ocorreu o erro no procedimento CommitTransaction." & Chr(13) & _
                        "Caso o erro persista, informe os seguintes dados ao administrador da rede:" & Chr(13) & _
                        "Nome do projeto: modConexao." & Chr(13) & _
                        "Tipo do módulo: Módulo de classe." & Chr(13) & _
                        "Nome do módulo: clsConexao." & Chr(13) & _
                        "Tipo de procedimento: Sub." & Chr(13) & _
                        "Nome do procedimento: CommitTransaction." & Chr(13) & _
                        "Numero do erro: " & Err.Number & "." & Chr(13) & _
                        "Descrição do erro: " & Err.Description & "." & Chr(13) & _
                        "Linha do erro: " & Erl & ".", vbOKOnly + vbCritical, "ERRO!"
End Sub

Public Sub RollbackTransaction()
10    On Error GoTo erro
          
20        If gblnTransacaoAberta Then
30            gobjConn.RollbackTrans
40            gblnTransacaoAberta = False
50        End If
60    Exit Sub
erro:
70        MsgBox "Ocorreu o erro no procedimento RollbackTransaction." & Chr(13) & _
                        "Caso o erro persista, informe os seguintes dados ao administrador da rede:" & Chr(13) & _
                        "Nome do projeto: modConexao." & Chr(13) & _
                        "Tipo do módulo: Módulo de classe." & Chr(13) & _
                        "Nome do módulo: clsConexao." & Chr(13) & _
                        "Tipo de procedimento: Sub." & Chr(13) & _
                        "Nome do procedimento: RollbackTransaction." & Chr(13) & _
                        "Numero do erro: " & Err.Number & "." & Chr(13) & _
                        "Descrição do erro: " & Err.Description & "." & Chr(13) & _
                        "Linha do erro: " & Erl & ".", vbOKOnly + vbCritical, "ERRO!"

End Sub

Public Sub conectar()
On Error GoTo erro
    With gobjConn
        .Provider = "Sqloledb"
        .CursorLocation = adUseClient
        .ConnectionString = "Server=" & gstrNomeServidor & ";Database=" & gstrNomeBaseDados & " ;User Id=" & gstrLoginUsuario & ";Password=" & gstrSenhaUsuario & ";"
        .Open
    End With
    gblnLoginRealizado = True
    Exit Sub
erro:
    'MsgBox "Falha ao realizar a conexão. Verifique o status da rede e tente novamente.", vbOKOnly + vbCritical, "Erro de conexão"
    gblnLoginRealizado = False
End Sub

Public Sub Desconectar()
    gobjConn.Close
End Sub




