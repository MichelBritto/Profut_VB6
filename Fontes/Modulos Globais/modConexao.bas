Attribute VB_Name = "mdConexao"
Option Explicit

'Vari�veis de banco de dADOs
Public gstrSQL                   As String     'Consultas sql
Public gobjConn                  As Connection 'Objeto ADO Connection
Public glngUsuario               As Long
Public gstrLoginUsuario          As String

Public gblnLoginRealizado        As Boolean

'Vari�veis expostas via propriedades
Public gstrNomeServidor          As String     'Nome do servidor
Public gstrNomeUsuario           As String
Public gstrEmailUsuario          As String
Public gstrSenhaUsuario          As String
Public gstrNomeBaseDados         As String

Public gstrCaminhoArquivoErros   As String

'Vari�veis de uso geral n�o expostas via propriedades
Public gstrMsg                   As String     'Usado em eventos externos

'Constantes de banco de dados
Public Const adStateClosed = 0
Public Const adStateOpen = 1
Public Const adStateConnecting = 2
Public Const adStateExecuting = 4
Public Const adStateFetching = 8

Public gblnTransacaoAberta      As Boolean




