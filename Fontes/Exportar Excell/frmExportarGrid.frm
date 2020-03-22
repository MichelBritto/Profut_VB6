VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Begin VB.Form frmExportarGrid 
   ClientHeight    =   630
   ClientLeft      =   7665
   ClientTop       =   2130
   ClientWidth     =   1740
   LinkTopic       =   "Form1"
   ScaleHeight     =   630
   ScaleWidth      =   1740
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmExportarGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Function ProcurarArquivoTexto(ByRef strPath As String) As Boolean
On Error GoTo Erro 'para ver se clicou em cancelar a única maneira é colocando a propriedade cancelerror = true e testando o número do erro

    cmnDlg.CancelError = False 'True
    'cmnDlg.InitDir = strPath
    cmnDlg.Flags = cdlOFNOverwritePrompt 'faz ficar center screen
    
    'cmnDlg.Filter = "Bancos (*.bb)(*.txt)(*.doc)|*.bb||*.txt||*.doc|" 'txt e doc são apenas para efeito de testes, a extensão de arquivos de retorno acho que é .ret
    cmnDlg.Filter = "Planilha Eletrônica (*.xls)|*.xls|"
    'Else
 '   cmnDlg.Filter = ""
    'End If
    'cmnDlg.DefaultExt
    cmnDlg.ShowSave
        
    'If Dir(cmnDlg.FileName, vbArchive) <> "" Then
    If cmnDlg.FileName <> "" Then
        strPath = cmnDlg.FileName
        ProcurarArquivoTexto = True
        'txtArquivo.Text = Dir(cmnDlg.FileName, vbArchive)
        'txtArquivo.Tag = cmnDlg.FileName
    End If
Exit Function
Erro:
    If Err.Number = 32755 Then 'clicou em cancelar
        strPath = ""
    Else
        MsgBox "Ocorreu um erro na rotina de abrir arquivo texto" & Chr(13) & _
                "Erro número: " & Err.Number & Chr(13) & _
                "Erro Descrição: " & Err.Description
    End If
    ProcurarArquivoTexto = False
End Function

