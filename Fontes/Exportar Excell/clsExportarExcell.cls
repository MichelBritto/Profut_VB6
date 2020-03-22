VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExportarExcell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mlngColunasAlteradas()          As Long
Private mlngQuantidadeColunasAlteradas  As Long


Public Sub ExportarGrid(ssgGrid As Object)
    Dim strPath             As String
    
    If ssgGrid Is Nothing Then Exit Sub
    
    If frmExportarGrid.ProcurarArquivoTexto(strPath) Then
        TrocarCheckBoxParaTexto ssgGrid 'Isso evita ficar aparecendo as caixas de MARCAR que trava o excell
        On Error Resume Next
        Call ssgGrid.ExportToFile(strPath, False) 'esse false é APPEND -- logo sobrescreve o arquivo
    End If
End Sub

Private Sub TrocarCheckBoxParaTexto(objGrid As TDBGrid)
On Error Resume Next
    Dim k           As Integer
    Dim a           As New ValueItem
    Dim b           As New ValueItem
    Dim c           As New ValueItem
    
    
    a.Value = "1"
    a.DisplayValue = "X"
    
    b.Value = "0"
    b.DisplayValue = ""
    
    c.Value = "-1"
    c.DisplayValue = "X"
    
    mlngQuantidadeColunasAlteradas = 0
    For k = 0 To objGrid.Columns.Count - 1
        If objGrid.Columns(k).ValueItems.Presentation = dbgCheckBox And objGrid.Columns(k).Visible = True Then
            objGrid.Columns(k).ValueItems.Presentation = dbgNormal
            objGrid.Columns(k).ValueItems.Translate = True
            objGrid.Columns(k).ValueItems.Add a
            objGrid.Columns(k).ValueItems.Add b
            objGrid.Columns(k).ValueItems.Add c
            
            If mlngQuantidadeColunasAlteradas = 0 Then
                ReDim mlngColunasAlteradas(0)
            Else
                ReDim Preserve mlngColunasAlteradas(UBound(mlngColunasAlteradas) + 1)
            End If
            
            mlngColunasAlteradas(UBound(mlngColunasAlteradas)) = k
            mlngQuantidadeColunasAlteradas = mlngQuantidadeColunasAlteradas + 1
        End If
    Next
On Error GoTo 0
End Sub


Private Sub TrocarTextoParaCheckBox(objGrid As TDBGrid)
    Dim j           As Integer
    Dim k           As Integer
    
    If mlngQuantidadeColunasAlteradas > 0 Then
        For k = 0 To mlngQuantidadeColunasAlteradas - 1
            objGrid.Columns(mlngColunasAlteradas(k)).ValueItems.Presentation = dbgCheckBox
            
            For j = 0 To objGrid.Columns(mlngColunasAlteradas(k)).ValueItems.Count - 1
                objGrid.Columns(mlngColunasAlteradas(k)).ValueItems.Remove 0
            Next
        Next
    End If
End Sub


