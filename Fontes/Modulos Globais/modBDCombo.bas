Attribute VB_Name = "modBDCombo"
Option Explicit

Public Sub modBDCombo_SelecionarCartecoriaJogador(ByRef objSSOleDBCombo As SSOleDBCombo, Optional ByVal lngCodigo As Long)

    Dim intIndex          As Integer
    Dim objrs             As New Recordset
   
On Error GoTo Erro
   
    intIndex = -1
    With objSSOleDBCombo
        .Columns.RemoveAll
        
        .Columns.Add (0)
        .Columns(0).Name = "chdescricao"
        .Columns(0).Caption = "Descricao"
        .Columns(0).Width = objSSOleDBCombo.Width
        .Columns(0).Visible = True
        
        .Columns.Add (1)
        .Columns(1).Name = "chcodigo"
        .Columns(1).Caption = "codigo"
        .Columns(1).Visible = False
        
        .DataFieldToDisplay = "column 0"
    End With
     
    objrs.Open "dbo.USP_SELECIONARCARTEGORIA", gSMConexao.Conexao, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    With objrs
        If Not objrs.RecordCount = 0 Then objrs.MoveFirst
        objSSOleDBCombo.RemoveAll
        Do While Not .EOF
            objSSOleDBCombo.AddItem NZ(!Descricao_VC) & vbTab & NS(!ID_IN)
            If NZ(!ID_IN) = lngCodigo Then intIndex = objrs.AbsolutePosition - 1
            .MoveNext
        Loop
    End With
    
    If intIndex = -1 Then
        objSSOleDBCombo.Text = ""
    Else
        objSSOleDBCombo.Bookmark = IIf(lngCodigo = 0, -1, objSSOleDBCombo.AddItemBookmark(intIndex))
        objSSOleDBCombo.Text = objSSOleDBCombo.Columns("chDescricao").CellValue(objSSOleDBCombo.Bookmark)
    End If

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "modBDCombo" & vbCrLf & "No Procedimento: " & "modBDCombo_SelecionarCartecoriaJogador" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Public Sub modBDCombo_SelecionarEstados(ByRef objSSOleDBCombo As SSOleDBCombo, Optional ByVal lngCodigo As Long)

    Dim intIndex          As Integer
    Dim objrs             As New Recordset
   
On Error GoTo Erro
   
    intIndex = -1
    With objSSOleDBCombo
        .Columns.RemoveAll
        
        .Columns.Add (0)
        .Columns(0).Name = "chdescricao"
        .Columns(0).Caption = "Descricao"
        .Columns(0).Width = objSSOleDBCombo.Width
        .Columns(0).Visible = True
        
        .Columns.Add (1)
        .Columns(1).Name = "chcodigo"
        .Columns(1).Caption = "codigo"
        .Columns(1).Visible = False
        
        .DataFieldToDisplay = "column 0"
    End With
     
    objrs.Open "dbo.USP_SELECIONARESTADOS", gSMConexao.Conexao, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    With objrs
        If Not objrs.RecordCount = 0 Then objrs.MoveFirst
        objSSOleDBCombo.RemoveAll
        Do While Not .EOF
            objSSOleDBCombo.AddItem NZ(!UF_CH) & vbTab & NS(!ID_IN)
            If NZ(!ID_IN) = lngCodigo Then intIndex = objrs.AbsolutePosition - 1
            .MoveNext
        Loop
    End With
    
    If intIndex = -1 Then
        objSSOleDBCombo.Text = ""
    Else
        objSSOleDBCombo.Bookmark = IIf(lngCodigo = 0, -1, objSSOleDBCombo.AddItemBookmark(intIndex))
        objSSOleDBCombo.Text = objSSOleDBCombo.Columns("chDescricao").CellValue(objSSOleDBCombo.Bookmark)
    End If

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "modBDCombo" & vbCrLf & "No Procedimento: " & "modBDCombo_SelecionarEstados" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Public Sub modBDCombo_SelecionarEquipePorCodigo(ByRef objSSOleDBCombo As SSOleDBCombo, Optional ByVal lngCodigo As Long)

    Dim intIndex          As Integer
    Dim objrs             As New Recordset
   
On Error GoTo Erro
   
    intIndex = -1
    With objSSOleDBCombo
        .Columns.RemoveAll
        
        .Columns.Add (0)
        .Columns(0).Name = "chdescricao"
        .Columns(0).Caption = "Descricao"
        .Columns(0).Width = objSSOleDBCombo.Width
        .Columns(0).Visible = True
        
        .Columns.Add (1)
        .Columns(1).Name = "chcodigo"
        .Columns(1).Caption = "codigo"
        .Columns(1).Visible = False
        
        .DataFieldToDisplay = "column 0"
    End With
     
    objrs.Open "dbo.USP_SELECIONAREQUIPEPORCODIGO", gSMConexao.Conexao, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    With objrs
        If Not objrs.RecordCount = 0 Then objrs.MoveFirst
        objSSOleDBCombo.RemoveAll
        Do While Not .EOF
            objSSOleDBCombo.AddItem NZ(!Nome_VC) & vbTab & NS(!ID_IN)
            If NZ(!ID_IN) = lngCodigo Then intIndex = objrs.AbsolutePosition - 1
            .MoveNext
        Loop
    End With
    
    If intIndex = -1 Then
        objSSOleDBCombo.Text = ""
    Else
        objSSOleDBCombo.Bookmark = IIf(lngCodigo = 0, -1, objSSOleDBCombo.AddItemBookmark(intIndex))
        objSSOleDBCombo.Text = objSSOleDBCombo.Columns("chDescricao").CellValue(objSSOleDBCombo.Bookmark)
    End If

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "modBDCombo" & vbCrLf & "No Procedimento: " & "modBDCombo_SelecionarEquipePorCodigo" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Public Sub modBDCombo_SelecionarCargos(ByRef objSSOleDBCombo As SSOleDBCombo, Optional ByVal lngCodigo As Long)

    Dim intIndex          As Integer
    Dim objrs             As New Recordset
   
On Error GoTo Erro
   
    intIndex = -1
    With objSSOleDBCombo
        .Columns.RemoveAll
        
        .Columns.Add (0)
        .Columns(0).Name = "chdescricao"
        .Columns(0).Caption = "Descricao"
        .Columns(0).Width = objSSOleDBCombo.Width
        .Columns(0).Visible = True
        
        .Columns.Add (1)
        .Columns(1).Name = "chcodigo"
        .Columns(1).Caption = "codigo"
        .Columns(1).Visible = False
        
        .DataFieldToDisplay = "column 0"
    End With
     
    objrs.Open "dbo.usp_SelecionarCargos", gSMConexao.Conexao, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    With objrs
        If Not objrs.RecordCount = 0 Then objrs.MoveFirst
        objSSOleDBCombo.RemoveAll
        Do While Not .EOF
            objSSOleDBCombo.AddItem NS(!Cargo_VC) & vbTab & NZ(!Cargo_IN)
            If NZ(!Cargo_IN) = lngCodigo Then intIndex = objrs.AbsolutePosition - 1
            .MoveNext
        Loop
    End With
    
    If intIndex = -1 Then
        objSSOleDBCombo.Text = ""
    Else
        objSSOleDBCombo.Bookmark = IIf(lngCodigo = 0, -1, objSSOleDBCombo.AddItemBookmark(intIndex))
        objSSOleDBCombo.Text = objSSOleDBCombo.Columns("chDescricao").CellValue(objSSOleDBCombo.Bookmark)
    End If

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "modBDCombo" & vbCrLf & "modBDCombo_SelecionarCargos" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Public Sub modBDCombo_SelecionarCidades(ByRef objSSOleDBCombo As SSOleDBCombo, Optional ByVal lngCodigoCidade As Long, Optional ByVal lngCodigoUF As Long)

    Dim intIndex          As Integer
    Dim objrs             As New Recordset

On Error GoTo Erro

    intIndex = -1
    With objSSOleDBCombo
        .Columns.RemoveAll

        .Columns.Add (0)
        .Columns(0).Name = "chdescricao"
        .Columns(0).Caption = "Descricao"
        .Columns(0).Width = objSSOleDBCombo.Width
        .Columns(0).Visible = True

        .Columns.Add (1)
        .Columns(1).Name = "chcodigoCidade"
        .Columns(1).Caption = "Cidade"
        .Columns(1).Visible = False

        .Columns.Add (2)
        .Columns(2).Name = "chcodigoUF"
        .Columns(2).Caption = "Estado"
        .Columns(2).Visible = False

        .DataFieldToDisplay = "column 0"
    End With

    objrs.Open "dbo.usp_SelecionarCidades", gSMConexao.Conexao, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    
    If lngCodigoUF <> 0 Then
        objrs.Filter = "Estado_IN=" & lngCodigoUF
    End If
    
    With objrs
        If Not objrs.RecordCount = 0 Then objrs.MoveFirst
        objSSOleDBCombo.RemoveAll
        Do While Not .EOF
            objSSOleDBCombo.AddItem NZ(!Nome_VC) & vbTab & NS(!Cidade_IN) & vbTab & NS(!Estado_IN)
            If NZ(!ID_IN) = lngCodigoCidade Then intIndex = objrs.AbsolutePosition - 1
            .MoveNext
        Loop
    End With

    If intIndex = -1 Then
        objSSOleDBCombo.Text = ""
    Else
        objSSOleDBCombo.Bookmark = IIf(lngCodigoCidade = 0, -1, objSSOleDBCombo.AddItemBookmark(intIndex))
        objSSOleDBCombo.Text = objSSOleDBCombo.Columns("chDescricao").CellValue(objSSOleDBCombo.Bookmark)
    End If

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "modBDCombo" & vbCrLf & "No Procedimento: " & "modBDCombo_SelecionarEstados" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub
