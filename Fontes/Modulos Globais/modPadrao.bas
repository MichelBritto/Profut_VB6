Attribute VB_Name = "modPadrao"
Option Explicit

''
'Utilizado nas funções
''
Public gobjCmd                         As New Command
''

Public gSMConexao                       As New clsConexaoMC
Public glngEmpresa                      As Long
Public gstrCaption                      As String
Public glngLOGPermissao                 As Long
Public glngTransacaoLOGPermissao        As Long

''
'Utilizado no módulo contábil
''
Public gdatDataCompetencia             As Date
''

Public Const VK_SHIFT = &H10
Public Const mConstButtonFace           As Variant = &H8000000F
Public Const mConstWrite                As Variant = &H80000005

Public Const gConstVermelhoTotais       As Variant = &H80&
Public Const gConstAzulTotais           As Variant = &H800000
Public Const gConstVerdeTotais          As Variant = &H8000&
Public Const gConstAmareloTotais        As Variant = &H8080&

Private Const conSwNormal = 1

Enum enuTipoCabecalhoRelatorio
    RELATORIO_INTERNO_RETRATO = 1
    RELATORIO_INTERNO_PAISAGEM = 2
    RELATORIO_EXTERNO_RETRATO = 3
    RELATORIO_EXTERNO_PAISAGEM = 4
End Enum

' EXECUTA O PROGRAMA ASSOCIADO A UM DOCUMENTO
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Usado para Salvar um arquivo passando o URL
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

'Pegar o nome do computador
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'Fechar form no evento LOAD
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_CLOSE = &H10
'Exemplo: PostMessage Me.hwnd, WM_CLOSE, 0, 0

'Modulo INI
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'Funçao SLEEP
Public Declare Sub Sleep2 Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

'Usado para Salvar um arquivo passando o URL
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

'#####################################################################'
'##        T R A T A M E N T O   D E   C A M P O S   N U L L        ##'
'#####################################################################'
Public Function NZ(ByVal valor As Variant, Optional ByVal lngValorPadrao As Long = 0) As Variant
10        If IsNull(valor) Then
20            NZ = lngValorPadrao
30        Else
40            NZ = valor
50            If CStr(valor) = Empty Then NZ = lngValorPadrao
60        End If
End Function

Public Function ND(ByVal valor As Variant, Optional ByVal datDataPadrao As Date = 0) As Variant
10        If IsNull(valor) Then
20            ND = datDataPadrao
30        Else
40            ND = valor
50            If CStr(valor) = Empty Then ND = datDataPadrao
60        End If
End Function

Public Function NS(ByVal valor As Variant) As Variant
10        NS = valor
20        If IsNull(valor) Then NS = ""
End Function

Public Function NB(ByVal Value As Variant) As Boolean

On Error GoTo Erro

    If IsNull(Value) Then
        Value = False
    End If
    
    Value = Trim(Value)

    If UCase(Value) = UCase("Verdadeiro") Then
        Value = True
    ElseIf UCase(Value) = UCase("Falso") Then
        Value = False
    ElseIf UCase(Value) = UCase("True") Then
        Value = True
    ElseIf UCase(Value) = UCase("False") Then
        Value = False
    ElseIf UCase(Value) = UCase("T") Then
        Value = True
    ElseIf UCase(Value) = UCase("V") Then
        Value = True
    ElseIf UCase(Value) = UCase("F") Then
        Value = False
    ElseIf UCase(Value) = UCase("S") Then
        Value = True
    ElseIf UCase(Value) = UCase("SIM") Then
        Value = True
    ElseIf UCase(Value) = UCase("N") Then
        Value = False
    ElseIf UCase(Value) = UCase("NAO") Then
        Value = False
    ElseIf UCase(Value) = UCase("NÃO") Then
        Value = False
    ElseIf Trim(Value) = "" Then
        Value = False
    ElseIf IsNumeric(Value) Then
        If Value = 1 Then
            Value = True
        ElseIf Value = -1 Then
            Value = True
        Else
            Value = False
        End If
    Else
        Value = CBool(Value)
    End If

Retorna:
    NB = CBool(Value)
Exit Function
Erro:
    Value = False
    GoTo Retorna
End Function
'#####################################################################'
'#####################################################################'
'
'
'Public Function IGlobal_RetornarDatadoServidor() As Date
'          Dim objRSData            As New Recordset
'10    On Error GoTo erro
'
'20        objRSData.Open "SELECT GETDATE()", gSMConexao.Conexao, adOpenForwardOnly
'
'30        IGlobal_RetornarDatadoServidor = Date
'40        If Not objRSData.EOF Then IGlobal_RetornarDatadoServidor = objRSData.Fields(0).Value
'50        objRSData.Close
'
'60    Exit Function
'erro:
'70        MsgBox "Erro ao selecionar a data do servidor", vbOKOnly + vbInformation, "modBDPadrao"
'End Function



Public Function DigitaNumero(ByVal KeyAscii As Integer, ByRef strTextoControle As String, _
                             Optional ByVal blnPermiteSomenteInteiro As Boolean = False) As Integer
                                 
    'Retorna o KeyAscii validado aceitando somente numeros
    Dim blnVirgulaEncontrada As Boolean
    
    If KeyAscii = 3 Then DigitaNumero = KeyAscii: Exit Function     'CTRL + C
    If KeyAscii = 22 Then DigitaNumero = KeyAscii: Exit Function      'CTRL + V
    If KeyAscii = vbKeyBack Then DigitaNumero = KeyAscii: Exit Function
    'If KeyAscii = vbKeyReturn Then DigitaNumero = KeyAscii: Exit Function
    

    blnVirgulaEncontrada = CBool(InStr(1, strTextoControle, ",", vbTextCompare))
    If blnPermiteSomenteInteiro And Chr(KeyAscii) = "," Then DigitaNumero = 0: Exit Function
    If (Chr(KeyAscii) = "," And blnVirgulaEncontrada) Then DigitaNumero = 0: Exit Function
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And Chr(KeyAscii) <> "," And KeyAscii <> vbKeyBack Then KeyAscii = 0
    DigitaNumero = KeyAscii
End Function


Public Function SomenteOsNumerosDaString(strTexto As String) As String
      Dim strParsed             As String
      Dim i                     As Long

      'Retorna somente a parte da string que é númerica, pode retornar vazio
10    strParsed = ""
20    For i = 1 To Len(strTexto)
30       If IsNumeric(Mid(strTexto, i, 1)) Then
40           strParsed = strParsed & Mid(strTexto, i, 1)
50       End If
60    Next
70    SomenteOsNumerosDaString = strParsed
End Function

Public Function FormataMoeda(strValor As String, Optional strFormato As String = Empty) As String
      Dim strParsed           As String
      Dim i                   As Long
          
          'Formata Moeda Validando o Texto do Controle extraindo apenas o valor numerico
10        strParsed = ""
20         For i = 1 To Len(strValor)
30            If IsNumeric(Mid(strValor, i, 1)) Then
40                strParsed = strParsed & Mid(strValor, i, 1)
50            Else
60                If Mid(strValor, i, 1) = "," Then
70                    strParsed = strParsed & Mid(strValor, i, 1)
80                End If
90            End If
100       Next
110       If strParsed = "" Or strParsed = "," Then strParsed = "0"
120       FormataMoeda = Format(strParsed, IIf(strFormato = Empty, "#,##0.00", strFormato))
End Function


Public Function FormataCPF(ByVal strCPF As String, Optional ByVal blnLimparSeForInvalido As Boolean = False) As String
10    On Error GoTo Erro
            

20        strCPF = SomenteOsNumerosDaString(strCPF)
30        If Len(strCPF) > 11 Then strCPF = Right(strCPF, 11)
          
40        strCPF = Replace(Replace(Replace(Format(strCPF, "000PT000PT000TRAC00"), "PT", ".", , , vbTextCompare), "BAR", "/", , , vbTextCompare), "TRAC", "-", , , vbTextCompare)
          
50        'If blnLimparSeForInvalido Then
60         '   If Not ValidarCPF(strCPF) Then
70          '      strCPF = ""
80           ' End If
90        'End If
100       FormataCPF = strCPF

110   Exit Function
Erro:
120       Call MsgBox("Erro na função FormataCPF" & Chr(13) & "Erro: " & Err.Description & Chr(13) & "Linha: " & Erl, vbOKOnly + vbCritical, "ERRO")

End Function

Public Function FormataTelefone(ByVal strTelefone As String, _
                                Optional ByVal blnAceitaCelular As Boolean = True) As String
          Dim strRetorno          As String
10    On Error GoTo Erro
            
20        strRetorno = Trim(SomenteOsNumerosDaString(strTelefone))
          
30        If Len(strRetorno) < 3 Then
40            strRetorno = ""
50        Else
60            If Len(strRetorno) > 8 And blnAceitaCelular Then
71                strRetorno = Right(strRetorno, 9)
70                strRetorno = Format(strRetorno, "# ####-####")
80            Else
81                strRetorno = Right(strRetorno, 8)
90                strRetorno = Format(strRetorno, "####-####")
100           End If
110       End If
          
120       FormataTelefone = strRetorno
          
130   Exit Function
Erro:
140       Call MsgBox("Erro na função FormataTelefone" & Chr(13) & "Erro: " & Err.Description & Chr(13) & "Linha: " & Erl, vbOKOnly + vbCritical, "ERRO")
End Function


Public Function ValidaTelefone(ByVal strTelefone As String, _
                                Optional ByVal blnAceitaCelular As Boolean = True) As Boolean
          Dim strRetorno          As String
10    On Error GoTo Erro
            
20        strRetorno = Trim(SomenteOsNumerosDaString(strTelefone))
30        ValidaTelefone = False
          
40        If Len(strRetorno) = 8 Then
50            ValidaTelefone = True
60        ElseIf Len(strRetorno) = 9 And blnAceitaCelular Then
70            ValidaTelefone = True
80        Else
90            ValidaTelefone = False
100       End If
          
110   Exit Function
Erro:
120       Call MsgBox("Erro na função ValidaTelefone" & Chr(13) & "Erro: " & Err.Description & Chr(13) & "Linha: " & Erl, vbOKOnly + vbCritical, "ERRO")
End Function

Public Function ValidaCelular(ByVal strCelular As String) As Boolean
          Dim strRetorno          As String
10    On Error GoTo Erro
            
20        strRetorno = Trim(SomenteOsNumerosDaString(strCelular))
30        ValidaCelular = False
          
40        If Len(strRetorno) = 9 Then
50            ValidaCelular = True
60        Else
70            ValidaCelular = False
80        End If
          
90    Exit Function
Erro:
100       Call MsgBox("Erro na função ValidaCelular" & Chr(13) & "Erro: " & Err.Description & Chr(13) & "Linha: " & Erl, vbOKOnly + vbCritical, "ERRO")
End Function


Public Function FormataCNPJ(ByVal strCNPJ As String, Optional ByVal blnLimparSeForInvalido As Boolean = False) As String
10    On Error GoTo Erro
            

20        strCNPJ = SomenteOsNumerosDaString(strCNPJ)
30        If Len(strCNPJ) > 15 Then strCNPJ = Right(strCNPJ, 15)
          
40        strCNPJ = Replace(Replace(Replace(Format(strCNPJ, "#00PT000PT000BAR0000TRAC00"), "PT", ".", , , vbTextCompare), "BAR", "/", , , vbTextCompare), "TRAC", "-", , , vbTextCompare)
          
      '50        If blnLimparSeForInvalido Then
      '60            If Not ValidarCNPJ(strCNPJ) Then
      '70                strCNPJ = ""
      '80            End If
      '90        End If
50        FormataCNPJ = strCNPJ

60    Exit Function
Erro:
70        Call MsgBox("Erro na função FormataCNPJ" & Chr(13) & "Erro: " & Err.Description & Chr(13) & "Linha: " & Erl, vbOKOnly + vbCritical, "ERRO")

          
End Function


Public Function FormataCNPJCPF(ByVal strCNPJCPF As String, Optional ByVal blnLimparSeForInvalido As Boolean = False) As String
10    On Error GoTo Erro
            

20        strCNPJCPF = SomenteOsNumerosDaString(strCNPJCPF)
30        If Len(strCNPJCPF) > 11 Then
40            FormataCNPJCPF = FormataCNPJ(strCNPJCPF, blnLimparSeForInvalido)
50        Else
60            FormataCNPJCPF = FormataCPF(strCNPJCPF, blnLimparSeForInvalido)
70        End If

80    Exit Function
Erro:
90        Call MsgBox("Erro na função FormataCNPJCPF" & Chr(13) & "Erro: " & Err.Description & Chr(13) & "Linha: " & Erl, vbOKOnly + vbCritical, "ERRO")

End Function




Public Function FormataPorcentagem(ByRef strTextoControle As String, _
                                   Optional ByVal PorcMaxima As Double = 100, _
                                   Optional ByVal strCasaDecimal As String = "0.00") As String
          Dim strParsed As String
          Dim i As Long
          
          'Formata Porcentagem Validando o Texto do Controle extraindo apenas o valor numerico
10        strParsed = ""
20        For i = 1 To Len(strTextoControle)
30            If IsNumeric(Mid(strTextoControle, i, 1)) Then
40                strParsed = strParsed & Mid(strTextoControle, i, 1)
50            Else
60                If Mid(strTextoControle, i, 1) = "," Then
70                    strParsed = strParsed & Mid(strTextoControle, i, 1)
80                End If
90            End If
100       Next
110       If strParsed = "" Then strParsed = "0"
120       If CDbl(strParsed) > PorcMaxima Then strParsed = PorcMaxima
130       FormataPorcentagem = Format(CDbl(strParsed) / 100, IIf(strCasaDecimal = "", "0.00", strCasaDecimal) & " %")
End Function


Public Function IGlobal_RetornarMesPorExtenso(ByVal lngMes As Long) As String
10        Select Case lngMes
              Case 1:  IGlobal_RetornarMesPorExtenso = "Janeiro"
20            Case 2:  IGlobal_RetornarMesPorExtenso = "Fevereiro"
30            Case 3:  IGlobal_RetornarMesPorExtenso = "Março"
40            Case 4:  IGlobal_RetornarMesPorExtenso = "Abril"
50            Case 5:  IGlobal_RetornarMesPorExtenso = "Maio"
60            Case 6:  IGlobal_RetornarMesPorExtenso = "Junho"
70            Case 7:  IGlobal_RetornarMesPorExtenso = "Julho"
80            Case 8:  IGlobal_RetornarMesPorExtenso = "Agosto"
90            Case 9:  IGlobal_RetornarMesPorExtenso = "Setembro"
100           Case 10: IGlobal_RetornarMesPorExtenso = "Outubro"
110           Case 11: IGlobal_RetornarMesPorExtenso = "Novembro"
120           Case 12: IGlobal_RetornarMesPorExtenso = "Dezembro"
130       End Select
End Function

'#####################################################################'
'##                     ESCREVER NÚMERO POR EXTENSO                 ##'
'#####################################################################'
Public Function IGlobal_EscreverNumeroPorExtenso(ByVal strNumero As String, Optional ByVal blnObrigaParteDecimal As Boolean = False, _
                                        Optional ByVal strDelimitadorDecimal As String = "ponto", _
                                        Optional ByVal blnNegativo As Boolean = False, _
                                        Optional ByVal strDelimitadorNegativo As String = "menos", _
                                        Optional ByVal strDescricaoParteInteira As String = "", _
                                        Optional ByVal strDescricaoParteDecimal As String = "") As String
          'strNumero                  -> número a ser convertido, deve ser enviado sem . na separação
          '                       de casas de milhar e com ',' como separador decimal
          'blnObrigaParteDecimal      -> Mesmo se o número vier inteiro o sistema escreverá sua parte decimal como sendo 'zero'
          'blnDelimitadorDecimal      -> descrição da ',', tipo "Vírgula", ou "ponto" ou "e"...etc
          'blnNegativo                -> essa rotina desconsidera o sinal do número, ou seja, só trata como negativo (usando o
          '                        strDelimitadorNegativo) se essa variável estiver true
          'strDelimitadorNegativo     -> É a descrição que vem antes do número por extenso no caso de negativo, ou seja, a descrição
          '                        Do '-'. Ex.: "Menos" cinco mil; "Negativo" cinco mil
          'strDescrição Parte Inteira -> É a descrição da parte inteira, vem entre o final da descrição da parte inteira do
          '                        número e a ','. Ex.: Mil "Reais" ponto trinca centavos, ou dez "Metros" e três centímetros
          '                       OBS.: Se a palavra vier "real" ou "dólar", é convertido para o plural "reais" ou "dólares" de
          '                       acordo com a necessidade: "Um real", "Dez Reais", "Um dólar", "Dez dólares"
          'strDescrição parte decimal -> É a descrição da parte decimal, vem no final da descrição da parte decimal.
          '                        Ex.: três metros e dez "centímetros", ou dez reais e quarenta "centavos"
          '                       OBS.: Se a palavra vier "centavo", então é convertida para o plural de acordo com a necessidade
          '                       Ex.: "dez centavos", "um centavo"
          
          Dim lngParteInteira     As Long
          Dim lngParteDecimal     As Long
          
          Dim strPartes()         As String
          
          Dim strParteInteira     As String
          Dim strParteDecimal     As String
          
          Dim strResultado        As String
              
10        strNumero = Replace(strNumero, "-", "", , , vbTextCompare)
20        strPartes = Split(strNumero, ",", , vbTextCompare)
          
30        If UBound(strPartes) > 0 Then
40            If CDbl(strPartes(0)) > 2147000000 Then 'mais do que isso da overflow(na verdade o overflow não é exatamente nesse número, mas é perto)
50                lngParteInteira = CCur(strPartes(0)) Mod 1000000000
60            Else
70                lngParteInteira = CLng(strPartes(0))
80            End If
              
90            If CDbl(strPartes(1)) > 2147000000 Then 'mais do que isso da overflow(na verdade o overflow não é exatamente nesse número, mas é perto)
100               lngParteDecimal = CCur(strPartes(1)) Mod 1000000000
110           Else
120               lngParteDecimal = CLng(strPartes(1))
130           End If
140       Else
150           If CDbl(strPartes(0)) > 2147000000 Then 'mais do que isso da overflow(na verdade o overflow não é exatamente nesse número, mas é perto)
160               lngParteInteira = CCur(strPartes(0)) Mod 1000000000
170           Else
180               lngParteInteira = CLng(strPartes(0))
190           End If
              
200           lngParteDecimal = 0
210       End If
          
220       If UCase(strDescricaoParteInteira) = "REAL" Then
230           If lngParteInteira <> 1 Then
240               strDescricaoParteInteira = "reais"
250           End If
260       End If
          
270       If UCase(strDescricaoParteInteira) = "DÓLAR" Then
280           If lngParteInteira <> 1 Then
290               strDescricaoParteInteira = "dólares"
300           End If
310       End If
          
320       If UCase(strDescricaoParteDecimal) = "CENTAVO" Then
330           If lngParteDecimal <> 1 Then
340               strDescricaoParteDecimal = "centavos"
350           End If
360       End If
          
370       If lngParteInteira = 0 Then
380           strParteInteira = "zero"
390       Else
400           strParteInteira = ConverteNumero(lngParteInteira)
410       End If
          
420       If lngParteDecimal = 0 Then
430           strParteDecimal = "zero"
440       Else
450           strParteDecimal = ConverteNumero(lngParteDecimal)
460       End If
          
470       If lngParteDecimal = 0 Then
480           If blnObrigaParteDecimal Then
490               strResultado = strParteInteira & " " & strDescricaoParteInteira & " " & strDelimitadorDecimal & " " & strParteDecimal & " " & strDescricaoParteDecimal
500           Else
510               strResultado = strParteInteira & " " & strDescricaoParteInteira
520           End If
530       Else
540           strResultado = strParteInteira & " " & strDescricaoParteInteira & " " & strDelimitadorDecimal & " " & strParteDecimal & " " & strDescricaoParteDecimal
550       End If
          
560       If blnNegativo Then
570           IGlobal_EscreverNumeroPorExtenso = strDelimitadorNegativo & " " & strResultado
580       Else
590           IGlobal_EscreverNumeroPorExtenso = strResultado
600       End If
          
End Function

Private Function ConverteNumero(ByVal lngNumero As Long) As String
          
          Dim strBilhao           As String
          Dim strMilhao           As String
          Dim strMilhar           As String
          Dim strCentena          As String
             
10        If lngNumero < 1000 Then
20            ConverteNumero = ConverteCentena(lngNumero)
30            Exit Function
          
40        ElseIf lngNumero >= 1000 And lngNumero < 1000000 Then
50            strCentena = ConverteCentena(lngNumero Mod 1000)
60            strMilhar = ConverteCentena(Val(lngNumero / 1000))
              
70            If strCentena = "" Then
80                ConverteNumero = strMilhar & " mil"
90            Else
100               ConverteNumero = strMilhar & " mil, " & strCentena
110           End If
          
120       ElseIf lngNumero >= 1000000 And lngNumero < 1000000000 Then
              
130           strCentena = ConverteCentena(lngNumero Mod 1000)
140           strMilhar = ConverteCentena(Val(lngNumero / 1000) Mod 1000)
150           strMilhao = ConverteCentena(Val(lngNumero / 1000000))
              
160           If Val(lngNumero / 1000000) = 1 Then
170               ConverteNumero = strMilhao & " milhão"
180           Else
190               ConverteNumero = strMilhao & " milhões"
200           End If
              
210           If strMilhar = "" Then
220               If Not strCentena = "" Then
230                   ConverteNumero = ConverteNumero & ", " & strCentena
240               End If
250           Else
260               ConverteNumero = ConverteNumero & ", " & strMilhar & " mil"
270               If Not strCentena = "" Then
280                   ConverteNumero = ConverteNumero & ", " & strCentena
290               End If
300           End If
              
310       ElseIf lngNumero >= 1000000000 And lngNumero < 1000000000000# Then
320           strCentena = ConverteCentena(lngNumero Mod 1000)
330           strMilhar = ConverteCentena(Val(lngNumero / 1000) Mod 1000)
340           strMilhao = ConverteCentena(Val(lngNumero / 1000000) Mod 1000)
350           strBilhao = ConverteCentena(Val(lngNumero / 1000000000))
              
360           If Val(lngNumero / 1000000000) = 1 Then
370               ConverteNumero = strBilhao & " bilhão"
380           Else
390               ConverteNumero = strBilhao & " bilhões"
400           End If
              
410           If Not strMilhao = "" Then
420               If Val(lngNumero / 1000000) = 1 Then
430                   ConverteNumero = ConverteNumero & ", " & strMilhao & " milhão"
440               Else
450                   ConverteNumero = ConverteNumero & ", " & strMilhao & " milhões"
460               End If
470           End If
              
480           If Not strMilhar = "" Then
490               ConverteNumero = ConverteNumero & ", " & strMilhar & " mil"
500           End If
              
510           If Not strCentena = "" Then
520               ConverteNumero = ConverteNumero & ", " & strCentena
530           End If
              
              
540       Else 'só formata até 999999999999 'antes de trilhão
550           ConverteNumero = ConverteNumero(lngNumero Mod 1000000000)
560           Exit Function
570       End If
                  
              
End Function

Private Function ConverteCentena(ByVal intCentena As Integer) As String
          Dim strCentena      As String
          Dim strDezena       As String
          
10        If intCentena > 0 And intCentena < 100 Then
20            ConverteCentena = ConverteDezena(intCentena)
30            Exit Function
                  
40        ElseIf intCentena = 100 Then
50            ConverteCentena = "cem"
60            Exit Function
              
70        ElseIf intCentena > 100 And intCentena < 200 Then
80            strCentena = "cento"
90            strDezena = ConverteDezena(intCentena Mod 100)
              
100       ElseIf intCentena >= 200 And intCentena < 300 Then
110           strCentena = "duzentos"
120           strDezena = ConverteDezena(intCentena Mod 200)
          
130       ElseIf intCentena >= 300 And intCentena < 400 Then
140           strCentena = "trezentos"
150           strDezena = ConverteDezena(intCentena Mod 300)
          
160       ElseIf intCentena >= 400 And intCentena < 500 Then
170           strCentena = "quatrocentos"
180           strDezena = ConverteDezena(intCentena Mod 400)
              
190       ElseIf intCentena >= 500 And intCentena < 600 Then
200           strCentena = "quinhentos"
210           strDezena = ConverteDezena(intCentena Mod 500)
              
220       ElseIf intCentena >= 600 And intCentena < 700 Then
230           strCentena = "seiscentos"
240           strDezena = ConverteDezena(intCentena Mod 600)
              
250       ElseIf intCentena >= 700 And intCentena < 800 Then
260           strCentena = "setecentos"
270           strDezena = ConverteDezena(intCentena Mod 700)
              
280       ElseIf intCentena >= 800 And intCentena < 900 Then
290           strCentena = "oitocentos"
300           strDezena = ConverteDezena(intCentena Mod 800)
              
310       ElseIf intCentena >= 900 Then
320           strCentena = "novecentos"
330           strDezena = ConverteDezena(intCentena Mod 900)
          
340       End If
          
350       If strDezena = "" Then
360           ConverteCentena = strCentena
370       Else
380           ConverteCentena = strCentena & " e " & strDezena
390       End If
          
End Function

Private Function ConverteDezena(ByVal intDezena As Integer) As String
          Dim strDezena       As String
          Dim strUnidade      As String
          
10        If intDezena > 0 And intDezena < 10 Then
20            ConverteDezena = ConverteUnidade(intDezena)
30            Exit Function
              
40        ElseIf intDezena >= 10 And intDezena < 20 Then
50            Select Case intDezena
                  Case 10: ConverteDezena = "dez"
60                Case 11: ConverteDezena = "onze"
70                Case 12: ConverteDezena = "doze"
80                Case 13: ConverteDezena = "treze"
90                Case 14: ConverteDezena = "quatorze"
100               Case 15: ConverteDezena = "quinze"
110               Case 16: ConverteDezena = "dezesseis"
120               Case 17: ConverteDezena = "dezessete"
130               Case 18: ConverteDezena = "dezoito"
140               Case 19: ConverteDezena = "dezenove"
150           End Select
160           Exit Function
              
170       ElseIf intDezena >= 20 And intDezena < 30 Then
180           strDezena = "vinte"
190           strUnidade = ConverteUnidade(intDezena Mod 20)
              
200       ElseIf intDezena >= 30 And intDezena < 40 Then
210           strDezena = "trinta"
220           strUnidade = ConverteUnidade(intDezena Mod 30)
              
230       ElseIf intDezena >= 40 And intDezena < 50 Then
240           strDezena = "quarenta"
250           strUnidade = ConverteUnidade(intDezena Mod 40)
              
260       ElseIf intDezena >= 50 And intDezena < 60 Then
270           strDezena = "cinqüenta"
280           strUnidade = ConverteUnidade(intDezena Mod 50)
              
290       ElseIf intDezena >= 60 And intDezena < 70 Then
300           strDezena = "sessenta"
310           strUnidade = ConverteUnidade(intDezena Mod 60)
              
320       ElseIf intDezena >= 70 And intDezena < 80 Then
330           strDezena = "setenta"
340           strUnidade = ConverteUnidade(intDezena Mod 70)
              
350       ElseIf intDezena >= 80 And intDezena < 90 Then
360           strDezena = "oitenta"
370           strUnidade = ConverteUnidade(intDezena Mod 80)
              
380       ElseIf intDezena >= 90 Then
390           strDezena = "noventa"
400           strUnidade = ConverteUnidade(intDezena Mod 90)
              
410       End If
          
420       If strUnidade = "" Then
430           ConverteDezena = strDezena
440       Else
450           ConverteDezena = strDezena & " e " & strUnidade
460       End If
              
End Function

Private Function ConverteUnidade(ByVal intUnidade As Integer) As String
10        Select Case intUnidade
              Case 1: ConverteUnidade = "um"
20            Case 2: ConverteUnidade = "dois"
30            Case 3: ConverteUnidade = "três"
40            Case 4: ConverteUnidade = "quatro"
50            Case 5: ConverteUnidade = "cinco"
60            Case 6: ConverteUnidade = "seis"
70            Case 7: ConverteUnidade = "sete"
80            Case 8: ConverteUnidade = "oito"
90            Case 9: ConverteUnidade = "nove"
100       End Select
End Function
'#####################################################################'
'#####################################################################'


'#####################################################################'
'##     TRATAMENDO DOS EVENTOS DOS COMPONENTES DA INTERFACE         ##'
'#####################################################################'
Public Sub TextBoxSomenteNumeros(ByVal strDado As String, _
                                      ByRef KeyAscii As Integer, _
                                      Optional ByVal blnAceitaVirgula As Boolean = False, _
                                      Optional ByVal blnAceitaNegativo As Boolean = False)
          
10        If KeyAscii = 3 Then Exit Sub      'CTRL + C
20        If KeyAscii = 22 Then Exit Sub      'CTRL + V
30        If KeyAscii = vbKeyBack Then Exit Sub
40        If KeyAscii = vbKeyReturn Then Exit Sub
50        If blnAceitaVirgula Then
60            If Chr(KeyAscii) = "," Then
70                If InStr(1, strDado, ",", vbTextCompare) Then
80                    KeyAscii = 0
90                End If
100               Exit Sub
110           End If
120       End If
130       If blnAceitaNegativo Then
140           If Chr(KeyAscii) = "-" Then
150               If InStr(1, strDado, "-", vbTextCompare) Then
160                   KeyAscii = 0
170               End If
180               Exit Sub
190           End If
200       End If
210       If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
End Sub

Public Sub DigitarSomenteCaracteresTelefone(ByVal strDado As String, _
                                        ByRef KeyAscii As Integer, _
                                        Optional ByVal blnAceitaCelular As Boolean = True)
    
    If KeyAscii = 3 Then Exit Sub      'CTRL + C
    If KeyAscii = 22 Then Exit Sub      'CTRL + V
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii = vbKeyReturn Then Exit Sub
    
    TextBoxSomenteNumeros strDado, KeyAscii, False, False
    If blnAceitaCelular Then
        If Len(SomenteOsNumerosDaString(strDado)) >= 9 Then KeyAscii = 0
    Else
        If Len(SomenteOsNumerosDaString(strDado)) >= 8 Then KeyAscii = 0
    End If
    
End Sub

Public Sub ComponenteKeyDown(KeyCode As Integer)
10        Select Case KeyCode
              Case vbKeyReturn: SendKeys "{TAB}"
20            Case vbKeyEscape: SendKeys "+{TAB}"
30        End Select
End Sub

Public Sub ComponenteGotFocus(objTextBox As Object)
10        If TypeOf objTextBox Is TextBox Then
20            objTextBox.SelStart = 0
30            objTextBox.SelLength = Len(objTextBox.Text)
40        End If
End Sub

Public Sub ComponenteLostFocusNumeric(objTextBox As Object, Optional strFormat As String = "#,###,##0.00")
    If Not IsNumeric(objTextBox.Text) Then
        objTextBox.Text = 0
    End If
    
    objTextBox.Text = Format(objTextBox.Text, strFormat)
End Sub
'#####################################################################'
'#####################################################################'




Public Function EmailValido(strEmail As String, Optional ByVal blnNaoExibirMensagem As Boolean = False) As Boolean
    Dim intCharacter    As Integer
    Dim Count           As Integer
    Dim sLetra          As String
    
    EmailValido = True
    strEmail = Trim$(LCase$(strEmail))
    
    If strEmail <> "" Then
        
        'Verifica se o e-mail tem no MÍNIMO 5 caracteres
        If Len(strEmail) < 5 Then   'O e-mail é inválido, pois tem menos de 5 caracteres
            EmailValido = False
        End If
        
        'Verificar a existencia de @s  no e-mail
        For intCharacter = 1 To Len(strEmail)
            If Mid(strEmail, intCharacter, 1) = "@" Then
                Count = Count + 1
            End If
        Next
        
        If Count <> 1 Then 'O e-mail é inválido, pois tem 0 ou mais de 1 @
            EmailValido = False
        Else
            If InStr(strEmail, "@") = 1 Then 'O e-mail é inválido, pois começa com uma @
                EmailValido = False
            ElseIf InStr(strEmail, "@") = Len(strEmail) Then 'O e-mail é inválido, pois termina com uma @
                EmailValido = False
            End If
        End If
        
        intCharacter = 0
        Count = 0
        'Verificar a existencia de pontos (.) no e-mail
        For intCharacter = 1 To Len(strEmail)
            If Mid(strEmail, intCharacter, 1) = "." Then
                Count = Count + 1
            End If
        Next
        'Verifica o número de pontos.TEM que ter PELO MENOS UM ponto.
        If Count < 1 Then 'O e-mail é inválido, pois não tem pontos.
            EmailValido = False
        Else 'O e-mail tem pelo menos 1 ponto.Verificar a posição do ponto:
            If InStr(strEmail, ".") = 1 Then 'O e-mail é inválido, pois começa com um ponto
                EmailValido = False
            ElseIf InStr(strEmail, ".") = Len(strEmail) Then 'O e-mail é inválido, pois termina com um ponto.
                EmailValido = False
            ElseIf InStr(strEmail, "@") = 0 Then
                EmailValido = False
            ElseIf InStr(InStr(strEmail, "@"), strEmail, ".") = 0 Then 'O e-mail é inválido, pois termina com um ponto.
                EmailValido = False
            End If
        End If
        intCharacter = 0
        Count = 0
        
        If InStr(strEmail, "..") > InStr(strEmail, "@") Then 'Verifica se o e-mail não tem pontos consecutivos (..) após a @ .
            EmailValido = False
        End If
        
        For intCharacter = 1 To Len(strEmail) 'Verifica se o e-mail tem caracteres inválidos
            sLetra = Mid$(strEmail, intCharacter, 1)
            If Not (LCase(sLetra) Like "[a-z]" Or sLetra = _
                "@" Or sLetra = "." Or sLetra = "-" Or _
                sLetra = "_" Or IsNumeric(sLetra)) Then 'O e-mail é inválido, pois tem caracteres inválidos
                EmailValido = False
            End If
        Next
    End If
    
    
    If EmailValido = False And blnNaoExibirMensagem = False Then
        MsgBox "O E-mail Informado é Inválido.", vbCritical, "E-mail Inválido"
    Else
        intCharacter = 0
    End If
End Function



Public Sub CriarPath(ByVal strPath As String)
On Error GoTo Erro
    Dim lngNumeroPastas         As Long
    Dim i                       As Integer
    Dim strNovoPath             As String

    Dim strPartesPath()     As String
        
    If Left(strPath, 2) = "\\" And Len(strPath) > 3 Then strPath = "##" & Mid(strPath, 3, Len(strPath) - 2)
    
    strPartesPath = Split(strPath, "\", , vbTextCompare)
    
    lngNumeroPastas = -1
    
    On Error Resume Next
    lngNumeroPastas = UBound(strPartesPath)
    On Error GoTo 0
    If lngNumeroPastas = -1 Then Exit Sub
    
    For i = 0 To lngNumeroPastas
        If Left(strPartesPath(i), 2) = "##" And Len(strPartesPath(i)) > 2 Then strPartesPath(i) = "\\" & Mid(strPartesPath(i), 3, Len(strPartesPath(i)) - 2)
        If strNovoPath <> "" Then
            strNovoPath = strNovoPath & "\" & strPartesPath(i)
        Else
            strNovoPath = strPartesPath(i)
        End If
        If i = 0 And Left(strPartesPath(i), 2) = "\\" Then
            GoTo Proximo
        ElseIf i = 1 And Right(strPartesPath(i), 1) = "$" And Left(strPartesPath(0), 2) = "\\" Then
            GoTo Proximo
        Else
            If Dir(strNovoPath, vbDirectory) = "" Then
                MkDir (strNovoPath)
            End If
        End If
Proximo:
    Next
        
Exit Sub
Erro:
    If Err.Number = 76 And Err.Description = "Path not found" Then
        MsgBox "Caminho não encontrado, por favor, peça alguém do suporte para criar a seguinte pasta:" & Chr(13) & _
            strPath, vbOKOnly + vbCritical, "Atenção"
    Else
       ' 'Call modBDTratarErro_TratarErroInterface("modBDPadrao", "CriarPath", Err.Description, Err.Number, Erl)
    End If
End Sub

Public Function VerificarDiretorioExiste(ByVal strPath As String) As Boolean
10    On Error GoTo Erro
          
20        VerificarDiretorioExiste = False

30        Do While Right(strPath, 1) = "\" 'tirando possíveis \ da direita, isso atrapalha o teste
40            If Len(strPath) <= 1 Then Exit Function
50            strPath = Mid(strPath, 1, Len(strPath) - 1)
60        Loop

70        VerificarDiretorioExiste = False
80        If Dir(strPath, vbDirectory) = "" Then
90            VerificarDiretorioExiste = False
100       Else
110           VerificarDiretorioExiste = True
120       End If

130   Exit Function
Erro:
140       ''Call modBDTratarErro_TratarErroInterface("modBDPadrao", "VerificarDiretorioExiste", Err.Description, Err.Number, Erl)
End Function

Public Function VerificarArquivoExiste(ByVal strPath As String) As Boolean
10    On Error GoTo Erro
20        VerificarArquivoExiste = False

30        If Dir(strPath, vbArchive) = "" Or strPath = "" Then
40            VerificarArquivoExiste = False
50        Else
60            VerificarArquivoExiste = True
70        End If

80    Exit Function
Erro:
90        ''Call modBDTratarErro_TratarErroInterface("modBDPadrao", "VerificarDiretorioExiste", Err.Description, Err.Number, Erl)
End Function

Public Sub TravarComponenteDoFrame(ByRef objComponente As Object)
10    On Error GoTo Erro
20        objComponente.Enabled = False
30        objComponente.BackColor = mConstButtonFace

40    Exit Sub
Erro:
50        ''Call modBDTratarErro_TratarErroInterface("modBDPadrao", "TravarComponenteDoFrame", Err.Description, Err.Number, Erl)
End Sub

Public Sub DestravarComponenteDoFrame(ByRef objComponente As Object)
10    On Error GoTo Erro
20        objComponente.Enabled = True
30        objComponente.BackColor = mConstWrite

40    Exit Sub
Erro:
50        'Call modBDTratarErro_TratarErroInterface("modBDPadrao", "DestravarComponenteDoFrame", Err.Description, Err.Number, Erl)
End Sub

Public Function NameOfPC() As String
On Error Resume Next
    Dim MachineName As String
    Dim NameSize As Long
    Dim X As Long
    
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
    
    NameOfPC = MachineName
On Error GoTo 0
End Function

'#####################################################################'
'#####################################################################'
'       PARTE REFERENTE A ARQUIVOS .INI
Public Function LerINI(Secao As String, Entrada As String, Path As String) As String
    Dim retlen As String
    Dim Ret    As String
    
    'Path = nome do arquivo ini
    'Secao=O que esta entre []
    'Entrada=nome do que se encontra antes do sinal de igual
    Ret = String$(255, 0)
    retlen = GetPrivateProfileString(Secao, Entrada, "", Ret, Len(Ret), Path)
    Ret = Left$(Ret, retlen)
    LerINI = Ret
End Function

Public Sub EscreverINI(Secao As String, Entrada As String, Texto As String, Path As String)
    'Path= nome do arquivo ini
    'Secao= O que esta entre []
    'Entrada= nome do que se encontra antes do sinal de igual
    'texto= valor que vem depois do igual
    WritePrivateProfileString Secao, Entrada, Texto, Path
End Sub
'#####################################################################'
'#####################################################################'

Public Function DownloadArquivoInternet(URL As String, LocalFilename As String) As Boolean
    Dim lngRetVal As Long
    
    lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
    If lngRetVal = 0 Then DownloadArquivoInternet = True
End Function

'#####################################################################'
'#####################################################################'
'FUNÇÃO RESPONSÁVEL EM FAZER ISSO: (Teste   Teste        Teste                 Teste)
'SE TRANSFORMAR EM:                (Teste Teste Teste Teste)
Public Function RemoverEspacos(ByRef strObservacao As String) As String
    strObservacao = Trim(strObservacao)
    strObservacao = Replace(strObservacao, Chr(9), " ", , , vbTextCompare)
    Do While InStr(1, strObservacao, "  ", vbTextCompare) > 0
        strObservacao = Replace(strObservacao, "  ", "", , , vbTextCompare)
    Loop
    RemoverEspacos = strObservacao
End Function
'#####################################################################'
'#####################################################################'

'#####################################################################'
'#####################################################################'
Public Sub TextBoxParaNumeros(TextBox As TextBox, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyEscape Then Exit Sub
    If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
    If KeyAscii = Asc(",") Then
        If InStr(1, TextBox.Text, ",", vbTextCompare) > 0 Then
            KeyAscii = 0
        End If
    Else
        If Not (KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Then
            KeyAscii = 0
        End If
    End If
End Sub


'
'Public Function RetornaLoginUsuarioLogado() As String
'Dim objRSUsuario  As New Recordset
'
'10    On Error GoTo erro
'
'20        objRSUsuario.Open "dbo.usp_SelecionarInformacoesdeLogin", gSMConexao.Conexao, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
'30        If Not (objRSUsuario.BOF And objRSUsuario.EOF) Then RetornaLoginUsuarioLogado = (objRSUsuario!Nome_VC)
'
'40    Exit Function
'erro:
'50        'Call modBDTratarErro_TratarErroInterface("modBDPadrao", "StatusBar_Atualizar", Err.Description, Err.Number, Erl)
'End Function

Public Function DivisaoEDecimal(dblDividendo As Double, dblDivisor As Double) As Boolean
    DivisaoEDecimal = False
    If InStr(1, CStr(dblDividendo / dblDivisor), ",") > 0 Then DivisaoEDecimal = True
End Function

Public Function RemoveAcentos(ByVal strTexto As String) As String
    DoEvents
    RemoveAcentos = Replace(strTexto, "Á", "A")
    RemoveAcentos = Replace(RemoveAcentos, "À", "A")
    RemoveAcentos = Replace(RemoveAcentos, "Ã", "A")
    RemoveAcentos = Replace(RemoveAcentos, "Â", "A")
    RemoveAcentos = Replace(RemoveAcentos, "á", "a")
    RemoveAcentos = Replace(RemoveAcentos, "à", "a")
    RemoveAcentos = Replace(RemoveAcentos, "ã", "a")
    RemoveAcentos = Replace(RemoveAcentos, "â", "a")
    RemoveAcentos = Replace(RemoveAcentos, "É", "E")
    RemoveAcentos = Replace(RemoveAcentos, "È", "E")
    RemoveAcentos = Replace(RemoveAcentos, "Ê", "E")
    RemoveAcentos = Replace(RemoveAcentos, "Ë", "E")
    RemoveAcentos = Replace(RemoveAcentos, "é", "e")
    RemoveAcentos = Replace(RemoveAcentos, "è", "e")
    RemoveAcentos = Replace(RemoveAcentos, "ê", "e")
    RemoveAcentos = Replace(RemoveAcentos, "ë", "e")
    RemoveAcentos = Replace(RemoveAcentos, "Í", "I")
    RemoveAcentos = Replace(RemoveAcentos, "Ì", "I")
    RemoveAcentos = Replace(RemoveAcentos, "Î", "I")
    RemoveAcentos = Replace(RemoveAcentos, "Î", "I")
    RemoveAcentos = Replace(RemoveAcentos, "Ï", "I")
    RemoveAcentos = Replace(RemoveAcentos, "í", "i")
    RemoveAcentos = Replace(RemoveAcentos, "ì", "i")
    RemoveAcentos = Replace(RemoveAcentos, "î", "i")
    RemoveAcentos = Replace(RemoveAcentos, "ï", "i")
    RemoveAcentos = Replace(RemoveAcentos, "Ó", "O")
    RemoveAcentos = Replace(RemoveAcentos, "Ò", "O")
    RemoveAcentos = Replace(RemoveAcentos, "Õ", "O")
    RemoveAcentos = Replace(RemoveAcentos, "Ô", "O")
    RemoveAcentos = Replace(RemoveAcentos, "Ö", "O")
    RemoveAcentos = Replace(RemoveAcentos, "ó", "o")
    RemoveAcentos = Replace(RemoveAcentos, "ò", "o")
    RemoveAcentos = Replace(RemoveAcentos, "ô", "o")
    RemoveAcentos = Replace(RemoveAcentos, "ô", "o")
    RemoveAcentos = Replace(RemoveAcentos, "ö", "o")
    RemoveAcentos = Replace(RemoveAcentos, "Ú", "O")
    RemoveAcentos = Replace(RemoveAcentos, "Ù", "O")
    RemoveAcentos = Replace(RemoveAcentos, "Û", "O")
    RemoveAcentos = Replace(RemoveAcentos, "Ü", "O")
    RemoveAcentos = Replace(RemoveAcentos, "ú", "O")
    RemoveAcentos = Replace(RemoveAcentos, "ù", "o")
    RemoveAcentos = Replace(RemoveAcentos, "û", "o")
    RemoveAcentos = Replace(RemoveAcentos, "ü", "o")
    RemoveAcentos = Replace(RemoveAcentos, "Ç", "C")
    RemoveAcentos = Replace(RemoveAcentos, "ç", "c")
End Function

Public Function PreparaTexto(ByVal strTexto As String) As String
    PreparaTexto = Replace(Replace(Replace(Replace(PreparaTexto, "&", ""), "ª", ""), "º", ""), "²", " ")
    PreparaTexto = Replace(Replace(Replace(Replace(Replace(PreparaTexto, "³", ""), "¹", ""), "/", ""), "//", ""), "///", "/")
    PreparaTexto = Replace(Replace(Replace(Replace(strTexto, "<", ""), ">", ""), "  ", ""), vbNewLine, " ")
    PreparaTexto = Replace(Replace(Replace(PreparaTexto, "ª", ""), "º", ""), "º", "")
    PreparaTexto = Replace(Replace(Replace(Replace(PreparaTexto, "³", ""), "¹", ""), "²", ""), "&", "e")
    PreparaTexto = Trim(Replace(Replace(PreparaTexto, "–", "-"), "º", " "))
    PreparaTexto = Replace(PreparaTexto, "°", "")
    PreparaTexto = Replace(PreparaTexto, "§", "")
    PreparaTexto = Replace(PreparaTexto, "Ê", "E")
    PreparaTexto = Replace(PreparaTexto, "ê", "e")
    PreparaTexto = Replace(PreparaTexto, "´", "")
    PreparaTexto = Replace(PreparaTexto, "'", "")
    PreparaTexto = Replace(PreparaTexto, "’", "")
    PreparaTexto = Replace(PreparaTexto, "™", "")
    PreparaTexto = Replace(Replace(PreparaTexto, "“", ""), "”", "")
    PreparaTexto = Replace(PreparaTexto, "   ", " ")  '3 espaços
    PreparaTexto = Replace(PreparaTexto, "  ", " ")   '2 espaços
    PreparaTexto = Replace(PreparaTexto, "    ", " ") '4 espaços
    PreparaTexto = Replace(PreparaTexto, "ã", "a")
    PreparaTexto = Replace(PreparaTexto, "Ã", "A")
    PreparaTexto = Replace(PreparaTexto, "õ", "o")
    PreparaTexto = Replace(PreparaTexto, "Õ", "O")
    PreparaTexto = Replace(PreparaTexto, "Ó", "O")
    PreparaTexto = Replace(PreparaTexto, "ô", "o")
    PreparaTexto = Replace(PreparaTexto, "í", "i")
    PreparaTexto = Replace(PreparaTexto, "I", "I")
    PreparaTexto = Replace(PreparaTexto, "Ç", "C")
    PreparaTexto = Replace(PreparaTexto, "ç", "c")
    PreparaTexto = Replace(PreparaTexto, "€", "C")
    PreparaTexto = Replace(PreparaTexto, "—", "-")
    PreparaTexto = RemoverEspacos(PreparaTexto)
End Function



Public Function AbrirSite(ByVal WWW As String, Optional ByVal hwnd As Long = 0)
    ShellExecute hwnd, "open", WWW, vbNullString, vbNullString, conSwNormal
End Function

Public Function EnviarEmail(ByVal Email As String, Optional ByVal hwnd As Long = 0, Optional ByVal Assunto As String = "", Optional ByVal Corpo As String = "", Optional ByVal ComCopia As String = "", Optional ByVal ComCopiaOculta As String = "")
    Dim strComando As String
    
    'constroi a string do email
    If Len(Assunto) Then strComando = "&Subject=" & Assunto
    If Len(Corpo) Then strComando = strComando & "&Body=" & Corpo
    If Len(ComCopia) Then strComando = strComando & "&CC=" & ComCopia
    If Len(ComCopiaOculta) Then strComando = strComando & "&BCC=" & ComCopiaOculta
    
    'substitui o primeiro &
    'com interrogacao
    If Len(strComando) Then
       Mid(strComando, 1, 1) = "?"
    End If
    
    'Inclui o comando mailto: e o endereço de e-mail
    strComando = "mailto:" & Email & strComando
    
    'executa o comando via API
    Call ShellExecute(hwnd, "open", strComando, vbNullString, vbNullString, conSwNormal)

End Function

Public Function AcertarValor(ByVal curValor As Currency, Optional intCasasDecimais As Integer = 2) As String
    Dim strCasas          As String
    Dim Str               As String
    
    strCasas = String(intCasasDecimais, "0")
    If intCasasDecimais = 0 Then
        AcertarValor = curValor
        Exit Function
    End If

    If curValor = 0 Then
        AcertarValor = "0.00"
        Exit Function
    Else
        Str = Replace(CStr(Format(curValor, "#####." & strCasas)), ",", ".")
        If Mid(Str, 1, 1) = "." Then
            Str = "0" & Str
        End If
    End If
    AcertarValor = Str
End Function

'
'Public Function modBDPadrao_RetornaDiasUteisNoPeriodo(datDataInicial As Date, datDataFinal As Date, blnSabadoEDiaUtil As Boolean, ByRef objRsDiasUteis As Recordset) As Long
'          'Retorna o número de dias ulteis no período
'10    On Error GoTo erro
'
'20        Set gobjCmd.ActiveConnection = gSMConexao.Conexao
'30        gobjCmd.CommandText = "dbo.usp_SelecionarDiasUteisNoPeriodo"
'40        gobjCmd.CommandType = adCmdStoredProc
'50        gobjCmd.CommandTimeout = 1000
'
'60        gobjCmd.Parameters("@DataInicial_DT").Value = datDataInicial
'70        gobjCmd.Parameters("@DataFinal_DT").Value = datDataFinal
'80        gobjCmd.Parameters("@SabadoEDiaUtil_BT").Value = blnSabadoEDiaUtil
'
'90        Set objRsDiasUteis = gobjCmd.Execute
'
'100       If Not objRsDiasUteis Is Nothing Then
'110           modBDPadrao_RetornaDiasUteisNoPeriodo = objRsDiasUteis.RecordCount
'120       End If
'
'130   Exit Function
'erro:
'140       'Call modBDTratarErro_TratarErroInterface("modBDPadrao", "modBDPadrao_RetornaDiasUteisNoPeriodo", Err.Description, Err.Number, Erl)
'End Function

Public Function IGlobal_RetornarMesPorExtensoAbreviado(ByVal lngMes As Long) As String
10        Select Case lngMes
              Case 1:  IGlobal_RetornarMesPorExtensoAbreviado = "Jan."
20            Case 2:  IGlobal_RetornarMesPorExtensoAbreviado = "Fev."
30            Case 3:  IGlobal_RetornarMesPorExtensoAbreviado = "Mar."
40            Case 4:  IGlobal_RetornarMesPorExtensoAbreviado = "Abr."
50            Case 5:  IGlobal_RetornarMesPorExtensoAbreviado = "Maio"
60            Case 6:  IGlobal_RetornarMesPorExtensoAbreviado = "Jun."
70            Case 7:  IGlobal_RetornarMesPorExtensoAbreviado = "Jul."
80            Case 8:  IGlobal_RetornarMesPorExtensoAbreviado = "Ago."
90            Case 9:  IGlobal_RetornarMesPorExtensoAbreviado = "Set."
100           Case 10: IGlobal_RetornarMesPorExtensoAbreviado = "Out."
110           Case 11: IGlobal_RetornarMesPorExtensoAbreviado = "Nov."
120           Case 12: IGlobal_RetornarMesPorExtensoAbreviado = "Dez."
130       End Select
End Function

Public Function CarregarConfiguracoesColunasGrid(objGrid As Object, strAPPTitle As String, strFORMANAME As String) As Boolean
       Dim i          As Integer
10     On Error GoTo Erro
20        For i = 0 To objGrid.Columns.Count - 1
30            objGrid.Columns(i).WrapText = True
40            objGrid.Columns(i).AllowSizing = True

50            objGrid.Columns(i).Width = GetSetting("MAS", strAPPTitle & "." & strFORMANAME & "." & objGrid.Name, _
                                                                "objGrid.Column." & objGrid.Columns(i).DataField & ".Width", objGrid.Columns(i).Width)
                  
60            objGrid.Columns(i).Order = GetSetting("MAS", strAPPTitle & "." & strFORMANAME & "." & objGrid.Name, _
                                                                "objGrid.Column." & objGrid.Columns(i).DataField & ".Order", objGrid.Columns(i).Order)
                                                                
                                                                
61            objGrid.Columns(i).Visible = IIf(GetSetting("MAS", strAPPTitle & "." & strFORMANAME & "." & objGrid.Name, _
                                                                 "objGrid.Column." & objGrid.Columns(i).DataField & ".Visible", _
                                                                 IIf(objGrid.Columns(i).Visible, "1", "0")) = "1", True, False)
70        Next
80    CarregarConfiguracoesColunasGrid = True
90    Exit Function
Erro:
100   CarregarConfiguracoesColunasGrid = False
110   On Error Resume Next
End Function

Public Function SalvarConfiguracoesColunasGrid(objGrid As Object, strAPPTitle As String, strFormName As String) As Boolean
      Dim i           As Integer
10     On Error GoTo Erro
20        For i = 0 To objGrid.Columns.Count - 1

30            Call SaveSetting("MAS", strAPPTitle & "." & strFormName & "." & objGrid.Name, _
                               "objGrid.Column." & objGrid.Columns(i).DataField & ".Width", objGrid.Columns(i).Width)
       
              
40            Call SaveSetting("MAS", strAPPTitle & "." & strFormName & "." & objGrid.Name, _
                               "objGrid.Column." & objGrid.Columns(i).DataField & ".Order", objGrid.Columns(i).Order)
                               
                               
41            Call SaveSetting("MAS", strAPPTitle & "." & strFormName & "." & objGrid.Name, _
                            "objGrid.Column." & objGrid.Columns(i).DataField & ".Visible", IIf(objGrid.Columns(i).Visible, "1", "0"))
                               
50        Next
60    SalvarConfiguracoesColunasGrid = True
Erro:
70    Exit Function
80    SalvarConfiguracoesColunasGrid = False
90    On Error Resume Next
End Function

Public Function RetornaValorTag(strStringPrincipal As String, strTag As String) As String
    Dim strRetorno              As String
    Dim strParte1Removida       As String
On Error Resume Next

    strRetorno = ""
    If InStr(1, strStringPrincipal, "<" & strTag & ">", vbTextCompare) > 0 And InStr(1, strStringPrincipal, "</" & strTag & ">", vbTextCompare) > 0 Then
        strParte1Removida = Mid(strStringPrincipal, InStr(1, strStringPrincipal, "<" & strTag & ">", vbTextCompare) + Len(strTag) + 2)
        strRetorno = Left(strParte1Removida, InStr(1, strParte1Removida, "</" & strTag & ">", vbTextCompare) - 1)
    End If
    RetornaValorTag = strRetorno
End Function



Public Function RetornaStringDaAPI(ByVal strURL As String, Optional ByVal strGETPOST As String = "GET", Optional ByVal strContentType As String = "application/x-www-form-urlencoded") As String
 On Error GoTo Erro:
 
    RetornaStringDaAPI = ""
    
    Dim objMyMSXML     As Variant
    Dim strRetornoApi As String
 
    Set objMyMSXML = CreateObject("Microsoft.XmlHttp")
    objMyMSXML.Open strGETPOST, strURL, False
    objMyMSXML.setRequestHeader "Content-Type", strContentType
    objMyMSXML.setRequestHeader "User-Agent", "Firefox 3.6.4"
    objMyMSXML.Send "" 'txtJson.Text 'Replace(Replace(Replace(a, """", " '", , , vbTextCompare), Chr(10), "", , , vbTextCompare), Chr(13), "", , , vbTextCompare) 'a '"param1=value2&param2=value2"
    strRetornoApi = objMyMSXML.responseText
    RetornaStringDaAPI = strRetornoApi
    'RecebeJSON = strRetornoAPI
    
    'VALIDAEMAILPORTO_VerificaEmailEValidoSimples = VALIDAEMAILPORTO_RetornaValorTagAPIIPortoBOOLEAN(strTag, strRetornoApi)
    
Erro:
    
End Function

Public Sub SomenteMaiusculas(KeyAscii As Integer)
    Dim lngAscii            As Long
    'If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
On Error GoTo Erro
    lngAscii = KeyAscii
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Erro:
    KeyAscii = lngAscii
End Sub

Public Sub SomenteMinusculas(KeyAscii As Integer)
    Dim lngAscii            As Long
    'If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
On Error GoTo Erro
    lngAscii = KeyAscii
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
Exit Sub
Erro:
    KeyAscii = lngAscii
End Sub
Public Function FormataCEP(strCEP As String) As String
    strCEP = SomenteOsNumerosDaString(strCEP)
    If Len(strCEP) > 3 Then
        FormataCEP = Mid(strCEP, 1, Len(strCEP) - 3) & "-" & Right(strCEP, 3)
    Else
        FormataCEP = strCEP
    End If
End Function

Public Function LimparTabulacoesDaString(strTexto As String) As String

Dim strParsed  As String
Dim i          As Integer
Dim oldChar    As String

'Se for encontrado um espaço no inicio ou final da String, será removido
'Se for encontrada uma tabulação, será removida da String

    strParsed = ""
    oldChar = ""
    
    For i = 1 To Len(strTexto)
    
        If Mid(strTexto, i, 1) = " " Then
            If oldChar = " " Then
                oldChar = Mid(strTexto, i, 1)
            ElseIf Mid(strTexto, i + 1, 1) <> " " And oldChar <> "" Then
                strParsed = strParsed & Mid(strTexto, i, 1)
                oldChar = Mid(strTexto, i, 1)
            End If
            oldChar = Mid(strTexto, i, 1)
        ElseIf Mid(strTexto, i, 1) = vbTab Then
           oldChar = Mid(strTexto, i, 1)
        ElseIf Mid(strTexto, i, 1) = "Char(13)" Then
           oldChar = Mid(strTexto, i, 1)
        ElseIf Mid(strTexto, i, 1) = "Char(9)" Then
           oldChar = Mid(strTexto, i, 1)
        Else
            strParsed = strParsed & Mid(strTexto, i, 1)
            oldChar = Mid(strTexto, i, 1)
        End If
        
    Next
    
    LimparTabulacoesDaString = strParsed
    
End Function


Public Sub OrdenarColunaTrueDB(ByRef objTDBGrid As Object, ByVal ColIndex As Long, _
                                                     ByRef imgcima As Object, ByRef imgbaixo As Object, _
                                                     Optional blnDestacarColOrdenada As Boolean = False)
          Dim intCont             As Integer
          Dim objrs               As New Recordset
          
          
On Error Resume Next
10        If objTDBGrid.DataSource Is Nothing Then Exit Sub
          
20        Set objrs = objTDBGrid.DataSource
30        If objTDBGrid.Columns(ColIndex).Tag > 0 Then
40            If objTDBGrid.Columns(ColIndex).Tag = 1 Then
50                Set objTDBGrid.Columns(ColIndex).HeadingStyle.ForegroundPicture = imgcima.Picture
60                objTDBGrid.Columns(ColIndex).HeadingStyle.ForegroundPicturePosition = 1 'dbgFPRight
70                objTDBGrid.Columns(ColIndex).Tag = 2
80                If Not objrs Is Nothing Then
90                    If Not (objrs.EOF And objrs.BOF) Then
100                       If objTDBGrid.Columns(ColIndex).DataField <> "" Then
110                           objrs.Sort = objTDBGrid.Columns(ColIndex).DataField & " ASC"
120                       End If
130                   End If
140               End If
150           Else
160               Set objTDBGrid.Columns(ColIndex).HeadingStyle.ForegroundPicture = imgbaixo.Picture
170               objTDBGrid.Columns(ColIndex).HeadingStyle.ForegroundPicturePosition = 1 'dbgFPRight
180               objTDBGrid.Columns(ColIndex).Tag = 1
190               If Not objrs Is Nothing Then
200                   If Not (objrs.EOF And objrs.BOF) Then
210                       If objTDBGrid.Columns(ColIndex).DataField <> "" Then
220                           objrs.Sort = objTDBGrid.Columns(ColIndex).DataField & " DESC"
230                       End If
240                   End If
250               End If
260           End If
270       Else
280           For intCont = 0 To objTDBGrid.Columns.Count - 1
290               objTDBGrid.Columns(intCont).HeadingStyle.ForegroundPicture = Null
300               objTDBGrid.Columns(intCont).Tag = 0
310           Next intCont
320           Set objTDBGrid.Columns(ColIndex).HeadingStyle.ForegroundPicture = imgcima.Picture
330           objTDBGrid.Columns(ColIndex).HeadingStyle.ForegroundPicturePosition = 1 'dbgFPRight
340           objTDBGrid.Columns(ColIndex).Tag = 2
350           If Not objrs Is Nothing Then
360               If Not (objrs.EOF And objrs.BOF) Then
370                       If objTDBGrid.Columns(ColIndex).DataField <> "" Then
380                           objrs.Sort = objTDBGrid.Columns(ColIndex).DataField & " ASC"
390                       End If
400               End If
410           End If
420       End If

430       If blnDestacarColOrdenada Then
440          For intCont = 0 To objTDBGrid.Columns.Count - 1
450              objTDBGrid.Columns(intCont).Style.BackColor = vbWhite
460          Next
470          objTDBGrid.Columns(ColIndex).Style.BackColor = vbYellow
480      End If
On Error GoTo 0
End Sub

Public Function RetornaAcessoPorUsuarioEPermissao(ByVal lngUsuario As Long, ByVal lngPermissao As Long) As Boolean
10    On Error GoTo Erro

20        Set gobjCmd.ActiveConnection = gSMConexao.Conexao
30        gobjCmd.CommandText = "usp_RetornaAcessoPorUsuarioEPermissao"
40        gobjCmd.CommandType = adCmdStoredProc
50        gobjCmd.CommandTimeout = 1000
        
60        With gobjCmd
70            .Parameters("@Permissao_IN").Value = lngPermissao
80            .Parameters("@Usuario_IN").Value = lngUsuario
90        End With
100       gobjCmd.Execute adExecuteNoRecords
110       RetornaAcessoPorUsuarioEPermissao = NB(gobjCmd.Parameters("@Acesso_BT").Value)

120   Exit Function
Erro:
130      Call MsgBox("Erro no módulo: " & "modPadrao" & vbCrLf & "RetornaAcessoPorUsuarioEPermissao" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Function

Public Function RetornaCodigoUsuarioPorLogin(ByVal strLogin As String) As Long
10    On Error GoTo Erro

20        Set gobjCmd.ActiveConnection = gSMConexao.Conexao
30        gobjCmd.CommandText = "usp_SelecionarUsuarioPorLogin"
40        gobjCmd.CommandType = adCmdStoredProc
50        gobjCmd.CommandTimeout = 1000
        
60        With gobjCmd
70            .Parameters("@LoginUsuario_VC").Value = strLogin
80        End With

90        gobjCmd.Execute , adExecuteNoRecords
          
100       RetornaCodigoUsuarioPorLogin = NZ(gobjCmd.Parameters("@CodigoUsuario_IN").Value)

110   Exit Function
Erro:
120      Call MsgBox("Erro no módulo: " & "modPadrao" & vbCrLf & "RetornaUsuarioPorLogin" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")


End Function
