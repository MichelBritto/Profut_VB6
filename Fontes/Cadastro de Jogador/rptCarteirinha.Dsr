VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} rptCarteirinha 
   Caption         =   "CadJogador - rptCarteirinha (ActiveReport)"
   ClientHeight    =   12570
   ClientLeft      =   30
   ClientTop       =   450
   ClientWidth     =   16470
   _ExtentX        =   29051
   _ExtentY        =   22172
   SectionData     =   "rptCarteirinha.dsx":0000
End
Attribute VB_Name = "rptCarteirinha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlngCodigo As Long
Dim mstrApelido As String
Dim mstrNomeCompleto As String
Dim mstrEquipe As String
Dim mstrCartegoria As String
Dim mdatNascimento As Date
Dim mstrPai As String
Dim mstrMae As String
Dim mlngCamisa As Long
Dim mlngSexo As Long '1 masculino 2 feminino

Dim mbitFoto() As Byte

Public Property Let Codigo(lngParam As Long)
    mlngCodigo = lngParam
End Property
Public Property Let Apelido(strParam As String)
    mstrApelido = strParam
End Property
Public Property Let Nome(strParam As String)
    mstrNomeCompleto = strParam
End Property
Public Property Let Equipe(strParam As String)
    mstrEquipe = strParam
End Property
Public Property Let Cartegoria(strParam As String)
    mstrCartegoria = strParam
End Property
Public Property Let Nascimento(datParam As Date)
    mdatNascimento = datParam
End Property
Public Property Let Pai(strParam As String)
    mstrPai = strParam
End Property
Public Property Let Mae(strParam As String)
    mstrMae = strParam
End Property
Public Property Let Camisa(strParam As String)
    mlngCamisa = strParam
End Property
Public Property Let Foto(bitParam() As Byte)
    mbitFoto() = bitParam()
End Property
Public Property Let Sexo(lngSexo As Long)
    mlngSexo = lngSexo
End Property

Private Sub ActiveReport_ReportStart()
Dim binIMG() As Byte

    fldCodigo.Text = mlngCodigo
    fldApelido.Text = mstrApelido
    fldNomeCompleto.Text = mstrNomeCompleto
    fldNomeEquipe.Text = mstrEquipe
    fldCartegoria.Text = mstrCartegoria
    
    fldDataNascimento.Text = Format(mdatNascimento, "dd/mm/yyyy")
    fldDiaNasc.Text = Left(fldDataNascimento.Text, 2)
    fldMesNasc.Text = Mid(fldDataNascimento.Text, 4, 2)
    fldAnoNasc.Text = Right(fldDataNascimento.Text, 2)
        
    fldNomePai.Text = mstrPai
    fldNomeMae.Text = mstrMae
    fldNumero.Text = mlngCamisa
    
    If mlngSexo = 1 Then
        lblXM.Caption = "X"
        lblXF.Caption = ""
    Else
        lblXM.Caption = ""
        lblXF.Caption = "X"
    End If
    
'---------------------------------------------------------
    'AQUI TRATO A IMAGEM BINÁRIA
    On Error Resume Next
    binIMG() = mbitFoto()
    If Val(binIMG(1)) <> 0 Then
        imgFoto.Picture = Nothing
        imgFoto.SizeMode = 1


        Dim b()  As Byte
        Dim ff   As Long
        Dim Arquivo As String
    
        'On Error GoTo ErrHandler
        'Call GetRandomArquivoName(Arquivo)
        Arquivo = "tempimg.bmp"
        ff = FreeFile
        Open Arquivo For Binary Access Write As ff
        b() = binIMG()
        Put ff, , b()
        Close ff
        Erase b
        imgFoto.Picture = LoadPicture(Arquivo)
        'Set GetImageFromField = LoadPicture(Arquivo)
        Kill Arquivo
    End If
'---------------------------------------------------------
    
End Sub

