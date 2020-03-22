VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} rptFichaJogador 
   Caption         =   "CadJogador - rptFichaJogador (ActiveReport)"
   ClientHeight    =   10590
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   11250
   _ExtentX        =   19844
   _ExtentY        =   18680
   SectionData     =   "rptFichaJogador.dsx":0000
End
Attribute VB_Name = "rptFichaJogador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mbitFoto() As Byte
Dim mlngCodigo As Long
Dim mstrEquipe As String
Dim mstrJogador As String
Dim mdatNascimento As Date
Dim mstrLocal As String
Dim mstrCertidao As String
Dim mstrCartorio As String
Dim mstrIdentidade As String
Dim mstrOrgao As String
Dim mstrPai As String
Dim mstrMae As String
Dim mstrEndereco As String
Dim mstrBairro As String
Dim mstrCidade As String
Dim mstrFacebook As String
Dim mstrEscola As String
Dim mstrEnderecoEscola As String
Dim mstrBairroEscola As String
Dim mstrCidadeEscola As String
Dim mstrTelefoneEscola As String
Dim mstrFacebookEscola As String


Public Property Let Parametros1(bitFoto() As Byte)
    mbitFoto() = bitFoto()
End Property
Public Property Let Parametros111(lngCodigo As Long)
    mlngCodigo = lngCodigo
End Property
Public Property Let Parametros2(strEquipe As String)
    mstrEquipe = strEquipe
End Property
Public Property Let Parametros3(strJogador As String)
    mstrJogador = strJogador
End Property
Public Property Let Parametros4(datNascimento As Date)
    mdatNascimento = datNascimento
End Property
Public Property Let Parametros5(strLocal As String)
    mstrLocal = strLocal
End Property
Public Property Let Parametros6(strCertidao As String)
    mstrCertidao = strCertidao
End Property
Public Property Let Parametros7(strCartorio As String)
    mstrCartorio = strCartorio
End Property
Public Property Let Parametros8(strIdentidade As String)
    mstrIdentidade = strIdentidade
End Property
Public Property Let Parametros9(strOrgao As String)
    mstrOrgao = strOrgao
End Property
Public Property Let Parametros10(strPai As String)
    mstrPai = strPai
End Property
Public Property Let Parametros11(strmae As String)
     mstrMae = strmae
End Property
Public Property Let Parametros12(strEndereco As String)
    mstrEndereco = strEndereco
End Property
Public Property Let Parametros13(strBairro As String)
    mstrBairro = strBairro
End Property
Public Property Let Parametros14(strCidade As String)
    mstrCidade = strCidade
End Property
Public Property Let Parametros15(strFacebook As String)
    mstrFacebook = strFacebook
End Property
Public Property Let Parametros16(strEscola As String)
    mstrEscola = strEscola
End Property
Public Property Let Parametros17(strEnderecoEscola As String)
    mstrEnderecoEscola = strEnderecoEscola
End Property
Public Property Let Parametros18(strBairroEscola As String)
    mstrBairroEscola = strBairroEscola
End Property
Public Property Let Parametros19(strCidadeEscola As String)
    mstrCidadeEscola = strCidadeEscola
End Property
Public Property Let Parametros20(strTelefoneEscola As String)
    mstrTelefoneEscola = strTelefoneEscola
End Property
Public Property Let Parametros21(strFacebookEscola As String)
    mstrFacebookEscola = strFacebookEscola
End Property

Private Sub ActiveReport_ReportStart()
Dim binIMG() As Byte

    fldCodigo.Text = mlngCodigo
    fldEquipe.Text = mstrEquipe
    fldJogador.Text = mstrJogador
    fldNascimento.Text = Format(IIf(Val(mdatNascimento) = 0, "", mdatNascimento), "dd/mm/yyyy")
    fldLocal.Text = mstrLocal
    fldCertidao.Text = mstrCertidao
    fldCartorio.Text = mstrCartorio
    fldIdentidade.Text = mstrIdentidade
    fldOrgao.Text = mstrOrgao
    fldPai.Text = mstrPai
    fldMae.Text = mstrMae
    fldEndereco.Text = mstrEndereco
    fldBairro.Text = mstrBairro
    fldCidade.Text = mstrCidade
    fldFacebook.Text = mstrFacebook
    fldEscola.Text = mstrEscola
    fldEnderecoEscola.Text = mstrEnderecoEscola
    fldBairroEscola.Text = mstrBairroEscola
    fldCidadeEscola.Text = mstrCidadeEscola
    fldTelefoneEscola.Text = mstrTelefoneEscola
    fldFacebookEscola.Text = mstrFacebookEscola
    
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

