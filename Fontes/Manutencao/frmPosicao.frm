VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPosicao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProFut - Manutenção de Posição do Atleta"
   ClientHeight    =   6420
   ClientLeft      =   7845
   ClientTop       =   3045
   ClientWidth     =   6585
   Icon            =   "frmPosicao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmPrincipal 
      Height          =   5505
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6585
      Begin VB.Frame fraAltercoes 
         Caption         =   "Adicionar/Alterar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   30
         TabIndex        =   1
         Top             =   120
         Width           =   6525
         Begin VB.TextBox txtDescricaoPosicao 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   60
            MaxLength       =   100
            TabIndex        =   5
            Top             =   450
            Width           =   5385
         End
         Begin VB.CheckBox chkAtibo 
            Caption         =   "Ativo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5580
            TabIndex        =   4
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton cmdAlterar 
            Appearance      =   0  'Flat
            Caption         =   "Alterar"
            Height          =   405
            Left            =   1560
            TabIndex        =   3
            Top             =   930
            Width           =   1395
         End
         Begin VB.CommandButton cmdAdicionar 
            Appearance      =   0  'Flat
            Caption         =   "Adicionar"
            Height          =   405
            Left            =   30
            Picture         =   "frmPosicao.frx":500A
            TabIndex        =   2
            Top             =   930
            Width           =   1395
         End
         Begin VB.Label Label 
            Caption         =   "Descricao"
            Height          =   285
            Left            =   90
            TabIndex        =   6
            Top             =   240
            Width           =   735
         End
      End
      Begin TrueOleDBGrid80.TDBGrid ssgPosicao 
         Height          =   3855
         Left            =   30
         TabIndex        =   7
         Top             =   1590
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6800
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Ativo"
         Columns(0).DataField=   "Ativo_BT"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Posição"
         Columns(1).DataField=   "Posicao_VC"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).Locked=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   3
         Splits(0).DividerColor=   -2147483633
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=714"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8705"
         Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=10054"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=9975"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8193"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   2
         DefColWidth     =   0
         HeadLines       =   2
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         InsertMode      =   0   'False
         DeadAreaBackColor=   -2147483633
         RowDividerColor =   -2147483633
         RowSubDividerColor=   -2147483633
         DirectionAfterEnter=   2
         DirectionAfterTab=   1
         MaxRows         =   250000
         ChildGrid       =   "ssgEventoCentroCusto"
         ChildGrid.vt    =   8
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H0&"
         _StyleDefs(7)   =   ":id=1,.borderColor=&HFFFF&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(8)   =   ":id=1,.strikethrough=0,.charset=0"
         _StyleDefs(9)   =   ":id=1,.fontname=Arial"
         _StyleDefs(10)  =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&HE3DFE0&,.fgcolor=&H0&"
         _StyleDefs(11)  =   ":id=4,.borderColor=&HFFFFFF&"
         _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8000000F&,.fgcolor=&H0&"
         _StyleDefs(13)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(14)  =   ":id=2,.fontname=Arial"
         _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.borderColor=&HFFFFFF&"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H808080&,.fgcolor=&H0&"
         _StyleDefs(18)  =   ":id=6,.borderColor=&H8080&"
         _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HCCFFFF&,.borderColor=&HFFFFFF&,.bold=0"
         _StyleDefs(20)  =   ":id=7,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(21)  =   ":id=7,.fontname=Arial"
         _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFFF00&"
         _StyleDefs(24)  =   ":id=9,.borderColor=&HFFFFFF&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(25)  =   ":id=9,.strikethrough=0,.charset=0"
         _StyleDefs(26)  =   ":id=9,.fontname=Arial"
         _StyleDefs(27)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.borderColor=&HFFFFFF&"
         _StyleDefs(28)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(29)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42,.borderColor=&HFFFF&"
         _StyleDefs(30)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&H80000005&"
         _StyleDefs(31)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HC0C0C0&"
         _StyleDefs(32)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H0&"
         _StyleDefs(33)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(34)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(35)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&H800000&,.fgcolor=&HFFFFFF&"
         _StyleDefs(36)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(37)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(38)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(39)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(40)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(41)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(42)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2,.wraptext=-1,.locked=-1"
         _StyleDefs(43)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
         _StyleDefs(44)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=2,.locked=-1"
         _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(50)  =   "Named:id=33:Normal"
         _StyleDefs(51)  =   ":id=33,.parent=0"
         _StyleDefs(52)  =   "Named:id=34:Heading"
         _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   ":id=34,.wraptext=-1"
         _StyleDefs(55)  =   "Named:id=35:Footing"
         _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   "Named:id=36:Selected"
         _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=37:Caption"
         _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(61)  =   "Named:id=38:HighlightRow"
         _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(63)  =   "Named:id=39:EvenRow"
         _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(65)  =   "Named:id=40:OddRow"
         _StyleDefs(66)  =   ":id=40,.parent=33"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33"
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosicao.frx":5674
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosicao.frx":5C0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosicao.frx":61A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosicao.frx":6742
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosicao.frx":6CDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosicao.frx":7276
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosicao.frx":7810
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosicao.frx":7DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosicao.frx":8344
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosicao.frx":88DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbBotoes 
      Height          =   570
      Left            =   5760
      TabIndex        =   8
      Top             =   5580
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1005
      ButtonWidth     =   1296
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F10-Sair"
            Key             =   "cmdSair"
            Object.ToolTipText     =   "Sair da tela"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   9
      Top             =   6195
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   397
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Image imgcima 
      Height          =   135
      Left            =   180
      Picture         =   "frmPosicao.frx":8E78
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgbaixo 
      Height          =   135
      Left            =   0
      Picture         =   "frmPosicao.frx":8FFE
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
End
Attribute VB_Name = "frmPosicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mobjRsPosicoes As Recordset

Private Sub cmdAdicionar_Click()
    AdicionarAlterarPosicao txtDescricaoPosicao.Text, IIf(chkAtibo.Value = vbChecked, True, False), 1
    CarregarPosicoes
End Sub

Private Sub cmdAlterar_Click()
    If cmdAlterar.Caption = "Alterar" Then
        cmdAlterar.Caption = "Gravar"
        cmdAdicionar.Enabled = False
        txtDescricaoPosicao.Text = mobjRsPosicoes!Posicao_VC
        chkAtibo.Value = IIf(mobjRsPosicoes!Ativo_BT = True, vbChecked, vbUnchecked)
        ssgPosicao.Enabled = False
    Else
        AdicionarAlterarPosicao txtDescricaoPosicao.Text, IIf(chkAtibo.Value = vbChecked, True, False), 2
        cmdAlterar.Caption = "Alterar"
        cmdAdicionar.Enabled = True
        ssgPosicao.Enabled = True
        CarregarPosicoes
    End If
End Sub

Private Sub Form_Load()
    CarregarPosicoes
    
    sta.Panels(1).Text = gSMConexao.LoginUsuario
    sta.Panels(1).Width = frmPosicao.Width / 3
    sta.Panels(2).Text = gSMConexao.NomeBaseDados
    sta.Panels(2).Width = frmPosicao.Width / 3
    sta.Panels(3).Text = gSMConexao.NomeServidor
    sta.Panels(3).Width = frmPosicao.Width / 3
End Sub

Private Sub ssgPosicao_Click()
10    On Error Resume Next
20        ssgPosicao.SelBookmarks.Clear
30        ssgPosicao.SelBookmarks.Add ssgPosicao.Bookmark
40    On Error GoTo 0

End Sub

Private Sub ssgPosicao_HeadClick(ByVal ColIndex As Integer)
10        OrdenarColunaTrueDB ssgPosicao, ColIndex, imgcima, imgbaixo
End Sub

Private Sub ssgPosicao_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
10        ssgPosicao_Click
End Sub

Private Sub tbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
10        Select Case Button.Key
              Case "cmdSair"
20                Unload Me
30        End Select
End Sub

Private Sub CarregarPosicoes()
10    On Error GoTo Erro
            
20        modManutencao_SelecionarPosicoesAtleta mobjRsPosicoes, , False
30        ssgPosicao.DataSource = mobjRsPosicoes

40    Exit Sub
Erro:
50       Call MsgBox("Erro no módulo: " & "frmPosicao" & vbCrLf & "CarregarPosicoes" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Private Sub AdicionarAlterarPosicao(strPosicao As String, blnAtivo As Boolean, lngOperacao As Long)
10    On Error GoTo Erro
            
          '1 - Adicionar
          '2 - Alterar
20        Select Case lngOperacao
                  Case 1
30                    modManutencao_AdicionarAlterarPosicao strPosicao, blnAtivo
                  
40                Case 2
50                    modManutencao_AdicionarAlterarPosicao strPosicao, blnAtivo, mobjRsPosicoes!POSICAO_IN
60        End Select
          
70    Exit Sub
Erro:
80       Call MsgBox("Erro no módulo: " & "frmPosicao" & vbCrLf & "AdicionarAlterarPosicao" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")



End Sub

