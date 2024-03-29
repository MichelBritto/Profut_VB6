VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmCadastroDeEquipe 
   Caption         =   "ProFut - Cadastro de Equipe"
   ClientHeight    =   7590
   ClientLeft      =   6210
   ClientTop       =   1695
   ClientWidth     =   6585
   Icon            =   "frmCadastroDeEquipe.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   6585
   Begin VB.Frame fraPrincipal 
      Height          =   6705
      Left            =   0
      TabIndex        =   14
      Top             =   -60
      Width           =   6585
      Begin VB.Frame fraJogadores 
         Caption         =   "Jogadores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   2640
         TabIndex        =   21
         Top             =   120
         Width           =   3915
         Begin TrueOleDBGrid80.TDBGrid ssgJogadoresEquipe 
            Height          =   2865
            Left            =   60
            TabIndex        =   30
            Top             =   210
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   5054
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Apelido"
            Columns(0).DataField=   "Apelido"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Cartegoria"
            Columns(1).DataField=   "Cartegoria"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3413"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3334"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=513"
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
            _StyleDefs(42)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2,.wraptext=-1"
            _StyleDefs(43)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
            _StyleDefs(44)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
            _StyleDefs(48)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
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
      Begin VB.Frame fraInfoClube 
         Caption         =   "Informa��es do Clube"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   30
         TabIndex        =   15
         Top             =   3180
         Width           =   6525
         Begin VB.TextBox txtUsuarioAlteracao 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   4440
            TabIndex        =   10
            Top             =   2310
            Width           =   1995
         End
         Begin VB.TextBox txtEmailResponsavel 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   120
            TabIndex        =   5
            Top             =   1680
            Width           =   6315
         End
         Begin VB.CheckBox chkWpp2 
            Height          =   195
            Left            =   2520
            Picture         =   "frmCadastroDeEquipe.frx":038A
            TabIndex        =   9
            Top             =   3030
            Width           =   255
         End
         Begin VB.PictureBox wpp2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   2790
            Picture         =   "frmCadastroDeEquipe.frx":0A9E
            ScaleHeight     =   345
            ScaleWidth      =   315
            TabIndex        =   24
            Top             =   2970
            Width           =   315
         End
         Begin VB.CheckBox chkWpp1 
            Height          =   195
            Left            =   2520
            Picture         =   "frmCadastroDeEquipe.frx":0F54
            TabIndex        =   7
            Top             =   2400
            Width           =   255
         End
         Begin VB.PictureBox wpp1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   2790
            Picture         =   "frmCadastroDeEquipe.frx":1668
            ScaleHeight     =   345
            ScaleWidth      =   315
            TabIndex        =   23
            Top             =   2340
            Width           =   315
         End
         Begin VB.TextBox txtTelefoneCelular2 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   120
            TabIndex        =   8
            Top             =   2910
            Width           =   2235
         End
         Begin VB.TextBox txtCodigoInterno 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   120
            MaxLength       =   6
            TabIndex        =   1
            Top             =   420
            Width           =   945
         End
         Begin VB.TextBox txtSiglaEquipe 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   5610
            MaxLength       =   3
            TabIndex        =   3
            Top             =   420
            Width           =   825
         End
         Begin VB.TextBox txtResponsavel 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   120
            TabIndex        =   4
            Top             =   1050
            Width           =   6315
         End
         Begin VB.TextBox txtTelefoneCelular1 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   120
            TabIndex        =   6
            Top             =   2310
            Width           =   2235
         End
         Begin VB.TextBox txtNomeEquipe 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   1110
            TabIndex        =   2
            Top             =   420
            Width           =   4455
         End
         Begin SSCalendarWidgets_A.SSDateCombo dtcDataUltimaAlteracao 
            Height          =   405
            Left            =   4440
            TabIndex        =   13
            Top             =   2940
            Width           =   1995
            _Version        =   65543
            _ExtentX        =   3519
            _ExtentY        =   714
            _StockProps     =   93
            Format          =   "DD MM,YYYY"
            BevelType       =   0
            Mask            =   2
         End
         Begin VB.Label Label26 
            Caption         =   "Usu�rio Ultima Altera��o"
            Height          =   285
            Left            =   4440
            TabIndex        =   27
            Top             =   2100
            Width           =   1965
         End
         Begin VB.Label Label27 
            Caption         =   "Data Ultima altera��o"
            Height          =   285
            Left            =   4440
            TabIndex        =   26
            Top             =   2730
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "E-mail Respons�vel"
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Top             =   1470
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Telefone/Celular"
            Height          =   285
            Left            =   90
            TabIndex        =   22
            Top             =   2730
            Width           =   1485
         End
         Begin VB.Label Label 
            Caption         =   "C�digo"
            Height          =   285
            Left            =   150
            TabIndex        =   20
            Top             =   210
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Nome Equipe"
            Height          =   285
            Left            =   1110
            TabIndex        =   19
            Top             =   210
            Width           =   1485
         End
         Begin VB.Label labrl 
            Caption         =   "Sigla"
            Height          =   285
            Left            =   5610
            TabIndex        =   18
            Top             =   210
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Respons�vel da Equipe"
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label5 
            Caption         =   "Telefone/Celular"
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   2100
            Width           =   1485
         End
      End
      Begin Threed.SSCommand cmdRemover 
         Height          =   330
         Left            =   1320
         TabIndex        =   12
         Top             =   2820
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   196609
         PictureFrames   =   1
         BackStyle       =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmCadastroDeEquipe.frx":1B1E
         Caption         =   "   Rem Foto"
         ButtonStyle     =   3
         PictureAlignment=   1
      End
      Begin Threed.SSCommand cmdAdicionar 
         Height          =   330
         Left            =   60
         TabIndex        =   11
         Top             =   2820
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   196609
         PictureFrames   =   1
         BackStyle       =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmCadastroDeEquipe.frx":1E40
         Caption         =   "        Add Foto"
         ButtonStyle     =   3
         PictureAlignment=   1
      End
      Begin Threed.SSFrame SSFrame 
         Height          =   2565
         Index           =   1
         Left            =   60
         TabIndex        =   28
         Top             =   210
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   4524
         _Version        =   196609
         Begin VB.Image imgClube 
            Height          =   2415
            Left            =   60
            Stretch         =   -1  'True
            Top             =   60
            Width           =   2430
         End
      End
   End
   Begin MSComctlLib.Toolbar tbBotoes 
      Height          =   570
      Left            =   30
      TabIndex        =   0
      Top             =   6720
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   1005
      ButtonWidth     =   2355
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F2 - Novo"
            Key             =   "cmdNovo"
            Object.ToolTipText     =   "Novo Jogador"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "F3 - Alterar"
            Key             =   "cmdAlterar"
            Object.ToolTipText     =   "Alterar Jogador"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "F6 - Abandonar"
            Key             =   "cmdLimpar"
            Object.ToolTipText     =   "Limpar dados da tela"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "F7-Gravar"
            Key             =   "cmdGravar"
            Object.ToolTipText     =   "Gravar Altera��es"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F10-Sair"
            Key             =   "cmdSair"
            Object.ToolTipText     =   "Sair da tela"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   30
      Top             =   4680
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
            Picture         =   "frmCadastroDeEquipe.frx":2552
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeEquipe.frx":2AEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeEquipe.frx":3086
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeEquipe.frx":3620
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeEquipe.frx":3BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeEquipe.frx":4154
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeEquipe.frx":46EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeEquipe.frx":4C88
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeEquipe.frx":5222
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeEquipe.frx":57BC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   29
      Top             =   7350
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4233
            MinWidth        =   4233
            Text            =   "CM Software"
            TextSave        =   "CM Software"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4260
            Text            =   "CadJogador"
            TextSave        =   "CadJogador"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Alpha 0.1"
            TextSave        =   "Alpha 0.1"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCadastroDeEquipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstrFlag As String
Dim mstrFoto As String
Dim mobjRsEquipe As Recordset
Dim mobjRsJogadores As Recordset

Public Property Let DiretorioFotoEquipe(strDiretorio As String)
    mstrFoto = strDiretorio
End Property

Private Sub CriarEPreencherRecordsetJogadores(ByRef objRsJogadores As Recordset)
On Error GoTo Erro
      
    Set mobjRsJogadores = Nothing
    Set mobjRsJogadores = New Recordset

    
    With mobjRsJogadores
        .Fields.Append "Apelido", adVarChar, 1024
        .Fields.Append "Cartegoria", adVarChar, 1024
        .CursorLocation = adUseClient
        .Open , Nothing, adOpenDynamic, adLockOptimistic
    End With
    
    If Not objRsJogadores Is Nothing Then
        If Not objRsJogadores.BOF And Not objRsJogadores.EOF Then
            objRsJogadores.MoveFirst
            Do While Not objRsJogadores.EOF
                mobjRsJogadores.AddNew
                
                mobjRsJogadores!Apelido = NS(objRsJogadores!APELIDO_VC)
                mobjRsJogadores!Cartegoria = NS(objRsJogadores!DESCRICAO_VC)
                
                mobjRsJogadores.Update
                objRsJogadores.MoveNext
            Loop
        End If
    End If
    
    ssgJogadoresEquipe.DataSource = mobjRsJogadores
    ssgJogadoresEquipe.Update

Exit Sub
Erro:
   Call MsgBox("Erro no m�dulo: " & "frmCadastroDeEquipe" & vbCrLf & "No Procedimento: " & "CriarEPreencherRecordsetJogadores" & vbCrLf & "Descri��o: " & Err.Description & vbCrLf & "N�mero: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Aten��o!")

End Sub

Private Sub cmdAdicionar_Click()
    On Error GoTo Erro
    
    frmAdicionarFotoEquipe.Show vbModal
  
    If mstrFoto <> "" Then
        imgClube.Picture = Nothing
        imgClube.Stretch = True
        imgClube.Picture = LoadPicture(mstrFoto)
    End If
    
'
'              modJogador_AdicionarAlterarFotoJogador udtJogador.lngCodigo
'
'    Call FileCopy(txtFoto.Text, "C:\Program Files\TesteDirPadrao")
    Exit Sub
Erro:
 Call MsgBox("Erro no m�dulo: " & "frmCadastroDeEquipe" & vbCrLf & "No Procedimento: " & "cmdAdicionar_Click" & vbCrLf & "Descri��o: " & Err.Description & vbCrLf & "N�mero: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Aten��o!")
End Sub

Private Sub cmdRemover_Click()
    imgClube.Picture = Nothing
    mstrFoto = ""
End Sub

Private Sub Form_Load()
    Sta.Panels(1).Text = gSMConexao.LoginUsuario
    Sta.Panels(2).Text = gSMConexao.NomeBaseDados
    
    mstrFlag = ""
    
    Call LimparCampos
    Call HabilitarCampos(False)
End Sub


Private Sub tbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Not (Button.Enabled) Then Exit Sub
    Select Case Button.Key
        
        Case "cmdNovo":
            mstrFlag = "I"
            Call LimparCampos
            Call HabilitarCampos(True)
            Call HabilitarTBBotoes(False, False, True, True, False)
        
        Case "cmdAlterar":
            mstrFlag = "A"
            Call HabilitarCampos(True)
            Call HabilitarTBBotoes(False, False, True, True, False)
        
        Case "cmdLimpar":
            mstrFlag = ""
            Call LimparCampos
            Call HabilitarCampos(False)
            Call HabilitarTBBotoes(True, False, False, False, True)
        
        Case "cmdGravar"
            'mstrFlag = ""
            If VerificarCampos Then
                GravarEquipe
                CarregarEquipe Val(txtCodigoInterno.Text)
                mstrFlag = "V"
            Else: Exit Sub
            End If
            Call HabilitarCampos(False)
            Call HabilitarTBBotoes(False, True, True, False, False)
        Case "cmdSair"
            Unload Me
        
    End Select
End Sub

Private Sub LimparCampos()
    txtCodigoInterno.Text = ""
    txtNomeEquipe.Text = ""
    txtSiglaEquipe.Text = ""
    txtEmailResponsavel.Text = ""
    txtResponsavel.Text = ""
    txtTelefoneCelular1.Text = ""
    txtTelefoneCelular2.Text = ""
    txtUsuarioAlteracao.Text = ""
    
    dtcDataUltimaAlteracao.DateValue = Empty
    
    chkWpp1.Value = vbUnchecked
    chkWpp2.Value = vbUnchecked
    
    imgClube.Picture = Nothing
    mstrFoto = ""
    
    ssgJogadoresEquipe.DataSource = Nothing
    ssgJogadoresEquipe.Update
End Sub

Private Sub HabilitarCampos(blnHabilitar As Boolean)

    If mstrFlag = "I" Or mstrFlag = "A" Then
        txtCodigoInterno.Locked = True
    Else
        txtCodigoInterno.Locked = fase
    End If
    txtNomeEquipe.Locked = Not blnHabilitar
    txtSiglaEquipe.Locked = Not blnHabilitar
    txtResponsavel.Locked = Not blnHabilitar
    txtEmailResponsavel.Locked = Not blnHabilitar
    txtTelefoneCelular1.Locked = Not blnHabilitar
    txtTelefoneCelular2.Locked = Not blnHabilitar
    txtUsuarioAlteracao.Locked = True
    
    dtcDataUltimaAlteracao.Enabled = False
    
    chkWpp1.Enabled = blnHabilitar
    chkWpp2.Enabled = blnHabilitar
    
    cmdAdicionar.Enabled = blnHabilitar
    cmdRemover.Enabled = blnHabilitar
End Sub
Private Sub HabilitarTBBotoes(blnNovo As Boolean, blnAlterar As Boolean, blnAbandonar As Boolean, blnGravar As Boolean, blnSair As Boolean)

    tbBotoes.Buttons("cmdNovo").Enabled = blnNovo
    tbBotoes.Buttons("cmdAlterar").Enabled = blnAlterar
    tbBotoes.Buttons("cmdLimpar").Enabled = blnAbandonar
    tbBotoes.Buttons("cmdGravar").Enabled = blnGravar
    tbBotoes.Buttons("cmdSair").Enabled = blnSair
    
End Sub


Private Function VerificarCampos()
On Error GoTo Erro
    Dim blnContinua As Boolean
    Dim strMensagem As String
    
    blnContinua = True
    
    If txtNomeEquipe.Text = "" Then
        strMensagem = strMensagem & "-> Nome da equipe n�o preenchido." & vbCrLf
        blnContinua = False
    End If
    
    If txtSiglaEquipe.Text = "" Then
        strMensagem = strMensagem & "-> Sigla da equipe n�o preenchido." & vbCrLf
        blnContinua = False
    End If
    
    If txtResponsavel.Text = "" Then
        strMensagem = strMensagem & "-> Respons�vel da equipe n�o preenchido." & vbCrLf
        blnContinua = False
    End If
    
    If txtEmailResponsavel.Text = "" Then
        strMensagem = strMensagem & "-> E-mail do respons�vel da equipe n�o preenchido." & vbCrLf
        blnContinua = False
    End If
    
    If txtTelefoneCelular1.Text = "" And txtTelefoneCelular2.Text = "" Then
        strMensagem = strMensagem & "-> � necess�rio ter pelo menos um n�mero de contato do respons�vel." & vbCrLf
        blnContinua = False
    End If
    
    If Not blnContinua Then
        MsgBox "O jogador n�o pode ser gravado pois possu� as seguintes pend�ncias: " & vbCrLf & strMensagem, vbOKOnly + vbInformation, "Aten��o!"
    End If
    
    VerificarCampos = blnContinua

Exit Function
Erro:
   VerificarCampos = False
   Call MsgBox("Erro no m�dulo: " & "frmCadastroDeEquipe" & vbCrLf & "No Procedimento: " & "VerificarCampos" & vbCrLf & "Descri��o: " & Err.Description & vbCrLf & "N�mero: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Aten��o!")

End Function
Private Sub GravarEquipe()
On Error GoTo Erro
Dim udtEquipe As TypEquipe
    
    With udtEquipe
        .strNome = txtNomeEquipe.Text
        .strSigla = txtSiglaEquipe.Text
        .strResponsavel = txtResponsavel.Text
        .strEmailContato = txtEmailResponsavel.Text
        .strContato1 = txtTelefoneCelular1.Text
        .blnWpp1 = IIf(chkWpp1.Value = vbChecked, True, False)
        .blnWpp2 = IIf(chkWpp2.Value = vbChecked, True, False)
        .strContato2 = txtTelefoneCelular2.Text
        .strEnderecoImagem = mstrFoto
    End With
    
    If mstrFlag = "I" Then
        Call modEquipe_AdicionarEquipe(udtEquipe)
        txtCodigoInterno.Text = udtEquipe.lngCodigo
    ElseIf mstrFlag = "A" Then
        udtEquipe.lngCodigo = Val(txtCodigoInterno.Text)
        Call modEquipe_AlterarEquipe(udtEquipe)
    End If
    
    

Exit Sub
Erro:
   Call MsgBox("Erro no m�dulo: " & "frmCadastroDeEquipe" & vbCrLf & "No Procedimento: " & "GravarEquipe" & vbCrLf & "Descri��o: " & Err.Description & vbCrLf & "N�mero: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Aten��o!")

End Sub

Private Sub CarregarEquipe(lngCodigo As Long)
On Error GoTo Erro
Dim objrs As Recordset

    Dim objRsEquipe As Recordset
    Set objRsEquipe = New Recordset
      
    Call LimparCampos
    modEquipe_SelecionarEquipePorCodigo lngCodigo, objRsEquipe
    
    If Not objRsEquipe Is Nothing Then
        If Not objRsEquipe.EOF And Not objRsEquipe.BOF Then
            
            txtNomeEquipe.Text = NS(objRsEquipe!NOME_VC)
            txtSiglaEquipe.Text = NS(objRsEquipe!SIGLA_VC)
            txtResponsavel.Text = NS(objRsEquipe!RESPONSAVEL_VC)
            txtEmailResponsavel.Text = NS(objRsEquipe!EMAILCONTATO_VC)
            txtTelefoneCelular1.Text = NS(objRsEquipe!CONTATO1_VC)
            txtTelefoneCelular2.Text = NS(objRsEquipe!CONTATO2_VC)
            txtUsuarioAlteracao.Text = NS(objRsEquipe!USUARIOULTIMAALTERACAO_VC)
            txtCodigoInterno.Text = lngCodigo
            dtcDataUltimaAlteracao.DateValue = ND(objRsEquipe!DATAULTIMAALTERACAO_DT)
            
            chkWpp1.Value = IIf(NB(objRsEquipe!WHATSAPP1_BT), vbChecked, vbUnchecked)
            chkWpp2.Value = IIf(NB(objRsEquipe!WHATSAP2_BT), vbChecked, vbUnchecked)
    
            mstrFoto = NS(objRsEquipe!ENDERECOIMAGEM_VC)
            If mstrFoto <> "" Then
                imgClube.Picture = Nothing
                imgClube.Stretch = True
                On Error Resume Next
                imgClube.Picture = LoadPicture(mstrFoto)
                On Error GoTo Erro
            End If
            
            Set objrs = objRsEquipe.NextRecordset
            
            CriarEPreencherRecordsetJogadores objrs
            
            mstrFlag = ""
            Call HabilitarCampos(False)
            Call HabilitarTBBotoes(False, True, True, False, False)
            
        Else
            MsgBox "Equipe n�o encontrada ou c�digo inv�lido.", vbOKOnly + vbInformation, "Aten��o!"
        End If
    Else
        MsgBox "Equipe n�o encontrada ou c�digo inv�lido.", vbOKOnly + vbInformation, "Aten��o!"
    End If

      

Exit Sub
Erro:
   Call MsgBox("Erro no m�dulo: " & "frmCadastroDeEquipe" & vbCrLf & "No Procedimento: " & "CarregarEquipe" & vbCrLf & "Descri��o: " & Err.Description & vbCrLf & "N�mero: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Aten��o!")

End Sub

Private Sub txtCodigoInterno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call CarregarEquipe(Val(txtCodigoInterno.Text))
    End If
End Sub
