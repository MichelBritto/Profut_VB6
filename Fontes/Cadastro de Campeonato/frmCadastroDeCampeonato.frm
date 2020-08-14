VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6330
   ClientLeft      =   6015
   ClientTop       =   2925
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   10860
   Begin VB.Frame fraPrincipal 
      Height          =   5685
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   10815
      Begin VB.Frame fraCartegoria 
         Appearance      =   0  'Flat
         Caption         =   "Cartegoria(s) Aceitas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   120
         TabIndex        =   25
         Top             =   3000
         Width           =   2895
         Begin TrueOleDBGrid80.TDBGrid ssgEquipes 
            Height          =   2175
            Left            =   60
            TabIndex        =   26
            Top             =   300
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   3836
            _LayoutType     =   4
            _RowHeight      =   15
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   100
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   "marcado_BT"
            Columns(0).DefaultValue=   "0"
            Columns(0).DefaultValue.vt=   8
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Cartegoria"
            Columns(1).DataField=   "NOME_VC"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   1
            Splits(0).MarqueeStyle=   5
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AllowColSelect=   0   'False
            Splits(0).DividerColor=   -2147483633
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=450"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=370"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=1"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=4286"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=4207"
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
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            InsertMode      =   0   'False
            MultiSelect     =   0
            DeadAreaBackColor=   -2147483633
            RowDividerColor =   -2147483633
            RowSubDividerColor=   -2147483633
            DirectionAfterEnter=   1
            DirectionAfterTab=   1
            MaxRows         =   250000
            ChildGrid       =   "ssgObrasEnderecos"
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
            _StyleDefs(7)   =   ":id=1,.borderColor=&HFFFFFF&,.bold=0,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(8)   =   ":id=1,.strikethrough=0,.charset=0"
            _StyleDefs(9)   =   ":id=1,.fontname=Arial"
            _StyleDefs(10)  =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&H0&,.fgcolor=&HFFFFFF&"
            _StyleDefs(11)  =   ":id=4,.appearance=0,.borderSize=0,.borderColor=&HFFFFFF&,.borderType=0,.bold=-1"
            _StyleDefs(12)  =   ":id=4,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(13)  =   ":id=4,.fontname=Arial"
            _StyleDefs(14)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8000000F&,.fgcolor=&H0&"
            _StyleDefs(15)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(16)  =   ":id=2,.fontname=Arial"
            _StyleDefs(17)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H808080&,.fgcolor=&H0&"
            _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1,.borderColor=&HFFFFFF&,.bold=0,.fontsize=825"
            _StyleDefs(21)  =   ":id=7,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(22)  =   ":id=7,.fontname=Arial"
            _StyleDefs(23)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(24)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFFF00&"
            _StyleDefs(25)  =   ":id=9,.borderColor=&HFFFFFF&,.bold=0,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(26)  =   ":id=9,.strikethrough=0,.charset=0"
            _StyleDefs(27)  =   ":id=9,.fontname=Arial"
            _StyleDefs(28)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(29)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(30)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(31)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(32)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HC0C0C0&"
            _StyleDefs(33)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H0&"
            _StyleDefs(34)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(35)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(36)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&H800000&,.fgcolor=&HFFFFFF&"
            _StyleDefs(37)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(38)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(39)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(40)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(41)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(42)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(43)  =   "Splits(0).Columns(0).Style:id=102,.parent=13,.alignment=2,.locked=0"
            _StyleDefs(44)  =   "Splits(0).Columns(0).HeadingStyle:id=99,.parent=14"
            _StyleDefs(45)  =   "Splits(0).Columns(0).FooterStyle:id=100,.parent=15"
            _StyleDefs(46)  =   "Splits(0).Columns(0).EditorStyle:id=101,.parent=17"
            _StyleDefs(47)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(48)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(50)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(51)  =   "Named:id=33:Normal"
            _StyleDefs(52)  =   ":id=33,.parent=0"
            _StyleDefs(53)  =   "Named:id=34:Heading"
            _StyleDefs(54)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(55)  =   ":id=34,.wraptext=-1"
            _StyleDefs(56)  =   "Named:id=35:Footing"
            _StyleDefs(57)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(58)  =   "Named:id=36:Selected"
            _StyleDefs(59)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(60)  =   "Named:id=37:Caption"
            _StyleDefs(61)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(62)  =   "Named:id=38:HighlightRow"
            _StyleDefs(63)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(64)  =   "Named:id=39:EvenRow"
            _StyleDefs(65)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(66)  =   "Named:id=40:OddRow"
            _StyleDefs(67)  =   ":id=40,.parent=33"
            _StyleDefs(68)  =   "Named:id=41:RecordSelector"
            _StyleDefs(69)  =   ":id=41,.parent=34"
            _StyleDefs(70)  =   "Named:id=42:FilterBar"
            _StyleDefs(71)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame fraInformacoes 
         Appearance      =   0  'Flat
         Caption         =   "Competição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2865
         Left            =   3000
         TabIndex        =   3
         Top             =   120
         Width           =   7695
         Begin VB.OptionButton optFeminino 
            Caption         =   "F"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6510
            TabIndex        =   23
            Top             =   2460
            Width           =   555
         End
         Begin VB.OptionButton optMasculino 
            Caption         =   "M"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5970
            TabIndex        =   22
            Top             =   2460
            Value           =   -1  'True
            Width           =   555
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3870
            MaxLength       =   20
            TabIndex        =   20
            Top             =   1710
            Width           =   3705
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   90
            MaxLength       =   20
            TabIndex        =   18
            Top             =   1710
            Width           =   3705
         End
         Begin VB.PictureBox wpp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   1
            Left            =   5490
            Picture         =   "frmCadastroDeCampeonato.frx":0000
            ScaleHeight     =   345
            ScaleWidth      =   315
            TabIndex        =   16
            Top             =   2400
            Width           =   315
         End
         Begin VB.CheckBox chkwpp2 
            Height          =   195
            Left            =   5250
            TabIndex        =   15
            Top             =   2490
            Width           =   225
         End
         Begin VB.TextBox txtTelCel2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2970
            MaxLength       =   11
            TabIndex        =   14
            Top             =   2370
            Width           =   2235
         End
         Begin VB.PictureBox wpp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   2550
            Picture         =   "frmCadastroDeCampeonato.frx":04B6
            ScaleHeight     =   345
            ScaleWidth      =   315
            TabIndex        =   12
            Top             =   2400
            Width           =   315
         End
         Begin VB.CheckBox chkWpp1 
            Height          =   195
            Left            =   2280
            Picture         =   "frmCadastroDeCampeonato.frx":096C
            TabIndex        =   11
            Top             =   2460
            Width           =   255
         End
         Begin VB.TextBox txtTelCel1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   90
            MaxLength       =   11
            TabIndex        =   10
            Top             =   2370
            Width           =   2145
         End
         Begin VB.TextBox txtApelido 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   6
            Top             =   420
            Width           =   6465
         End
         Begin VB.TextBox txtCodigoInterno 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   90
            MaxLength       =   8
            TabIndex        =   5
            Top             =   420
            Width           =   945
         End
         Begin VB.TextBox txtNomeJogador 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   90
            MaxLength       =   128
            TabIndex        =   4
            Top             =   1080
            Width           =   7485
         End
         Begin VB.Label Label30 
            Caption         =   "Sexo"
            Height          =   285
            Left            =   5970
            TabIndex        =   24
            Top             =   2160
            Width           =   885
         End
         Begin VB.Label Label3 
            Caption         =   "Rede Social 2"
            Height          =   285
            Left            =   3900
            TabIndex        =   21
            Top             =   1500
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Rede Social"
            Height          =   285
            Left            =   90
            TabIndex        =   19
            Top             =   1500
            Width           =   1485
         End
         Begin VB.Label Label16 
            Caption         =   "Telefone/Celular 2"
            Height          =   285
            Left            =   2940
            TabIndex        =   17
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Label Label17 
            Caption         =   "Telefone/Celular"
            Height          =   285
            Left            =   90
            TabIndex        =   13
            Top             =   2130
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Responsável/Responsáveis pela Competição"
            Height          =   285
            Left            =   90
            TabIndex        =   9
            Top             =   870
            Width           =   3675
         End
         Begin VB.Label Apelido 
            Caption         =   "Nome Competição"
            Height          =   285
            Left            =   1080
            TabIndex        =   8
            Top             =   210
            Width           =   1485
         End
         Begin VB.Label Label 
            Caption         =   "Código"
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   210
            Width           =   855
         End
      End
      Begin VB.Frame fraLogo 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2865
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2865
         Begin VB.Image imgLogoCompeticao 
            Height          =   2655
            Left            =   60
            Top             =   150
            Width           =   2745
         End
      End
   End
   Begin MSComctlLib.Toolbar tbBotoes 
      Height          =   570
      Left            =   1545
      TabIndex        =   1
      Top             =   5685
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   1005
      ButtonWidth     =   2355
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
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
            Caption         =   "F4 - Procurar"
            Key             =   "cmdProcurar"
            Object.ToolTipText     =   "Procurar Jogador"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F5 - Inativar"
            Key             =   "cmdExcluir"
            Object.ToolTipText     =   "Inativar um jogador do sistema"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "F6 - Abandonar"
            Key             =   "cmdLimpar"
            Object.ToolTipText     =   "Limpar dados da tela"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "F7-Gravar"
            Key             =   "cmdGravar"
            Object.ToolTipText     =   "Gravar Alterações"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F10-Sair"
            Key             =   "cmdSair"
            Object.ToolTipText     =   "Sair da tela"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   60
      Top             =   5580
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
            Picture         =   "frmCadastroDeCampeonato.frx":1080
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":161A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":1BB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":214E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":26E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":2C82
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":321C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":37B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":3D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":42EA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
