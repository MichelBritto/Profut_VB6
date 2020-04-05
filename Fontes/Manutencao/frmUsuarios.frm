VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmUsuarios 
   Caption         =   "ProFut  - Usuários"
   ClientHeight    =   9420
   ClientLeft      =   6690
   ClientTop       =   2265
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   10590
   Begin VB.Frame fraPermissao 
      Height          =   8835
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   10515
      Begin VB.Frame fraCadastro 
         Caption         =   "Novo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   60
         TabIndex        =   4
         Top             =   120
         Width           =   10425
         Begin VB.CommandButton cmdAlterarUsuario 
            Appearance      =   0  'Flat
            Caption         =   "Alterar"
            Height          =   405
            Left            =   8850
            TabIndex        =   16
            Top             =   1080
            Width           =   1395
         End
         Begin VB.CommandButton cmdAdicionarUsuário 
            Appearance      =   0  'Flat
            Caption         =   "Adicionar"
            Height          =   405
            Left            =   7410
            Picture         =   "frmUsuarios.frx":0000
            TabIndex        =   15
            Top             =   1080
            Width           =   1395
         End
         Begin EditLib.fpMask fpTelefone 
            Height          =   405
            Left            =   4830
            TabIndex        =   12
            Top             =   1110
            Width           =   2475
            _Version        =   196608
            _ExtentX        =   4366
            _ExtentY        =   714
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   0
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   0
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   1
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            AllowOverflow   =   0   'False
            BestFit         =   0   'False
            ClipMode        =   0
            DataFormatEx    =   0
            Mask            =   "( __ ) ____-_____"
            PromptChar      =   "_"
            PromptInclude   =   0   'False
            RequireFill     =   0   'False
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   1
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            AutoTab         =   0   'False
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.TextBox txtEmail 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   60
            MaxLength       =   6
            TabIndex        =   9
            Top             =   1110
            Width           =   4695
         End
         Begin VB.TextBox txtLogin 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   4830
            MaxLength       =   6
            TabIndex        =   7
            Top             =   450
            Width           =   2475
         End
         Begin VB.TextBox txtNome 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   60
            MaxLength       =   6
            TabIndex        =   5
            Top             =   450
            Width           =   4695
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscCargo 
            Height          =   390
            Left            =   7350
            TabIndex        =   13
            Top             =   450
            Width           =   2175
            DataFieldList   =   "Column 0"
            BevelType       =   0
            _Version        =   196617
            DataMode        =   2
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorOdd    =   15724527
            RowHeight       =   476
            Columns(0).Width=   3200
            Columns(0).DataType=   8
            Columns(0).FieldLen=   4096
            _ExtentX        =   3836
            _ExtentY        =   688
            _StockProps     =   93
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label29 
            Caption         =   "Cargo"
            Height          =   285
            Left            =   7380
            TabIndex        =   14
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label3 
            Caption         =   "Telefone/Celular"
            Height          =   285
            Left            =   4860
            TabIndex        =   11
            Top             =   900
            Width           =   1365
         End
         Begin VB.Label Label2 
            Caption         =   "E-mail"
            Height          =   285
            Left            =   90
            TabIndex        =   10
            Top             =   900
            Width           =   3135
         End
         Begin VB.Label Label1 
            Caption         =   "Login"
            Height          =   285
            Left            =   4860
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label 
            Caption         =   "Nome"
            Height          =   285
            Left            =   90
            TabIndex        =   6
            Top             =   240
            Width           =   555
         End
      End
      Begin VB.Frame fraUsuários 
         Caption         =   "Usuários"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7035
         Left            =   60
         TabIndex        =   2
         Top             =   1770
         Width           =   10425
         Begin TrueOleDBGrid80.TDBGrid ssgJogadoresEquipe 
            Height          =   6765
            Left            =   60
            TabIndex        =   3
            Top             =   210
            Width           =   10275
            _ExtentX        =   18124
            _ExtentY        =   11933
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nome"
            Columns(0).DataField=   "Apelido"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Login"
            Columns(1).DataField=   "Camisa"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Cargo"
            Columns(2).DataField=   "Cartegoria"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Telefone"
            Columns(3).DataField=   "Nome"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   16
            Columns(4)._MaxComboItems=   5
            Columns(4).ValueItems(0)._DefaultItem=   0
            Columns(4).ValueItems(0).Value=   "1"
            Columns(4).ValueItems(0).Value.vt=   8
            Columns(4).ValueItems(0).DisplayValue=   "MASCULINO"
            Columns(4).ValueItems(0).DisplayValue.vt=   8
            Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(4).ValueItems(1)._DefaultItem=   0
            Columns(4).ValueItems(1).Value=   "2"
            Columns(4).ValueItems(1).Value.vt=   8
            Columns(4).ValueItems(1).DisplayValue=   "FEMINIMO"
            Columns(4).ValueItems(1).DisplayValue.vt=   8
            Columns(4).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(4).ValueItems(2)._DefaultItem=   0
            Columns(4).ValueItems(2).Value=   "-1"
            Columns(4).ValueItems(2).Value.vt=   8
            Columns(4).ValueItems(2).DisplayValue=   "-1"
            Columns(4).ValueItems(2).DisplayValue.vt=   8
            Columns(4).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
            Columns(4).ValueItems.Count=   3
            Columns(4).Caption=   "E-mail"
            Columns(4).DataField=   "Sexo"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=4498"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4419"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=3307"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3228"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=1"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2699"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2619"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=2910"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2831"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=1"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=5847"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=5768"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=1"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
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
            _StyleDefs(46)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(51)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14,.alignment=2"
            _StyleDefs(52)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(55)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(59)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(62)  =   "Named:id=33:Normal"
            _StyleDefs(63)  =   ":id=33,.parent=0"
            _StyleDefs(64)  =   "Named:id=34:Heading"
            _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(66)  =   ":id=34,.wraptext=-1"
            _StyleDefs(67)  =   "Named:id=35:Footing"
            _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(69)  =   "Named:id=36:Selected"
            _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(71)  =   "Named:id=37:Caption"
            _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(73)  =   "Named:id=38:HighlightRow"
            _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(75)  =   "Named:id=39:EvenRow"
            _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(77)  =   "Named:id=40:OddRow"
            _StyleDefs(78)  =   ":id=40,.parent=33"
            _StyleDefs(79)  =   "Named:id=41:RecordSelector"
            _StyleDefs(80)  =   ":id=41,.parent=34"
            _StyleDefs(81)  =   "Named:id=42:FilterBar"
            _StyleDefs(82)  =   ":id=42,.parent=33"
         End
      End
   End
   Begin MSComctlLib.Toolbar tbBotoes 
      Height          =   570
      Left            =   7710
      TabIndex        =   1
      Top             =   8820
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   1005
      ButtonWidth     =   1720
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F3 - Alterar"
            Key             =   "cmdAlterar"
            Object.ToolTipText     =   "Alterar Jogador"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "F7-Gravar"
            Key             =   "cmdGravar"
            Object.ToolTipText     =   "Gravar Alterações"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F10-Sair"
            Key             =   "cmdSair"
            Object.ToolTipText     =   "Sair da tela"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   90
      Top             =   8850
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
            Picture         =   "frmUsuarios.frx":066A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":0C04
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":119E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":1738
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":1CD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":226C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":2806
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":2DA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":333A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":38D4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
