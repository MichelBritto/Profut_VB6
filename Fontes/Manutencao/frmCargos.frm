VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPermissao 
   Caption         =   "ProFut - Cargos e Permissões"
   ClientHeight    =   8430
   ClientLeft      =   6705
   ClientTop       =   2265
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   10590
   Begin VB.Frame fraPrincipal 
      Height          =   7755
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   10605
      Begin VB.Frame fraPermissao 
         Caption         =   "Permisões"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5205
         Left            =   60
         TabIndex        =   3
         Top             =   2460
         Width           =   10485
         Begin TrueOleDBGrid80.TDBGrid ssgPermissoes 
            Height          =   4905
            Left            =   60
            TabIndex        =   4
            Top             =   210
            Width           =   10350
            _ExtentX        =   18256
            _ExtentY        =   8652
            _LayoutType     =   4
            _RowHeight      =   15
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   100
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   "check"
            Columns(0).DefaultValue=   "0"
            Columns(0).DefaultValue.vt=   8
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Código"
            Columns(1).DataField=   "Permissao_IN"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Permissão"
            Columns(2).DataField=   "Descricao_VC"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=635"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=556"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=1"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1005"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=926"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=16536"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=16457"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
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
            _StyleDefs(51)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(52)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(54)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(55)  =   "Named:id=33:Normal"
            _StyleDefs(56)  =   ":id=33,.parent=0"
            _StyleDefs(57)  =   "Named:id=34:Heading"
            _StyleDefs(58)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(59)  =   ":id=34,.wraptext=-1"
            _StyleDefs(60)  =   "Named:id=35:Footing"
            _StyleDefs(61)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(62)  =   "Named:id=36:Selected"
            _StyleDefs(63)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(64)  =   "Named:id=37:Caption"
            _StyleDefs(65)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(66)  =   "Named:id=38:HighlightRow"
            _StyleDefs(67)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(68)  =   "Named:id=39:EvenRow"
            _StyleDefs(69)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(70)  =   "Named:id=40:OddRow"
            _StyleDefs(71)  =   ":id=40,.parent=33"
            _StyleDefs(72)  =   "Named:id=41:RecordSelector"
            _StyleDefs(73)  =   ":id=41,.parent=34"
            _StyleDefs(74)  =   "Named:id=42:FilterBar"
            _StyleDefs(75)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame fraUsuarios 
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
         Height          =   2415
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   10485
         Begin TrueOleDBGrid80.TDBGrid ssgUsuarios 
            Height          =   2115
            Left            =   60
            TabIndex        =   5
            Top             =   210
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   3731
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nome"
            Columns(0).DataField=   "Nome_VC"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Login"
            Columns(1).DataField=   "Login_VC"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Cargo"
            Columns(2).DataField=   "Descricao_VC"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=8229"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=8149"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=5609"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=5530"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=1"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=3863"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=3784"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
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
            _StyleDefs(54)  =   "Named:id=33:Normal"
            _StyleDefs(55)  =   ":id=33,.parent=0"
            _StyleDefs(56)  =   "Named:id=34:Heading"
            _StyleDefs(57)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(58)  =   ":id=34,.wraptext=-1"
            _StyleDefs(59)  =   "Named:id=35:Footing"
            _StyleDefs(60)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(61)  =   "Named:id=36:Selected"
            _StyleDefs(62)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(63)  =   "Named:id=37:Caption"
            _StyleDefs(64)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(65)  =   "Named:id=38:HighlightRow"
            _StyleDefs(66)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(67)  =   "Named:id=39:EvenRow"
            _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(69)  =   "Named:id=40:OddRow"
            _StyleDefs(70)  =   ":id=40,.parent=33"
            _StyleDefs(71)  =   "Named:id=41:RecordSelector"
            _StyleDefs(72)  =   ":id=41,.parent=34"
            _StyleDefs(73)  =   "Named:id=42:FilterBar"
            _StyleDefs(74)  =   ":id=42,.parent=33"
         End
      End
   End
   Begin MSComctlLib.Toolbar tbBotoes 
      Height          =   570
      Left            =   7650
      TabIndex        =   1
      Top             =   7800
      Width           =   2880
      _ExtentX        =   5080
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
      Left            =   60
      Top             =   7800
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
            Picture         =   "frmCargos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargos.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargos.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargos.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargos.frx":1668
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargos.frx":1C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargos.frx":219C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargos.frx":2736
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargos.frx":2CD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargos.frx":326A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgcima 
      Height          =   135
      Left            =   180
      Picture         =   "frmCargos.frx":3804
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgbaixo 
      Height          =   135
      Left            =   0
      Picture         =   "frmCargos.frx":398A
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
End
Attribute VB_Name = "frmPermissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mobjRsUsuarios As Recordset
Dim mobjrsPermissao As Recordset

Dim mstrFlag As String
Private Sub Form_Load()
    
    CriarEPreencherRecordsets
    ssgPermissoes.Columns(0).Locked = True
End Sub

Private Sub CriarEPreencherRecordsets(Optional blnFiltrando As Boolean)
      Dim objRsUsuarios As Recordset
      Dim objRsPermissao As Recordset
10    On Error GoTo Erro
20        If blnFiltrando = False Then
30            Set mobjRsUsuarios = Nothing
40            Set mobjRsUsuarios = New Recordset
              
50            With mobjRsUsuarios
60                .Fields.Append "ID_IN", adInteger
70                .Fields.Append "Nome_VC", adVarChar, 1024
80                .Fields.Append "Login_VC", adVarChar, 1024
90                .Fields.Append "Descricao_VC", adVarChar, 1024 'CARGO
100               .CursorLocation = adUseClient
110               .Open , Nothing, adOpenDynamic, adLockOptimistic
120           End With
              
130           Call modManutencao_SelecionarUsuario(objRsUsuarios)
              
140           If Not objRsUsuarios Is Nothing Then
150               If Not objRsUsuarios.BOF And Not objRsUsuarios.EOF Then
160                   If objRsUsuarios.RecordCount > 0 Then
170                       objRsUsuarios.MoveFirst
                          
180                       Do While Not objRsUsuarios.EOF
190                           mobjRsUsuarios.AddNew
                                  
200                           mobjRsUsuarios!ID_IN = NZ(objRsUsuarios!ID_IN)
210                           mobjRsUsuarios!Nome_VC = NS(objRsUsuarios!Nome_VC)
220                           mobjRsUsuarios!Login_VC = NS(objRsUsuarios!Login_VC)
230                           mobjRsUsuarios!Descricao_VC = NS(objRsUsuarios!Descricao_VC)
                              
240                           objRsUsuarios.MoveNext
250                       Loop
260                       ssgUsuarios.DataSource = mobjRsUsuarios
270                   End If
280               End If
290           End If
          
300           mobjRsUsuarios.MoveFirst
310       End If
          
320       Set mobjrsPermissao = Nothing
330       Set mobjrsPermissao = New Recordset
          
340       With mobjrsPermissao
350           .Fields.Append "Permissao_IN", adInteger
360           .Fields.Append "Descricao_VC", adVarChar, 1024
370           .Fields.Append "Status_BT", adBoolean
380           .Fields.Append "check", adBoolean
390           .CursorLocation = adUseClient
400           .Open , Nothing, adOpenDynamic, adLockOptimistic
410       End With
          
420       Call modManutencao_SelecionarPermissaoPorUsuario(mobjRsUsuarios!ID_IN, objRsPermissao)
          
430       If Not objRsPermissao Is Nothing Then
440           If Not objRsPermissao.BOF And Not objRsPermissao.EOF Then
450               If objRsPermissao.RecordCount > 0 Then
460                   objRsPermissao.MoveFirst
                      
470                   Do While Not objRsPermissao.EOF
480                       mobjrsPermissao.AddNew
490                       mobjrsPermissao!Permissao_IN = NZ(objRsPermissao!Permissao_IN)
500                       mobjrsPermissao!Descricao_VC = NS(objRsPermissao!Descricao_VC)
510                       mobjrsPermissao!Status_BT = NB(objRsPermissao!Status_BT)
520                       mobjrsPermissao!check = NB(objRsPermissao!Status_BT)
530                       objRsPermissao.MoveNext
540                   Loop
                      
550                   ssgPermissoes.DataSource = mobjrsPermissao
560               End If
570           End If
580       End If
        On Error Resume Next
590       mobjrsPermissao.MoveFirst
        On Error GoTo Erro
600   Exit Sub
Erro:
610      Call MsgBox("Erro no módulo: " & "frmCargos" & vbCrLf & "CriarEPreencherRecordsets" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")



End Sub

Private Sub ssgUsuarios_Click()
On Error Resume Next
    ssgUsuarios.SelBookmarks.Clear
    ssgUsuarios.SelBookmarks.Add ssgUsuarios.Bookmark
On Error GoTo 0
    
End Sub

Private Sub ssgUsuarios_HeadClick(ByVal ColIndex As Integer)
    OrdenarColunaTrueDB ssgUsuarios, ColIndex, imgcima, imgbaixo
End Sub

Private Sub ssgUsuarios_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ssgUsuarios_Click
    CriarEPreencherRecordsets True
End Sub


Private Sub ssgPermissoes_Click()
On Error Resume Next
    ssgPermissoes.SelBookmarks.Clear
    ssgPermissoes.SelBookmarks.Add ssgPermissoes.Bookmark
On Error GoTo 0

End Sub

'Private Sub ssgPermissoes_HeadClick(ByVal ColIndex As Integer)
'    OrdenarColunaTrueDB ssgPermissoes, ColIndex, imgcima, imgbaixo
'End Sub

Private Sub ssgPermissoes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ssgPermissoes_Click
End Sub

Private Sub tbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Not (Button.Enabled) Then Exit Sub
    Select Case Button.Key

        Case "cmdAlterar":
            mstrFlag = "A"
            Call HabilitarCampos(True)
            Call HabilitarTBBotoes(False, True, True)

        Case "cmdGravar"
            mstrFlag = ""
            GravarAlteracoes
            Call HabilitarCampos(False)
            Call HabilitarTBBotoes(True, False, True)

        Case "cmdSair"
            Unload Me
  
    End Select
End Sub

Private Sub HabilitarCampos(blnHabilitar As Boolean)
On Error GoTo Erro
      
    ssgPermissoes.Columns(0).Locked = Not blnHabilitar
    ssgUsuarios.Enabled = Not blnHabilitar

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmPermissao" & vbCrLf & "HabilitarCampos" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Private Sub HabilitarTBBotoes(blnAlterar As Boolean, blnGravar As Boolean, blnsair As Boolean)
On Error GoTo Erro
      
    tbBotoes.Buttons("cmdAlterar").Enabled = blnAlterar
    tbBotoes.Buttons("cmdGravar").Enabled = blnGravar
    tbBotoes.Buttons("cmdSair").Enabled = blnsair
    
Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmPermissao" & vbCrLf & "HabilitarTBBotoes" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")


End Sub

Private Sub GravarAlteracoes()
On Error GoTo Erro
    Dim objRsPermissaoClone As Recordset
      
    Set objRsPermissaoClone = mobjrsPermissao.Clone
      
    objRsPermissaoClone.MoveFirst
    
    gSMConexao.BeginTransaction
    
    Do While Not objRsPermissaoClone.EOF
    
        Call modManutencao_AdicionarAlterarPermissaoPorUsuario(mobjRsUsuarios!ID_IN, objRsPermissaoClone!Permissao_IN, objRsPermissaoClone!check)
            
        objRsPermissaoClone.MoveNext
    Loop
    
    gSMConexao.CommitTransaction
    MsgBox "Alterações gravadas com sucesso!", vbOKOnly + vbInformation, "Sucesso!"
    
Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmPermissao" & vbCrLf & "GravarAlteracoes" & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
   MsgBox "Alterações não foram gravadas!", vbOKOnly + vbCritical, "Atenção!"
   gSMConexao.RollbackTransaction

End Sub
