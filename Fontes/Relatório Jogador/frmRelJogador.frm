VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form frmRelJogador 
   Caption         =   "ProFut - Relatório de Jogador"
   ClientHeight    =   7755
   ClientLeft      =   4050
   ClientTop       =   1860
   ClientWidth     =   14715
   Icon            =   "frmRelJogador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   14715
   Begin VB.Frame fraLegenda 
      Caption         =   "Legenda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   60
      TabIndex        =   35
      Top             =   7140
      Width           =   2835
      Begin VB.Label Label5 
         Caption         =   "- Feminino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1740
         TabIndex        =   37
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "- Masculino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   390
         TabIndex        =   36
         Top             =   240
         Width           =   1035
      End
      Begin VB.Shape shpRosa 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C00000&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF00FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   1470
         Top             =   240
         Width           =   225
      End
      Begin VB.Shape shpMasculino 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C00000&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C00000&
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   120
         Top             =   240
         Width           =   225
      End
   End
   Begin VB.CommandButton cmdExportar 
      Height          =   345
      Left            =   3030
      Picture         =   "frmRelJogador.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Exportar para Excell"
      Top             =   7290
      Width           =   375
   End
   Begin VB.Frame fraPrincipal 
      Height          =   7185
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   14655
      Begin VB.Frame fraResultado 
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4635
         Left            =   60
         TabIndex        =   3
         Top             =   2490
         Width           =   14535
         Begin TrueOleDBGrid80.TDBGrid ssgResultado 
            Height          =   4335
            Left            =   60
            TabIndex        =   27
            Top             =   240
            Width           =   14400
            _ExtentX        =   25400
            _ExtentY        =   7646
            _LayoutType     =   4
            _RowHeight      =   15
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   16
            Columns(0)._MaxComboItems=   5
            Columns(0).ValueItems(0)._DefaultItem=   0
            Columns(0).ValueItems(0).Value=   "1"
            Columns(0).ValueItems(0).Value.vt=   8
            Columns(0).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(0).ValueItems(0).DisplayValue(0)=   "bHQAAGYCAABCTWYCAAAAAAAANgAAACgAAAANAAAADgAAAAEAGAAAAAAAMAIAAAAAAAAAAAAAAAAA"
            Columns(0).ValueItems(0).DisplayValue(1)=   "AAAAAAD///////////////////////////////////////////////////8A////zBMJzBMJzBMJ"
            Columns(0).ValueItems(0).DisplayValue(2)=   "zBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJ////AP///8wTCcwTCcwTCcwTCcwTCcwTCcwTCcwTCcwT"
            Columns(0).ValueItems(0).DisplayValue(3)=   "CcwTCcwTCf///wD////MEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwn///8A////zBMJ"
            Columns(0).ValueItems(0).DisplayValue(4)=   "zBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJ////AP///8wTCcwTCcwTCcwTCcwTCcwTCcwT"
            Columns(0).ValueItems(0).DisplayValue(5)=   "CcwTCcwTCcwTCcwTCf///wD////MEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwn///8A"
            Columns(0).ValueItems(0).DisplayValue(6)=   "////zBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJ////AP///8wTCcwTCcwTCcwTCcsT"
            Columns(0).ValueItems(0).DisplayValue(7)=   "CcwTCcwTCcwTCcwTCcwTCcwTCf///wD////MEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnM"
            Columns(0).ValueItems(0).DisplayValue(8)=   "Ewn///8A////zBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJ////AP///8wTCcwTCcwT"
            Columns(0).ValueItems(0).DisplayValue(9)=   "CcwTCcwTCcwTCcwTCcwTCcwTCcwTCcwTCf///wD////MEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnM"
            Columns(0).ValueItems(0).DisplayValue(10)=   "EwnMEwnMEwn///8A////////////////////////////////////////////////////AA=="
            Columns(0).ValueItems(0).DisplayValue.vt=   9
            Columns(0).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(0).ValueItems(1)._DefaultItem=   0
            Columns(0).ValueItems(1).Value=   "-1"
            Columns(0).ValueItems(1).Value.vt=   8
            Columns(0).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(0).ValueItems(1).DisplayValue(0)=   "bHQAAGYCAABCTWYCAAAAAAAANgAAACgAAAANAAAADgAAAAEAGAAAAAAAMAIAAAAAAAAAAAAAAAAA"
            Columns(0).ValueItems(1).DisplayValue(1)=   "AAAAAAD///////////////////////////////////////////////////8A////zBMJzBMJzBMJ"
            Columns(0).ValueItems(1).DisplayValue(2)=   "zBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJ////AP///8wTCcwTCcwTCcwTCcwTCcwTCcwTCcwTCcwT"
            Columns(0).ValueItems(1).DisplayValue(3)=   "CcwTCcwTCf///wD////MEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwn///8A////zBMJ"
            Columns(0).ValueItems(1).DisplayValue(4)=   "zBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJ////AP///8wTCcwTCcwTCcwTCcwTCcwTCcwT"
            Columns(0).ValueItems(1).DisplayValue(5)=   "CcwTCcwTCcwTCcwTCf///wD////MEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwn///8A"
            Columns(0).ValueItems(1).DisplayValue(6)=   "////zBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJ////AP///8wTCcwTCcwTCcwTCcsT"
            Columns(0).ValueItems(1).DisplayValue(7)=   "CcwTCcwTCcwTCcwTCcwTCcwTCf///wD////MEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnM"
            Columns(0).ValueItems(1).DisplayValue(8)=   "Ewn///8A////zBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJzBMJ////AP///8wTCcwTCcwT"
            Columns(0).ValueItems(1).DisplayValue(9)=   "CcwTCcwTCcwTCcwTCcwTCcwTCcwTCcwTCf///wD////MEwnMEwnMEwnMEwnMEwnMEwnMEwnMEwnM"
            Columns(0).ValueItems(1).DisplayValue(10)=   "EwnMEwnMEwn///8A////////////////////////////////////////////////////AA=="
            Columns(0).ValueItems(1).DisplayValue.vt=   9
            Columns(0).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(0).ValueItems(2)._DefaultItem=   0
            Columns(0).ValueItems(2).Value=   "2"
            Columns(0).ValueItems(2).Value.vt=   8
            Columns(0).ValueItems(2).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(0).ValueItems(2).DisplayValue(0)=   "bHQAAGYCAABCTWYCAAAAAAAANgAAACgAAAANAAAADgAAAAEAGAAAAAAAMAIAAAAAAAAAAAAAAAAA"
            Columns(0).ValueItems(2).DisplayValue(1)=   "AAAAAAD///////////////////////////////////////////////////8A////ya7/ya7/ya7/"
            Columns(0).ValueItems(2).DisplayValue(2)=   "ya7/ya7/ya7/ya7/ya7/ya7/ya7/ya7/////AP///8mu/8mu/8mu/8mu/8mu/8mu/8mu/8mu/8mu"
            Columns(0).ValueItems(2).DisplayValue(3)=   "/8mu/8mu/////wD////Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/Jrv////8A////ya7/"
            Columns(0).ValueItems(2).DisplayValue(4)=   "ya7/ya7/ya7/ya7/ya7/ya7/ya7/ya7/ya7/ya7/////AP///8mu/8mu/8mu/8mu/8mu/8mu/8mu"
            Columns(0).ValueItems(2).DisplayValue(5)=   "/8mu/8mu/8mu/8mu/////wD////Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/Jrv////8A"
            Columns(0).ValueItems(2).DisplayValue(6)=   "////ya7/ya7/ya7/ya7/ya7/ya7/ya7/ya7/ya7/ya7/ya7/////AP///8mu/8mu/8mu/8mu/8mu"
            Columns(0).ValueItems(2).DisplayValue(7)=   "/8mu/8mu/8mu/8mu/8mu/8mu/////wD////Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/J"
            Columns(0).ValueItems(2).DisplayValue(8)=   "rv////8A////ya7/ya7/ya7/ya7/ya7/ya7/ya7/ya7/ya7/ya7/ya7/////AP///8mu/8mu/8mu"
            Columns(0).ValueItems(2).DisplayValue(9)=   "/8mu/8mu/8mu/8mu/8mu/8mu/8mu/8mu/////wD////Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/Jrv/J"
            Columns(0).ValueItems(2).DisplayValue(10)=   "rv/Jrv/Jrv////8A////////////////////////////////////////////////////AA=="
            Columns(0).ValueItems(2).DisplayValue.vt=   9
            Columns(0).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
            Columns(0).ValueItems.Count=   3
            Columns(0).DataField=   "sexo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Cod."
            Columns(1).DataField=   "Codigo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nome"
            Columns(2).DataField=   "Nome"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Apelido"
            Columns(3).DataField=   "Apelido"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Equipe"
            Columns(4).DataField=   "Equipe"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Cartegoria"
            Columns(5).DataField=   "Cartegoria"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Data Nascimento"
            Columns(6).DataField=   "DataNascimento"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=529"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=450"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=714"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=635"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=8652"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=8573"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(17)=   "Column(3).Width=5424"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=5345"
            Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(23)=   "Column(4).Width=4630"
            Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=4551"
            Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(29)=   "Column(5).Width=3413"
            Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=3334"
            Splits(0)._ColumnProps(32)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(35)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(38)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._ColStyle=512"
            Splits(0)._ColumnProps(40)=   "Column(6).Order=7"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
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
            _StyleDefs(31)  =   "Splits(0).Style:id=13,.parent=1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(32)  =   ":id=13,.strikethrough=0,.charset=0"
            _StyleDefs(33)  =   ":id=13,.fontname=Arial"
            _StyleDefs(34)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HC0C0C0&"
            _StyleDefs(35)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H0&"
            _StyleDefs(36)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(37)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(38)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&H800000&,.fgcolor=&HFFFFFF&"
            _StyleDefs(39)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(40)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(41)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(42)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(43)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(44)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(45)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
            _StyleDefs(46)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(47)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(48)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(49)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
            _StyleDefs(50)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
            _StyleDefs(51)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
            _StyleDefs(52)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
            _StyleDefs(53)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=0"
            _StyleDefs(54)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14,.alignment=2"
            _StyleDefs(55)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(56)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(57)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(58)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14,.alignment=2"
            _StyleDefs(59)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(60)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(61)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(62)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14,.alignment=2"
            _StyleDefs(63)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(64)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(65)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(66)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14,.alignment=2"
            _StyleDefs(67)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(68)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(69)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=0"
            _StyleDefs(70)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14,.alignment=2"
            _StyleDefs(71)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(72)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(73)  =   "Named:id=33:Normal"
            _StyleDefs(74)  =   ":id=33,.parent=0"
            _StyleDefs(75)  =   "Named:id=34:Heading"
            _StyleDefs(76)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(77)  =   ":id=34,.wraptext=-1"
            _StyleDefs(78)  =   "Named:id=35:Footing"
            _StyleDefs(79)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(80)  =   "Named:id=36:Selected"
            _StyleDefs(81)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(82)  =   "Named:id=37:Caption"
            _StyleDefs(83)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(84)  =   "Named:id=38:HighlightRow"
            _StyleDefs(85)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(86)  =   "Named:id=39:EvenRow"
            _StyleDefs(87)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(88)  =   "Named:id=40:OddRow"
            _StyleDefs(89)  =   ":id=40,.parent=33"
            _StyleDefs(90)  =   "Named:id=41:RecordSelector"
            _StyleDefs(91)  =   ":id=41,.parent=34"
            _StyleDefs(92)  =   "Named:id=42:FilterBar"
            _StyleDefs(93)  =   ":id=42,.parent=33"
         End
         Begin VB.Label lblinfo 
            Caption         =   "INFORMAÇÃO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   4740
            TabIndex        =   33
            Top             =   2130
            Width           =   6015
         End
      End
      Begin VB.Frame fraFiltros 
         Caption         =   "Filtros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   60
         TabIndex        =   2
         Top             =   120
         Width           =   14535
         Begin VB.Frame fraEquipes 
            Height          =   2055
            Left            =   11220
            TabIndex        =   25
            Top             =   120
            Width           =   3225
            Begin VB.CheckBox chkEquipe 
               Caption         =   "Equipe"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   -30
               Width           =   975
            End
            Begin TrueOleDBGrid80.TDBGrid ssgEquipes 
               Height          =   1725
               Left            =   90
               TabIndex        =   28
               Top             =   240
               Width           =   3030
               _ExtentX        =   5345
               _ExtentY        =   3043
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
               Columns(1).Caption=   "Equipe"
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
               Splits(0)._ColumnProps(1)=   "Column(0).Width=635"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=556"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
               Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=1"
               Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(8)=   "Column(1).Width=4630"
               Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=4551"
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
         Begin VB.Frame Frame1 
            Height          =   1005
            Left            =   3390
            TabIndex        =   18
            Top             =   1170
            Width           =   7785
            Begin VB.TextBox txtBairroEnderecoAtleta 
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
               Left            =   4470
               MaxLength       =   128
               TabIndex        =   20
               Top             =   480
               Width           =   3255
            End
            Begin VB.CheckBox chkEndereco 
               Caption         =   "Endereço"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   150
               TabIndex        =   19
               Top             =   -30
               Width           =   1185
            End
            Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscUfEnderecoAtleta 
               Height          =   390
               Left            =   60
               TabIndex        =   21
               Top             =   480
               Width           =   795
               DataFieldList   =   "Column 0"
               AllowInput      =   0   'False
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
               _ExtentX        =   1402
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
            Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscCidadeEnderecoAtleta 
               Height          =   390
               Left            =   900
               TabIndex        =   34
               Top             =   480
               Width           =   3555
               DataFieldList   =   "Column 0"
               AllowInput      =   0   'False
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
               _ExtentX        =   6271
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
            Begin VB.Label Label12 
               Caption         =   "Bairro"
               Height          =   285
               Left            =   4470
               TabIndex        =   24
               Top             =   240
               Width           =   2235
            End
            Begin VB.Label Label13 
               Caption         =   "Cidade"
               Height          =   285
               Left            =   900
               TabIndex        =   23
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label14 
               Caption         =   "UF"
               Height          =   285
               Left            =   60
               TabIndex        =   22
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame fraCartegoria 
            Height          =   885
            Left            =   3390
            TabIndex        =   11
            Top             =   270
            Width           =   3525
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
               Left            =   2790
               TabIndex        =   32
               Top             =   390
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
               Left            =   2250
               TabIndex        =   31
               Top             =   390
               Value           =   -1  'True
               Width           =   555
            End
            Begin VB.CheckBox chkCartegoria 
               Caption         =   "Cartegoria"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   150
               TabIndex        =   16
               Top             =   -30
               Width           =   1245
            End
            Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscCartegoria 
               Height          =   390
               Left            =   90
               TabIndex        =   26
               Top             =   300
               Width           =   1995
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
               _ExtentX        =   3519
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
         End
         Begin VB.Frame Frame2 
            Height          =   885
            Left            =   6930
            TabIndex        =   9
            Top             =   270
            Width           =   4245
            Begin VB.CheckBox chkDataNascimento 
               Caption         =   "Data de Nascimento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   150
               TabIndex        =   17
               Top             =   -30
               Width           =   2055
            End
            Begin SSCalendarWidgets_A.SSDateCombo dtcDataNascimentoInicial 
               Height          =   405
               Left            =   360
               TabIndex        =   10
               Top             =   270
               Width           =   1665
               _Version        =   65543
               _ExtentX        =   2937
               _ExtentY        =   714
               _StockProps     =   93
               BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "DD/MM/YYYY"
               BevelType       =   0
            End
            Begin SSCalendarWidgets_A.SSDateCombo dtcDataNascimentoFinal 
               Height          =   405
               Left            =   2490
               TabIndex        =   12
               Top             =   270
               Width           =   1665
               _Version        =   65543
               _ExtentX        =   2937
               _ExtentY        =   714
               _StockProps     =   93
               Format          =   "DD/MM/YYYY"
               BevelType       =   0
            End
            Begin VB.Label Label3 
               Caption         =   "Até :"
               Height          =   285
               Left            =   2100
               TabIndex        =   14
               Top             =   360
               Width           =   345
            End
            Begin VB.Label Label1 
               Caption         =   "De :"
               Height          =   285
               Left            =   30
               TabIndex        =   13
               Top             =   360
               Width           =   315
            End
         End
         Begin VB.Frame fraInfoJogador 
            Height          =   1905
            Left            =   60
            TabIndex        =   4
            Top             =   270
            Width           =   3315
            Begin VB.CheckBox chkJogador 
               Caption         =   "Jogador"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   -30
               Width           =   1035
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
               Left            =   60
               MaxLength       =   128
               TabIndex        =   6
               Top             =   1260
               Width           =   3105
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
               Left            =   60
               MaxLength       =   20
               TabIndex        =   5
               Top             =   450
               Width           =   3105
            End
            Begin VB.Label Apelido 
               Caption         =   "Apelido do Jogador"
               Height          =   285
               Left            =   60
               TabIndex        =   8
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label2 
               Caption         =   "Nome Completo"
               Height          =   285
               Left            =   60
               TabIndex        =   7
               Top             =   1050
               Width           =   975
            End
         End
      End
   End
   Begin MSComctlLib.Toolbar tbBotoes 
      Height          =   570
      Left            =   13020
      TabIndex        =   1
      Top             =   7170
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   1005
      ButtonWidth     =   1376
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F7-Gerar"
            Key             =   "cmdGerar"
            Object.ToolTipText     =   "Gravar Alterações"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F10-Sair"
            Key             =   "cmdSair"
            Object.ToolTipText     =   "Sair da tela"
            ImageIndex      =   9
         EndProperty
      EndProperty
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
            Picture         =   "frmRelJogador.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelJogador.frx":0BCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelJogador.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelJogador.frx":1702
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelJogador.frx":1C9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelJogador.frx":2236
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelJogador.frx":27D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelJogador.frx":2D6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelJogador.frx":3304
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelJogador.frx":389E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgcima 
      Height          =   135
      Left            =   180
      Picture         =   "frmRelJogador.frx":3E38
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgbaixo 
      Height          =   135
      Left            =   0
      Picture         =   "frmRelJogador.frx":3FBE
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Menu mnuVisualizar 
      Caption         =   "mnuVisualizar"
      Visible         =   0   'False
      Begin VB.Menu mnuVisualizarJogasor 
         Caption         =   "Visualizar Jogador"
      End
   End
End
Attribute VB_Name = "frmRelJogador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mobjRsEquipes As Recordset
Dim mobjRsResultado As Recordset
Dim mblnCarregadoPorcurar As Boolean

Dim mstrEquipes As String

Public Property Let CarregadoViaProcurar(blnProcurar As Boolean)
    mblnCarregadoPorcurar = blnProcurar
End Property

Public Property Get IDJogador() As Integer
    IDJogador = IIf(mobjRsResultado!Codigo = 0, 0, mobjRsResultado!Codigo)
End Property

Private Sub chkCartegoria_Click()
    If chkCartegoria.Value = vbChecked Then
        sscCartegoria.Enabled = True
        optMasculino.Enabled = True
        optFeminino.Enabled = True
    Else
        optMasculino.Enabled = False
        optFeminino.Enabled = False
        sscCartegoria.Enabled = False
    End If
End Sub

Private Sub chkDataNascimento_Click()
    If chkDataNascimento.Value = vbChecked Then
        dtcDataNascimentoInicial.Enabled = True
        dtcDataNascimentoFinal.Enabled = True
    Else
        dtcDataNascimentoInicial.Enabled = False
        dtcDataNascimentoFinal.Enabled = False
    End If
End Sub

Private Sub chkEndereco_Click()
    If chkEndereco.Value = vbChecked Then
        sscUfEnderecoAtleta.Enabled = True
        sscCidadeEnderecoAtleta.Enabled = True
        txtBairroEnderecoAtleta.Enabled = True
    Else
        sscUfEnderecoAtleta.Enabled = False
        sscCidadeEnderecoAtleta.Enabled = False
        txtBairroEnderecoAtleta.Enabled = False
    End If
End Sub

Private Sub chkEquipe_Click()
    If chkEquipe.Value = vbChecked Then
        ssgEquipes.Enabled = True
    Else
        ssgEquipes.Enabled = False
    End If
End Sub

Private Sub chkJogador_Click()
    If chkJogador.Value = vbChecked Then
        txtApelido.Enabled = True
        txtNomeJogador.Enabled = True
    Else
        txtApelido.Enabled = False
        txtNomeJogador.Enabled = False
    End If
End Sub



Private Sub cmdExportar_Click()
On Error GoTo Erro
Dim clsExportar As clsExportarExcell
    
    Set clsExportar = New clsExportarExcell
        
    If mobjRsResultado Is Nothing Then Exit Sub
    
    MousePointer = vbHourglass

    
    DoEvents
    
    If Not mobjRsResultado.EOF And Not mobjRsResultado.BOF Then
        
        lblinfo.Visible = True
        lblinfo.ZOrder 0
        ssgResultado.Visible = False
        lblinfo.Caption = "Gerando Tabela Excell..."
        Call clsExportar.ExportarGrid(ssgResultado)
        lblinfo.Visible = False
        ssgResultado.Visible = True
    End If
    
    MousePointer = vbDefault
    
    
    DoEvents
Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmRelJogador" & vbCrLf & "No Procedimento: " & "cmdExportar_Click" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Private Sub Form_Load()
Dim objrs As Recordset
Dim lngUsuario As Long

    lngUsuario = gSMConexao.CodigoUsuario

    lblinfo.Visible = False
        
    chkCartegoria_Click
    chkDataNascimento_Click
    chkEndereco_Click
    chkEquipe_Click
    chkJogador_Click
    
    modBDCombo_SelecionarEstados sscUfEnderecoAtleta
    modBDCombo_SelecionarCartecoriaJogador sscCartegoria
    
    If RetornaAcessoPorUsuarioEPermissao(lngUsuario, 12) = True Then
        modEquipe_SelecionarEquipePorCodigo 0, objrs
    Else
        modEquipe_SelecionarEquipePorCodigo RetornaClubePorUsuario(lngUsuario), objrs
        chkEquipe.Value = vbChecked
        chkEquipe_Click
        chkEquipe.Enabled = False
        ssgEquipes.Enabled = False
    End If
    
    
    CriarEPreencherRecordsetEquipes objrs
    
    ssgEquipes.DataSource = mobjRsEquipes
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF7:  tbBotoes.Buttons("cmdGerar").Value = tbrPressed
        Case vbKeyF10: tbBotoes.Buttons("cmdSair").Value = tbrPressed
   End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    tbBotoes.Buttons("cmdGerar").Value = tbrUnpressed
    tbBotoes.Buttons("cmdSair").Value = tbrUnpressed
  
    Select Case KeyCode
        Case vbKeyF7:  If tbBotoes.Buttons("cmdGerar").Enabled Then Call tbBotoes_ButtonClick(tbBotoes.Buttons("cmdGerar"))
        Case vbKeyF10: If tbBotoes.Buttons("cmdSair").Enabled Then Call tbBotoes_ButtonClick(tbBotoes.Buttons("cmdSair"))
    End Select

End Sub

Private Sub Form_Resize()

    If WindowState = vbMinimized Then Exit Sub
    If Height < 8325 Then Height = 8325
    If Width < 14955 Then Width = 14955

    fraPrincipal.Width = frmRelJogador.Width - 300
    fraPrincipal.Height = frmRelJogador.Height - 1140
    
    cmdExportar.Top = fraPrincipal.Height + 135
    
    fraResultado.Height = fraPrincipal.Height - 2550
    fraResultado.Width = fraPrincipal.Width - 125
    
    ssgResultado.Height = fraResultado.Height - 300
    ssgResultado.Width = fraResultado.Width - 150
    
    fraLegenda.Top = fraPrincipal.Height
    
    
    tbBotoes.Left = Me.Width - 1800
    tbBotoes.Top = fraPrincipal.Top + fraPrincipal.Height + 50

End Sub

Private Sub mnuVisualizarJogasor_Click()
On Error GoTo Erro
      
        If Not mobjRsResultado Is Nothing Then
            If Not mobjRsResultado.EOF And Not mobjRsResultado.BOF Then
                If Not mobjRsResultado.RecordCount = 0 Then
                    If mblnCarregadoPorcurar = True Then
                        Unload Me
                    Else
                        Dim objCadJogador As clsCadJogador
                        Set objCadJogador = New clsCadJogador
                        
                        objCadJogador.Show gSMConexao, , vbModeless, , mobjRsResultado!Codigo, True
                    End If
                End If
            End If
        End If

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmRelJogador" & vbCrLf & "No Procedimento: " & "mnuVisualizarJogasor_Click" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub



Private Sub sscUfEnderecoAtleta_LostFocus()
    Call modBDCombo_SelecionarCidades(sscCidadeEnderecoAtleta, , sscUfEnderecoAtleta.Columns("chcodigo").Value)
End Sub

Private Sub ssgEquipes_Click()
On Error Resume Next
    ssgEquipes.SelBookmarks.Clear
    ssgEquipes.SelBookmarks.Add ssgEquipes.Bookmark
On Error GoTo 0
End Sub

Private Sub ssgEquipes_HeadClick(ByVal ColIndex As Integer)
    OrdenarColunaTrueDB ssgEquipes, ColIndex, imgcima, imgbaixo
End Sub

Private Sub ssgEquipes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ssgEquipes_Click
End Sub


Private Sub ssgResultado_Click()
On Error Resume Next
    ssgResultado.SelBookmarks.Clear
    ssgResultado.SelBookmarks.Add ssgResultado.Bookmark
On Error GoTo 0
End Sub

Private Sub ssgResultado_HeadClick(ByVal ColIndex As Integer)
    OrdenarColunaTrueDB ssgResultado, ColIndex, imgcima, imgbaixo
End Sub

Private Sub ssgResultado_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If Not mobjRsResultado Is Nothing Then
            If Not mobjRsResultado.EOF And Not mobjRsResultado.BOF Then
                If Not mobjRsResultado.RecordCount = 0 Then
                    PopupMenu mnuVisualizar
                End If
            End If
        End If
    End If
End Sub

Private Sub ssgResultado_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ssgResultado_Click
End Sub

Private Sub tbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Not (Button.Enabled) Then Exit Sub
    Select Case Button.Key
    
        Case "cmdGerar"
            GerarRelatorio
        
        Case "cmdSair"
            Unload Me
        
    End Select
End Sub

Private Sub CriarEPreencherRecordsetResultadoResultado(ByRef objRsResultado As Recordset)
On Error GoTo Erro
      
      
    Set mobjRsResultado = Nothing
    Set mobjRsResultado = New Recordset

    
    With mobjRsResultado
        .Fields.Append "Codigo", adInteger
        .Fields.Append "Sexo", adInteger
        .Fields.Append "Nome", adVarChar, 1024
        .Fields.Append "Apelido", adVarChar, 1024
        .Fields.Append "Equipe", adVarChar, 1024
        .Fields.Append "Cartegoria", adVarChar, 1024
        .Fields.Append "DataNascimento", adDate
        .CursorLocation = adUseClient
        .Open , Nothing, adOpenDynamic, adLockOptimistic
    End With
    
    If Not objRsResultado Is Nothing Then
        If Not objRsResultado.BOF Or Not objRsResultado.EOF Then
            If objRsResultado.RecordCount > 0 Then
                objRsResultado.MoveFirst
                Do While Not objRsResultado.EOF
                    mobjRsResultado.AddNew
                    
                    mobjRsResultado!Codigo = NZ(objRsResultado!Codigo)
                    mobjRsResultado!Sexo = NZ(objRsResultado!Sexo)
                    mobjRsResultado!Nome = NS(objRsResultado!Nome)
                    mobjRsResultado!Apelido = NS(objRsResultado!Apelido)
                    mobjRsResultado!Equipe = NS(objRsResultado!Equipe)
                    mobjRsResultado!Cartegoria = NS(objRsResultado!Cartegoria)
                    mobjRsResultado!DataNascimento = NS(objRsResultado!DataNascimento)
                    
                    mobjRsResultado.Update
                    objRsResultado.MoveNext
                Loop
            End If
        End If
    End If
    
    ssgResultado.DataSource = mobjRsResultado
    
    If Not mobjRsResultado Is Nothing Then
        If Not mobjRsResultado.BOF And Not mobjRsResultado.EOF Then
            If mobjRsResultado.RecordCount = 0 Then
                           
                lblinfo.Visible = True
                ssgResultado.Visible = False
                lblinfo.ZOrder 0
                lblinfo.Caption = "Sem Resultados"
            Else
                ssgResultado.Visible = True
                lblinfo.Visible = False
            End If
        Else
            lblinfo.Visible = True
            lblinfo.ZOrder 0
            ssgResultado.Visible = False
            lblinfo.Caption = "Sem Resultados"
        End If
    Else
        lblinfo.Visible = True
        lblinfo.ZOrder 0
        ssgResultado.Visible = False
        lblinfo.Caption = "Sem Resultados"
    End If
    
Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmRelJogador" & vbCrLf & "No Procedimento: " & "CriarEPreencherRecordsetResultado" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Private Sub CriarEPreencherRecordsetEquipes(ByRef objRsEquipes As Recordset)
On Error GoTo Erro
      
      'modEquipe_SelecionarEquipePorCodigo 0, mobjRsEquipes
      
      Set mobjRsEquipes = Nothing
      Set mobjRsEquipes = New Recordset
      
      With mobjRsEquipes
        
        .Fields.Append "marcado_BT", adBoolean
        .Fields.Append "ID_IN", adInteger
        .Fields.Append "nome_VC", adVarChar, 1024
        
        .CursorLocation = adUseClient
        .Open , Nothing, adOpenDynamic, adLockOptimistic
      End With
      
      If Not objRsEquipes Is Nothing Then
        If Not objRsEquipes.BOF And Not objRsEquipes.EOF Then
            objRsEquipes.MoveFirst
            If objRsEquipes.RecordCount > 0 Then
                Do While Not objRsEquipes.EOF
                
                    mobjRsEquipes.AddNew
                    
                    mobjRsEquipes!marcado_bt = IIf(RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 12), False, True)
                    mobjRsEquipes!ID_IN = NZ(objRsEquipes!ID_IN)
                    mobjRsEquipes!Nome_VC = NS(objRsEquipes!Nome_VC)
                                    
                    objRsEquipes.MoveNext
                Loop
            End If
        End If
      End If
      
      ssgEquipes.DataSource = mobjRsEquipes

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmRelJogador" & vbCrLf & "No Procedimento: " & "CriarEPreencherRecordsetEquipes" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Private Sub GerarRelatorio()
On Error GoTo Erro
      
Dim udtJogador As TypJogador
Dim objRsJogador As Recordset

Dim objRsEquipes As Recordset

    If chkJogador.Value = vbChecked Then
        udtJogador.strApelido = Trim(txtApelido.Text)
        udtJogador.strNomeAtleta = Trim(txtNomeJogador.Text)
    End If
    
    If chkCartegoria.Value = vbChecked Then
        udtJogador.lngCartegoria = IIf(sscCartegoria.Columns("chcodigo").Value = 1 And sscCartegoria.Text = "", 0, sscCartegoria.Columns("chcodigo").Value)
        udtJogador.lngSexo = IIf(optMasculino.Value = True, 1, 2)
    End If
    
    If chkDataNascimento.Value = vbChecked Then
        udtJogador.datDataNascimentoDE = IIf(dtcDataNascimentoInicial.DateValue < 10, Empty, dtcDataNascimentoInicial.DateValue)
        udtJogador.datDataNascimentoATE = IIf(dtcDataNascimentoFinal.DateValue < 10, Empty, dtcDataNascimentoFinal.DateValue)
    End If
    
    If chkEndereco.Value = vbChecked Then
        udtJogador.lngEstado = IIf(sscUfEnderecoAtleta.Columns("chcodigo").Value = 1 And sscUfEnderecoAtleta.Text = "", 0, sscUfEnderecoAtleta.Columns("chcodigo").Value)
        udtJogador.strCidade = Trim(sscCidadeEnderecoAtleta.Text)
        udtJogador.strBairro = Trim(txtBairroEnderecoAtleta.Text)
    End If
    
    
    If chkEquipe.Value = vbChecked Then
        
        If Not mobjRsEquipes Is Nothing Then
            If Not mobjRsEquipes.BOF And Not mobjRsEquipes.EOF Then
                mstrEquipes = ""
                Set objRsEquipes = mobjRsEquipes.Clone
                objRsEquipes.MoveFirst
                Do While Not objRsEquipes.EOF
                    If objRsEquipes!marcado_bt = True Then
                        mstrEquipes = mstrEquipes & objRsEquipes!ID_IN & ","
                    End If
                    objRsEquipes.MoveNext
                Loop
                udtJogador.strEquipes = mstrEquipes
            End If
        End If
    End If
'
    modJogador_SelecionarDadosParaRelatorioDeJogador udtJogador, objRsJogador
    
    CriarEPreencherRecordsetResultadoResultado objRsJogador

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmRelJogador" & vbCrLf & "No Procedimento: " & "GerarRelatorio" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub
