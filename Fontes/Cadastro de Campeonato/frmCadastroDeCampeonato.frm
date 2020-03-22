VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmCadastroDeCampeonato 
   Caption         =   "ProFut - Cadastro de Campeonato"
   ClientHeight    =   8475
   ClientLeft      =   10770
   ClientTop       =   1650
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   6660
   Begin VB.Frame fraPrincipal 
      Height          =   7905
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   6585
      Begin VB.Frame fraEquipesCampeonato 
         Caption         =   "Equipes do Campeonato / Classificação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   60
         TabIndex        =   9
         Top             =   4110
         Width           =   6465
         Begin VB.Frame fraJogos 
            Caption         =   "Jogos do Campeonato"
            Height          =   615
            Left            =   60
            TabIndex        =   32
            Top             =   3060
            Width           =   6345
            Begin VB.CommandButton cmdVisualizarJogos 
               Appearance      =   0  'Flat
               Caption         =   "Visualizar Jogos"
               Height          =   315
               Left            =   3210
               TabIndex        =   34
               Top             =   210
               Width           =   1635
            End
            Begin VB.CommandButton cmdNovoJogo 
               Appearance      =   0  'Flat
               Caption         =   "Adicionar Jogo"
               Height          =   315
               Left            =   1500
               TabIndex        =   33
               Top             =   210
               Width           =   1635
            End
         End
         Begin TrueOleDBGrid80.TDBGrid ssgCompeticoesDoClube 
            Height          =   2415
            Left            =   60
            TabIndex        =   10
            Top             =   210
            Width           =   6345
            _ExtentX        =   11192
            _ExtentY        =   4260
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Colocação"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Grupo"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Equipe"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Cartegoria"
            Columns(3).DataField=   ""
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Nº Jogos"
            Columns(4).DataField=   ""
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Vitórias"
            Columns(5).DataField=   ""
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Empates"
            Columns(6).DataField=   ""
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Derrotas"
            Columns(7).DataField=   ""
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "GP"
            Columns(8).DataField=   ""
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "GC"
            Columns(9).DataField=   ""
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "SG"
            Columns(10).DataField=   ""
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   11
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   953
            Splits(0)._SavedRecordSelectors=   -1  'True
            Splits(0).ScrollBars=   3
            Splits(0).DividerColor=   15790320
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=11"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1614"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1535"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=1005"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=926"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=3413"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3334"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=2090"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2011"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(4).Width=1402"
            Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1323"
            Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(21)=   "Column(5).Width=1164"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1085"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=1349"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1270"
            Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(29)=   "Column(7).Width=1244"
            Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=1164"
            Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(33)=   "Column(8).Width=953"
            Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=873"
            Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(37)=   "Column(9).Width=847"
            Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=767"
            Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(41)=   "Column(10).Width=1032"
            Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=953"
            Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   2
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   15790320
            RowDividerColor =   15790320
            RowSubDividerColor=   15790320
            DirectionAfterEnter=   1
            DirectionAfterTab=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
            _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
            _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
            _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
            _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
            _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=78,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=75,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=76,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=77,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=54,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=58,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
            _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
            _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=74,.parent=13"
            _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
            _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
            _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
            _StyleDefs(74)  =   "Named:id=33:Normal"
            _StyleDefs(75)  =   ":id=33,.parent=0"
            _StyleDefs(76)  =   "Named:id=34:Heading"
            _StyleDefs(77)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(78)  =   ":id=34,.wraptext=-1"
            _StyleDefs(79)  =   "Named:id=35:Footing"
            _StyleDefs(80)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(81)  =   "Named:id=36:Selected"
            _StyleDefs(82)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(83)  =   "Named:id=37:Caption"
            _StyleDefs(84)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(85)  =   "Named:id=38:HighlightRow"
            _StyleDefs(86)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(87)  =   "Named:id=39:EvenRow"
            _StyleDefs(88)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(89)  =   "Named:id=40:OddRow"
            _StyleDefs(90)  =   ":id=40,.parent=33"
            _StyleDefs(91)  =   "Named:id=41:RecordSelector"
            _StyleDefs(92)  =   ":id=41,.parent=34"
            _StyleDefs(93)  =   "Named:id=42:FilterBar"
            _StyleDefs(94)  =   ":id=42,.parent=33"
         End
         Begin Threed.SSCommand cmdRemover 
            Height          =   330
            Left            =   1920
            TabIndex        =   26
            Top             =   2700
            Width           =   1725
            _ExtentX        =   3043
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
            Picture         =   "frmCadastroDeCampeonato.frx":0000
            Caption         =   "     Remover Equipe"
            ButtonStyle     =   3
            PictureAlignment=   1
         End
         Begin Threed.SSCommand cmdAdicionar 
            Height          =   330
            Left            =   60
            TabIndex        =   27
            Top             =   2700
            Width           =   1845
            _ExtentX        =   3254
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
            Picture         =   "frmCadastroDeCampeonato.frx":0322
            Caption         =   "     Adicionar Equipe"
            ButtonStyle     =   3
            PictureAlignment=   1
         End
         Begin VB.Label lblQtdEquipes 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5880
            TabIndex        =   29
            Top             =   2760
            Width           =   555
         End
         Begin VB.Label Label7 
            Caption         =   "Quantidade de Equipes :"
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
            Left            =   3720
            TabIndex        =   28
            Top             =   2760
            Width           =   2115
         End
      End
      Begin VB.Frame fraInfoClube 
         Caption         =   "Informações do Campeonato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4005
         Left            =   60
         TabIndex        =   2
         Top             =   120
         Width           =   6465
         Begin VB.CommandButton cmdFinalizarCampeonato 
            Appearance      =   0  'Flat
            Caption         =   "Finalizar Campeonato"
            Height          =   315
            Left            =   4710
            TabIndex        =   31
            Top             =   2700
            Width           =   1635
         End
         Begin VB.CommandButton cmdIniciarCampeonato 
            Appearance      =   0  'Flat
            Caption         =   "Iniciar Campeonato"
            Height          =   315
            Left            =   3000
            TabIndex        =   30
            Top             =   2700
            Width           =   1635
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   1005
            Left            =   120
            TabIndex        =   24
            Top             =   1650
            Width           =   6255
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   5160
            TabIndex        =   20
            Top             =   420
            Width           =   1215
         End
         Begin VB.Frame fraInfoSistema 
            Caption         =   "Informações Sistema"
            Height          =   915
            Left            =   60
            TabIndex        =   15
            Top             =   3000
            Width           =   6345
            Begin VB.TextBox txtStatusCampeonato 
               Appearance      =   0  'Flat
               Height          =   405
               Left            =   60
               TabIndex        =   22
               Top             =   420
               Width           =   2295
            End
            Begin VB.TextBox txtUsuarioAlteracao 
               Appearance      =   0  'Flat
               Height          =   405
               Left            =   2400
               TabIndex        =   16
               Top             =   420
               Width           =   1995
            End
            Begin SSCalendarWidgets_A.SSDateCombo dtcDataUltimaAlteracao 
               Height          =   405
               Left            =   4440
               TabIndex        =   17
               Top             =   420
               Width           =   1815
               _Version        =   65543
               _ExtentX        =   3201
               _ExtentY        =   714
               _StockProps     =   93
               Format          =   "DD MM,YYYY"
               BevelType       =   0
               Mask            =   2
            End
            Begin VB.Label Label2 
               Caption         =   "Status Campeonato"
               Height          =   285
               Left            =   60
               TabIndex        =   23
               Top             =   210
               Width           =   1665
            End
            Begin VB.Label Label26 
               Caption         =   "Usuário Ultima Alteração"
               Height          =   285
               Left            =   2400
               TabIndex        =   19
               Top             =   210
               Width           =   1965
            End
            Begin VB.Label Label27 
               Caption         =   "Data Ultima alteração"
               Height          =   285
               Left            =   4440
               TabIndex        =   18
               Top             =   210
               Width           =   1575
            End
         End
         Begin VB.TextBox txtNomeCampeonato 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   1110
            TabIndex        =   4
            Top             =   420
            Width           =   4005
         End
         Begin VB.TextBox txtCodigoInterno 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   120
            TabIndex        =   3
            Top             =   420
            Width           =   945
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscCartegoria 
            Height          =   390
            Left            =   120
            TabIndex        =   8
            Top             =   1050
            Width           =   2175
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
         Begin SSCalendarWidgets_A.SSDateCombo dtcDataInicio 
            Height          =   405
            Left            =   2340
            TabIndex        =   11
            Top             =   1050
            Width           =   1995
            _Version        =   65543
            _ExtentX        =   3519
            _ExtentY        =   714
            _StockProps     =   93
            Format          =   "DD MM,YYYY"
            BevelType       =   0
            Mask            =   2
         End
         Begin SSCalendarWidgets_A.SSDateCombo dtcDataFinal 
            Height          =   405
            Left            =   4380
            TabIndex        =   13
            Top             =   1050
            Width           =   1995
            _Version        =   65543
            _ExtentX        =   3519
            _ExtentY        =   714
            _StockProps     =   93
            Format          =   "DD MM,YYYY"
            BevelType       =   0
            Mask            =   2
         End
         Begin VB.Label Label6 
            Caption         =   "Descrição/Observação"
            Height          =   285
            Left            =   150
            TabIndex        =   25
            Top             =   1440
            Width           =   1725
         End
         Begin VB.Label Label5 
            Caption         =   "Edição"
            Height          =   285
            Left            =   5190
            TabIndex        =   21
            Top             =   210
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Data Encerramento"
            Height          =   285
            Left            =   4380
            TabIndex        =   14
            Top             =   840
            Width           =   1425
         End
         Begin VB.Label Label3 
            Caption         =   "Data Início"
            Height          =   285
            Left            =   2340
            TabIndex        =   12
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label labrl 
            Caption         =   "Cartegoria"
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Nome do Campeonato"
            Height          =   285
            Left            =   1110
            TabIndex        =   6
            Top             =   210
            Width           =   1845
         End
         Begin VB.Label Label 
            Caption         =   "Código"
            Height          =   285
            Left            =   150
            TabIndex        =   5
            Top             =   210
            Width           =   855
         End
      End
   End
   Begin MSComctlLib.Toolbar tbBotoes 
      Height          =   570
      Left            =   15
      TabIndex        =   1
      Top             =   7875
      Width           =   6600
      _ExtentX        =   11642
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
            Object.ToolTipText     =   "Novo Campeonato"
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
            Object.ToolTipText     =   "Gravar Alterações"
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
      Left            =   120
      Top             =   6000
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
            Picture         =   "frmCadastroDeCampeonato.frx":0A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":0FCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":1568
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":1B02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":209C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":2636
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":2BD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":316A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":3704
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeCampeonato.frx":3C9E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCadastroDeCampeonato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fraEquipesCampeonato_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub fraInfoClube_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub fraPrincipal_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label1_Click()

End Sub
