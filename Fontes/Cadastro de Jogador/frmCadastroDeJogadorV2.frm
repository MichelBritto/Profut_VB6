VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Object = "{B074BC93-5A5B-11CE-98BD-0000C0E6B88E}#2.0#0"; "sstabs32.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmCadastroDeJogadorV2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProFut - Cadastro De Jogador"
   ClientHeight    =   6375
   ClientLeft      =   4515
   ClientTop       =   2355
   ClientWidth     =   10815
   Icon            =   "frmCadastroDeJogadorV2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin SSDesignerWidgetsTabs.SSIndexTab tabPrincipal 
      Height          =   5835
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   10815
      _Version        =   131078
      _ExtentX        =   19076
      _ExtentY        =   10292
      _StockProps     =   13
      BackColor       =   -2147483630
      CoverAllowClose =   0   'False
      CoverMarginX    =   200
      CoverMarginY    =   200
      RingHoleMargin  =   500
      RingMarginTop   =   100
      RingMarginBottom=   100
      RingSeparator   =   200
      RingSize        =   1000
      RingWidth       =   300
      ActualTTO       =   500
      GutterWidth     =   100
      PageAnimationFrames=   20
      TabVisibleLast  =   2
      RingCount       =   9
      RingGroups      =   3
      PageTabOrientation=   0
      PageAlignmentCaption=   7
      PageAlignmentPicture=   1
      ActiveTab3D     =   0
      BeginProperty FontSub {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ActivePageFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs(IC).PictureMetaWidth=   0
      Tabs(IC).PictureMetaHeight=   0
      Tabs(IC).Page   =   0
      Tabs(IC).ControlCount=   0
      Tabs(IC).ControlEnabled=   0   'False
      Tabs(IC).Pages(0).PictureMetaWidth=   0
      Tabs(IC).Pages(0).PictureMetaHeight=   0
      Tabs(IC).Pages(0).Tag=   ""
      Tabs(IC).Pages(0).Caption=   "Page 0"
      Tabs(IC).Pages(0).Name=   "page 0"
      Tabs(IC).Pages(0).CtlCount=   0
      Tabs(IC).Pages(0).CtlEnabled=   0   'False
      Tabs(IC).Tag    =   ""
      Tabs(IC).Caption=   ""
      Tabs(IC).Name   =   ""
      Tabs(0).PictureMetaWidth=   0
      Tabs(0).PictureMetaHeight=   0
      Tabs(0).Page    =   0
      Tabs(0).ControlCount=   0
      Tabs(0).ControlEnabled=   0   'False
      Tabs(0).Pages(0).PictureMetaWidth=   0
      Tabs(0).Pages(0).PictureMetaHeight=   0
      Tabs(0).Pages(0).Tag=   ""
      Tabs(0).Pages(0).Caption=   "Page 0"
      Tabs(0).Pages(0).Name=   "page 0"
      Tabs(0).Pages(0).CtlCount=   3
      Tabs(0).Pages(0).CtlEnabled=   -1  'True
      Tabs(0).Pages(0).Ctl(0)=   "fraInfoSistema"
      Tabs(0).Pages(0).Ctl(1)=   "fraDadosCadastrais"
      Tabs(0).Pages(0).Ctl(2)=   "fraFoto"
      Tabs(0).Tag     =   ""
      Tabs(0).Caption =   "Principal"
      Tabs(0).Name    =   "tab 0"
      Tabs(1).PictureMetaWidth=   0
      Tabs(1).PictureMetaHeight=   0
      Tabs(1).Page    =   0
      Tabs(1).ControlCount=   0
      Tabs(1).ControlEnabled=   0   'False
      Tabs(1).Pages(0).PictureMetaWidth=   0
      Tabs(1).Pages(0).PictureMetaHeight=   0
      Tabs(1).Pages(0).Tag=   ""
      Tabs(1).Pages(0).Caption=   "Page 0"
      Tabs(1).Pages(0).Name=   "page 0"
      Tabs(1).Pages(0).CtlCount=   2
      Tabs(1).Pages(0).CtlEnabled=   0   'False
      Tabs(1).Pages(0).Ctl(0)=   "fraDocumentos"
      Tabs(1).Pages(0).Ctl(1)=   "fraInfoEscolares"
      Tabs(1).Tag     =   ""
      Tabs(1).Caption =   "Documentação"
      Tabs(1).Name    =   "tab 1"
      Tabs(2).PictureMetaWidth=   0
      Tabs(2).PictureMetaHeight=   0
      Tabs(2).Page    =   0
      Tabs(2).ControlCount=   0
      Tabs(2).ControlEnabled=   0   'False
      Tabs(2).Pages(0).PictureMetaWidth=   0
      Tabs(2).Pages(0).PictureMetaHeight=   0
      Tabs(2).Pages(0).Tag=   ""
      Tabs(2).Pages(0).Caption=   "Page 0"
      Tabs(2).Pages(0).Name=   "page 0"
      Tabs(2).Pages(0).CtlCount=   1
      Tabs(2).Pages(0).CtlEnabled=   0   'False
      Tabs(2).Pages(0).Ctl(0)=   "Frame1"
      Tabs(2).Tag     =   ""
      Tabs(2).Caption =   "Endereço e Contatos"
      Tabs(2).Name    =   "tab 2"
      Templates(0).PictureMetaWidth=   0
      Templates(0).PictureMetaHeight=   0
      Templates(0).Tag=   ""
      Templates(0).Caption=   "Page 0"
      Templates(0).Name=   "page 0"
      Templates(0).CtlCount=   0
      Templates(0).CtlEnabled=   -1  'True
      Templates(1).PictureMetaWidth=   0
      Templates(1).PictureMetaHeight=   0
      Templates(1).Tag=   ""
      Templates(1).Caption=   "Page 1"
      Templates(1).Name=   "page 1"
      Templates(1).CtlCount=   0
      Templates(1).CtlEnabled=   0   'False
      Begin VB.Frame Frame1 
         Caption         =   "Endereço e Contato do Atleta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5325
         Left            =   -74940
         TabIndex        =   69
         Top             =   330
         Width           =   10695
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
            Left            =   5430
            MaxLength       =   128
            TabIndex        =   24
            Top             =   1560
            Width           =   4845
         End
         Begin VB.TextBox txtEnderecoAtleta 
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
            Left            =   360
            MaxLength       =   128
            TabIndex        =   22
            Top             =   930
            Width           =   8895
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
            Left            =   5430
            MaxLength       =   11
            TabIndex        =   27
            Top             =   2220
            Width           =   2235
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
            Left            =   360
            MaxLength       =   11
            TabIndex        =   25
            Top             =   2220
            Width           =   2145
         End
         Begin VB.TextBox txtEmail 
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
            Left            =   360
            MaxLength       =   128
            TabIndex        =   29
            Top             =   2850
            Width           =   9915
         End
         Begin VB.TextBox txtFacebookAtleta 
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
            Left            =   360
            MaxLength       =   128
            TabIndex        =   30
            Top             =   3480
            Width           =   9915
         End
         Begin VB.TextBox txtInstagramAtleta 
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
            Left            =   360
            MaxLength       =   128
            TabIndex        =   31
            Top             =   4110
            Width           =   9915
         End
         Begin VB.CheckBox chkWpp1 
            Height          =   195
            Left            =   2550
            Picture         =   "frmCadastroDeJogadorV2.frx":038A
            TabIndex        =   26
            Top             =   2310
            Width           =   255
         End
         Begin VB.CheckBox chkwpp2 
            Height          =   195
            Left            =   7710
            TabIndex        =   28
            Top             =   2340
            Width           =   225
         End
         Begin VB.PictureBox wpp2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   2820
            Picture         =   "frmCadastroDeJogadorV2.frx":0A9E
            ScaleHeight     =   345
            ScaleWidth      =   315
            TabIndex        =   71
            Top             =   2250
            Width           =   315
         End
         Begin VB.PictureBox wpp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   7950
            Picture         =   "frmCadastroDeJogadorV2.frx":0F54
            ScaleHeight     =   345
            ScaleWidth      =   315
            TabIndex        =   70
            Top             =   2250
            Width           =   315
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscUfEnderecoAtleta 
            Height          =   390
            Left            =   9330
            TabIndex        =   23
            Top             =   930
            Width           =   945
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
            _ExtentX        =   1667
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
            Left            =   360
            TabIndex        =   83
            Top             =   1560
            Width           =   5055
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
            _ExtentX        =   8916
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
         Begin VB.Label Label31 
            Caption         =   "Bairro"
            Height          =   285
            Left            =   5430
            TabIndex        =   80
            Top             =   1350
            Width           =   3885
         End
         Begin VB.Label Label13 
            Caption         =   "Cidade"
            Height          =   285
            Left            =   360
            TabIndex        =   79
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label15 
            Caption         =   "Endereço "
            Height          =   285
            Left            =   360
            TabIndex        =   78
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label14 
            Caption         =   "UF"
            Height          =   285
            Left            =   9360
            TabIndex        =   77
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Telefone/Celular 2"
            Height          =   285
            Left            =   5400
            TabIndex        =   76
            Top             =   2010
            Width           =   1815
         End
         Begin VB.Label Label17 
            Caption         =   "Telefone/Celular"
            Height          =   285
            Left            =   360
            TabIndex        =   75
            Top             =   1980
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "E-mail Contato"
            Height          =   285
            Left            =   360
            TabIndex        =   74
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label Label19 
            Caption         =   "Facebook"
            Height          =   285
            Left            =   360
            TabIndex        =   73
            Top             =   3270
            Width           =   1815
         End
         Begin VB.Label Label20 
            Caption         =   "Instagram"
            Height          =   285
            Left            =   360
            TabIndex        =   72
            Top             =   3900
            Width           =   1815
         End
      End
      Begin VB.Frame fraDocumentos 
         Caption         =   "Documentos do Atleta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   -74940
         TabIndex        =   64
         Top             =   360
         Width           =   10725
         Begin VB.TextBox txtCertidaoNascimento 
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
            TabIndex        =   13
            Top             =   420
            Width           =   5355
         End
         Begin VB.TextBox txtCartorio 
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
            Left            =   5490
            MaxLength       =   128
            TabIndex        =   14
            Top             =   420
            Width           =   5205
         End
         Begin VB.TextBox txtOrgao 
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
            Left            =   5490
            MaxLength       =   128
            TabIndex        =   16
            Top             =   1050
            Width           =   5205
         End
         Begin EditLib.fpMask txtIdentidade 
            Height          =   405
            Left            =   60
            TabIndex        =   15
            Top             =   1050
            Width           =   5355
            _Version        =   196608
            _ExtentX        =   9446
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
            Mask            =   "12 345 678-9"
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
         Begin VB.Label Label12 
            Caption         =   "Certidão de Nascimento"
            Height          =   285
            Left            =   90
            TabIndex        =   68
            Top             =   210
            Width           =   4125
         End
         Begin VB.Label Label9 
            Caption         =   "Cartório Responsável"
            Height          =   285
            Left            =   5490
            TabIndex        =   67
            Top             =   210
            Width           =   3765
         End
         Begin VB.Label Label10 
            Caption         =   "Identidade(RG)"
            Height          =   285
            Left            =   90
            TabIndex        =   66
            Top             =   840
            Width           =   4125
         End
         Begin VB.Label Label11 
            Caption         =   "Órgão Expedidor"
            Height          =   285
            Left            =   5490
            TabIndex        =   65
            Top             =   840
            Width           =   3795
         End
      End
      Begin VB.Frame fraInfoEscolares 
         Caption         =   "Informações Escolares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3765
         Left            =   -74940
         TabIndex        =   57
         Top             =   1890
         Width           =   10725
         Begin VB.TextBox txtNomeEscola 
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
            TabIndex        =   17
            Top             =   720
            Width           =   10605
         End
         Begin VB.TextBox txtEnderecoEscola 
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
            TabIndex        =   18
            Top             =   1320
            Width           =   10605
         End
         Begin VB.TextBox txtBairroEscola 
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
            Left            =   5640
            MaxLength       =   128
            TabIndex        =   20
            Top             =   1980
            Width           =   4995
         End
         Begin VB.TextBox txtRedeSocialEscola 
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
            Height          =   375
            Left            =   60
            MaxLength       =   128
            TabIndex        =   21
            Top             =   2610
            Width           =   10575
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscUfEscola 
            Height          =   390
            Left            =   90
            TabIndex        =   19
            Top             =   1980
            Width           =   945
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
            _ExtentX        =   1667
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
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscCidadeEscola 
            Height          =   390
            Left            =   1050
            TabIndex        =   82
            Top             =   1980
            Width           =   4575
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
            _ExtentX        =   8070
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
         Begin VB.Label Label21 
            Caption         =   "Nome Instituição"
            Height          =   285
            Left            =   60
            TabIndex        =   63
            Top             =   510
            Width           =   1485
         End
         Begin VB.Label Label22 
            Caption         =   "Endereço"
            Height          =   285
            Left            =   60
            TabIndex        =   62
            Top             =   1110
            Width           =   5745
         End
         Begin VB.Label Label23 
            Caption         =   "Bairro"
            Height          =   285
            Left            =   5640
            TabIndex        =   61
            Top             =   1770
            Width           =   4695
         End
         Begin VB.Label Label24 
            Caption         =   "Cidade"
            Height          =   285
            Left            =   1050
            TabIndex        =   60
            Top             =   1740
            Width           =   2235
         End
         Begin VB.Label Label25 
            Caption         =   "Rede Social"
            Height          =   285
            Left            =   60
            TabIndex        =   59
            Top             =   2400
            Width           =   1785
         End
         Begin VB.Label Label28 
            Caption         =   "UF"
            Height          =   285
            Left            =   120
            TabIndex        =   58
            Top             =   1770
            Width           =   615
         End
      End
      Begin VB.Frame fraInfoSistema 
         Caption         =   "Informações Sistema"
         Height          =   1545
         Left            =   3960
         TabIndex        =   45
         Top             =   4110
         Width           =   6825
         Begin VB.TextBox txtUsuarioCadastro 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   405
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   420
            Width           =   4785
         End
         Begin VB.TextBox txtUsuarioAlteracao 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   405
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   1080
            Width           =   4785
         End
         Begin SSCalendarWidgets_A.SSDateCombo dtcDataUltimaAlteracao 
            Height          =   405
            Left            =   4920
            TabIndex        =   47
            Top             =   1080
            Width           =   1815
            _Version        =   65543
            _ExtentX        =   3201
            _ExtentY        =   714
            _StockProps     =   93
            Enabled         =   0   'False
            Format          =   "DD/MM/YY"
            BevelType       =   0
         End
         Begin SSCalendarWidgets_A.SSDateCombo dtcDataCadastro 
            Height          =   405
            Left            =   4920
            TabIndex        =   51
            Top             =   420
            Width           =   1815
            _Version        =   65543
            _ExtentX        =   3201
            _ExtentY        =   714
            _StockProps     =   93
            Enabled         =   0   'False
            Format          =   "DD/MM/YY"
            BevelType       =   0
         End
         Begin VB.Label Label3 
            Caption         =   "Usuário cadastro"
            Height          =   285
            Left            =   60
            TabIndex        =   53
            Top             =   210
            Width           =   1965
         End
         Begin VB.Label Label1 
            Caption         =   "Data Cadastro"
            Height          =   285
            Left            =   4920
            TabIndex        =   52
            Top             =   210
            Width           =   1575
         End
         Begin VB.Label Label26 
            Caption         =   "Usuário Ultima Alteração"
            Height          =   285
            Left            =   60
            TabIndex        =   49
            Top             =   870
            Width           =   1965
         End
         Begin VB.Label Label27 
            Caption         =   "Data Ultima alteração"
            Height          =   285
            Left            =   4920
            TabIndex        =   48
            Top             =   870
            Width           =   1575
         End
      End
      Begin VB.Frame fraDadosCadastrais 
         Caption         =   "Dados Cadastrais"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3795
         Left            =   3960
         TabIndex        =   36
         Top             =   330
         Width           =   6825
         Begin VB.TextBox txtNumCamisa 
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
            Left            =   4140
            MaxLength       =   20
            TabIndex        =   3
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtNomePai 
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
            ForeColor       =   &H80000012&
            Height          =   405
            Left            =   60
            MaxLength       =   128
            TabIndex        =   11
            Top             =   2550
            Width           =   6705
         End
         Begin VB.TextBox txtNomeMae 
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
            TabIndex        =   12
            Top             =   3150
            Width           =   6705
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
            Left            =   5640
            TabIndex        =   4
            Top             =   690
            Value           =   -1  'True
            Width           =   555
         End
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
            Left            =   6180
            TabIndex        =   5
            Top             =   690
            Width           =   555
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
            TabIndex        =   6
            Top             =   1260
            Width           =   4785
         End
         Begin VB.TextBox txtCodigoInterno 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   90
            MaxLength       =   8
            TabIndex        =   1
            Top             =   600
            Width           =   945
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
            TabIndex        =   2
            Top             =   600
            Width           =   3015
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscCartegoria 
            Height          =   390
            Left            =   2700
            TabIndex        =   9
            Top             =   1890
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
         Begin SSCalendarWidgets_A.SSDateCombo dtcDataNascimento 
            Height          =   405
            Left            =   4920
            TabIndex        =   10
            Top             =   1890
            Width           =   1815
            _Version        =   65543
            _ExtentX        =   3201
            _ExtentY        =   714
            _StockProps     =   93
            Format          =   "DD/MM/YY"
            BevelType       =   0
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscClube 
            Height          =   390
            Left            =   60
            TabIndex        =   8
            Top             =   1890
            Width           =   2595
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
            _ExtentX        =   4577
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
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscPosicao 
            Height          =   390
            Left            =   4920
            TabIndex        =   7
            Top             =   1260
            Width           =   1815
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
            _ExtentX        =   3201
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
         Begin VB.Label Label32 
            Caption         =   "Posição"
            Height          =   285
            Left            =   4920
            TabIndex        =   84
            Top             =   1050
            Width           =   645
         End
         Begin VB.Label Label8 
            Caption         =   "Equipe"
            Height          =   285
            Left            =   60
            TabIndex        =   56
            Top             =   1680
            Width           =   525
         End
         Begin VB.Label Label6 
            Caption         =   "Número da camisa"
            Height          =   285
            Left            =   4140
            TabIndex        =   55
            Top             =   390
            Width           =   1365
         End
         Begin VB.Label Label4 
            Caption         =   "Nome do Pai"
            Height          =   285
            Left            =   60
            TabIndex        =   44
            Top             =   2340
            Width           =   1665
         End
         Begin VB.Label Label5 
            Caption         =   "Nome da Mãe"
            Height          =   285
            Left            =   60
            TabIndex        =   43
            Top             =   2940
            Width           =   1725
         End
         Begin VB.Label Label30 
            Caption         =   "Sexo"
            Height          =   285
            Left            =   5640
            TabIndex        =   42
            Top             =   390
            Width           =   885
         End
         Begin VB.Label Label7 
            Caption         =   "Data de Nascimento"
            Height          =   285
            Left            =   4920
            TabIndex        =   41
            Top             =   1680
            Width           =   1545
         End
         Begin VB.Label Label29 
            Caption         =   "Cartegoria"
            Height          =   285
            Left            =   2700
            TabIndex        =   40
            Top             =   1680
            Width           =   1365
         End
         Begin VB.Label Label 
            Caption         =   "Código"
            Height          =   285
            Left            =   120
            TabIndex        =   39
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Apelido 
            Caption         =   "Apelido do Jogador"
            Height          =   285
            Left            =   1080
            TabIndex        =   38
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label2 
            Caption         =   "Nome Completo"
            Height          =   285
            Left            =   90
            TabIndex        =   37
            Top             =   1050
            Width           =   1905
         End
      End
      Begin VB.Frame fraFoto 
         Height          =   5325
         Left            =   30
         TabIndex        =   32
         Top             =   330
         Width           =   3915
         Begin Threed.SSCommand cmdRemover 
            Height          =   330
            Left            =   1950
            TabIndex        =   33
            Top             =   4710
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   582
            _Version        =   196609
            PictureFrames   =   1
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmCadastroDeJogadorV2.frx":140A
            Caption         =   "Remover Foto"
            ButtonStyle     =   3
            PictureAlignment=   1
         End
         Begin Threed.SSCommand cmdAdicionar 
            Height          =   330
            Left            =   90
            TabIndex        =   34
            Top             =   4710
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   582
            _Version        =   196609
            PictureFrames   =   1
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmCadastroDeJogadorV2.frx":172C
            Caption         =   "        Adicionar Foto"
            ButtonStyle     =   3
            PictureAlignment=   1
         End
         Begin Threed.SSFrame SSFrame 
            Height          =   4305
            Index           =   1
            Left            =   90
            TabIndex        =   35
            Top             =   180
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   7594
            _Version        =   196609
            Begin VB.Label lblInativo 
               AutoSize        =   -1  'True
               Caption         =   "INATIVO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   555
               Left            =   840
               TabIndex        =   81
               Top             =   2010
               Visible         =   0   'False
               Width           =   2055
            End
            Begin VB.Image imgFotoJogador 
               Height          =   4215
               Left            =   30
               Stretch         =   -1  'True
               Top             =   30
               Width           =   3690
            End
         End
      End
   End
   Begin MSComctlLib.Toolbar tbBotoes 
      Height          =   570
      Left            =   105
      TabIndex        =   54
      Top             =   5775
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   1005
      ButtonWidth     =   2355
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
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
            Description     =   "Inativar um jogador do sistema"
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
            Enabled         =   0   'False
            Caption         =   "F8 - Imprimir"
            Key             =   "cmdimprimir"
            Object.ToolTipText     =   "Impirimir carteirinha ou Ficha do jogador"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F10-Sair"
            Key             =   "cmdSair"
            Object.ToolTipText     =   "Sair da tela"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   0
      Top             =   5670
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
            Picture         =   "frmCadastroDeJogadorV2.frx":1E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogadorV2.frx":23D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogadorV2.frx":2972
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogadorV2.frx":2F0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogadorV2.frx":34A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogadorV2.frx":3A40
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogadorV2.frx":3FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogadorV2.frx":4574
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogadorV2.frx":4B0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogadorV2.frx":50A8
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCadastroDeJogadorV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstrFlag As String
Dim mstrFoto As String
Dim mlngOpcao As Long

Dim mblnCarregando As Boolean
Dim mlngJogador As Long

Dim mbitFoto() As Byte

Dim blnRemoveuFoto As Boolean

Public Property Let DiretorioFotoJogador(strDiretorio As String)
10        mstrFoto = strDiretorio
End Property
Public Property Let OpcaoImpressao(lngOpcao As Long)
10        mlngOpcao = lngOpcao
End Property

Public Property Let Jogador(lngJogador As Long)
10        mlngJogador = lngJogador
End Property

Public Property Let Carregando(blnCarregando As Boolean)
10        mblnCarregando = blnCarregando
End Property


Private Sub cmdAdicionar_Click()
10        On Error GoTo Erro
          
20        frmAdicionarFotoJogador.Show vbModal
        
30        If mstrFoto <> "" Then
40            imgFotoJogador.Picture = Nothing
50            imgFotoJogador.Stretch = True
60            imgFotoJogador.Picture = LoadPicture(mstrFoto)
70            blnRemoveuFoto = False
80        End If
          
90        Exit Sub
Erro:
100    Call MsgBox("Erro no módulo: " & "frmCadastroDeJogador" & vbCrLf & "No Procedimento: " & "cmdAdicionar_Click" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Private Sub ImprimirFicha()
10    On Error GoTo Erro
            
20        If (Val(txtCodigoInterno.Text)) = 0 Then Exit Sub
          
30       rptFichaJogador.Parametros1 = mbitFoto()
40       rptFichaJogador.Parametros111 = (Val(txtCodigoInterno.Text))
50       rptFichaJogador.Parametros2 = (sscClube.Columns("chdescricao").Value)
60       rptFichaJogador.Parametros3 = (txtNomeJogador.Text)
70       rptFichaJogador.Parametros4 = (dtcDataNascimento.DateValue)
80       rptFichaJogador.Parametros5 = ""
90       rptFichaJogador.Parametros6 = txtCertidaoNascimento.Text
100      rptFichaJogador.Parametros7 = txtCartorio.Text
110      rptFichaJogador.Parametros8 = txtIdentidade.Text
120      rptFichaJogador.Parametros9 = txtOrgao.Text
130      rptFichaJogador.Parametros10 = txtNomePai.Text
140      rptFichaJogador.Parametros11 = txtNomeMae.Text
150      rptFichaJogador.Parametros12 = txtEnderecoAtleta.Text
160      rptFichaJogador.Parametros13 = txtBairroEnderecoAtleta.Text
170      rptFichaJogador.Parametros14 = sscCidadeEnderecoAtleta.Text
180      rptFichaJogador.Parametros15 = txtFacebookAtleta.Text
190      rptFichaJogador.Parametros16 = txtNomeEscola.Text
200      rptFichaJogador.Parametros17 = txtEnderecoEscola.Text
210      rptFichaJogador.Parametros18 = txtBairroEscola.Text
220      rptFichaJogador.Parametros19 = sscCidadeEscola.Text
230      rptFichaJogador.Parametros20 = ""
240      rptFichaJogador.Parametros21 = txtRedeSocialEscola.Text
         
250       rptFichaJogador.Show vbModal, Me

260   Exit Sub
Erro:
270      Call MsgBox("Erro no módulo: " & "frmCadastroDeJogadorV2" & vbCrLf & "No Procedimento: " & "ImprimirFicha" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Private Sub ImprimirCarteirinha()
10    On Error GoTo Erro
          
20        If (Val(txtCodigoInterno.Text)) = 0 Then Exit Sub
          
30        rptCarteirinha.Codigo = (Val(txtCodigoInterno.Text))
40        rptCarteirinha.Apelido = (txtApelido.Text)
50        rptCarteirinha.Nome = (txtNomeJogador.Text)
60        rptCarteirinha.Camisa = (Val(txtCodigoInterno.Text))
70        rptCarteirinha.Equipe = (sscClube.Columns("chdescricao").Value)
80        rptCarteirinha.Cartegoria = (sscCartegoria.Columns("chdescricao").Value)
90        rptCarteirinha.Nascimento = (dtcDataNascimento.DateValue)
100       rptCarteirinha.Mae = (txtNomePai.Text)
110       rptCarteirinha.Pai = (txtNomeMae.Text)
120       rptCarteirinha.Foto = mbitFoto()
130       rptCarteirinha.Show vbModal, Me
140   Exit Sub
Erro:
150      Call MsgBox("Erro no módulo: " & "frmCadastroDeJogadorV2" & vbCrLf & "No Procedimento: " & "ImprimirCarteirinha" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Private Sub Command1_Click()
    rptFichaJogador.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
10        Select Case KeyCode
              Case vbKeyF2:  tbBotoes.Buttons("cmdNovo").Value = tbrPressed
20            Case vbKeyF3:  tbBotoes.Buttons("cmdAlterar").Value = tbrPressed
30            Case vbKeyF4:  tbBotoes.Buttons("cmdProcurar").Value = tbrPressed
40            Case vbKeyF5:  tbBotoes.Buttons("cmdExcluir").Value = tbrPressed
50            Case vbKeyF6:  tbBotoes.Buttons("cmdLimpar").Value = tbrPressed
60            Case vbKeyF7:  tbBotoes.Buttons("cmdGravar").Value = tbrPressed
70            Case vbKeyF8:  tbBotoes.Buttons("cmdimprimir").Value = tbrPressed
80            Case vbKeyF10: tbBotoes.Buttons("cmdSair").Value = tbrPressed
90       End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

10    On Error GoTo Erro

20        tbBotoes.Buttons("cmdNovo").Value = tbrUnpressed
30        tbBotoes.Buttons("cmdAlterar").Value = tbrUnpressed
40        tbBotoes.Buttons("cmdProcurar").Value = tbrUnpressed
50        tbBotoes.Buttons("cmdExcluir").Value = tbrUnpressed
60        tbBotoes.Buttons("cmdLimpar").Value = tbrUnpressed
70        tbBotoes.Buttons("cmdGravar").Value = tbrUnpressed
80        tbBotoes.Buttons("cmdimprimir").Value = tbrUnpressed
90        tbBotoes.Buttons("cmdSair").Value = tbrUnpressed
        
100       Select Case KeyCode
              Case vbKeyF2:  If tbBotoes.Buttons("cmdNovo").Enabled Then Call tbBotoes_ButtonClick(tbBotoes.Buttons("cmdNovo"))
110           Case vbKeyF3:  If tbBotoes.Buttons("cmdAlterar").Enabled Then Call tbBotoes_ButtonClick(tbBotoes.Buttons("cmdAlterar"))
120           Case vbKeyF4:  If tbBotoes.Buttons("cmdProcurar").Enabled Then Call tbBotoes_ButtonClick(tbBotoes.Buttons("cmdProcurar"))
130           Case vbKeyF5:  If tbBotoes.Buttons("cmdExcluir").Enabled Then Call tbBotoes_ButtonClick(tbBotoes.Buttons("cmdExcluir"))
140           Case vbKeyF6:  If tbBotoes.Buttons("cmdLimpar").Enabled Then Call tbBotoes_ButtonClick(tbBotoes.Buttons("cmdLimpar"))
150           Case vbKeyF7:  If tbBotoes.Buttons("cmdGravar").Enabled Then Call tbBotoes_ButtonClick(tbBotoes.Buttons("cmdGravar"))
160           Case vbKeyF8:  If tbBotoes.Buttons("cmdimprimir").Enabled Then Call tbBotoes_ButtonClick(tbBotoes.Buttons("cmdimprimir"))
170           Case vbKeyF10: If tbBotoes.Buttons("cmdSair").Enabled Then Call tbBotoes_ButtonClick(tbBotoes.Buttons("cmdSair"))
              
180           Case vbKeyRight
190               On Error Resume Next
200               tabPrincipal.ActiveTab = tabPrincipal.ActiveTab + 1
210               On Error GoTo Erro
220           Case vbKeyLeft
230                On Error Resume Next
240               tabPrincipal.ActiveTab = tabPrincipal.ActiveTab - 1
250               On Error GoTo Erro
260       End Select
          
270   Exit Sub
Erro:
280      Call MsgBox("Erro no módulo: " & "frmCadastroDeJogadorV2" & vbCrLf & "No Procedimento: " & "Form_KeyUp" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub
Private Sub cmdRemover_Click()
10        imgFotoJogador.Picture = Nothing
20        mstrFoto = ""
30        blnRemoveuFoto = True
End Sub

Private Sub Form_Load()
          
10        mstrFlag = ""
          
20        If mblnCarregando = True Then
30            txtCodigoInterno.Text = mlngJogador
40            Call txtCodigoInterno_KeyDown(vbKeyReturn, 0)
50        Else
60            Call LimparCampos
70            Call HabilitarTBBotoes(True, False, True, False, False, True, False, False)
80            Call HabilitarCampos(False)
90        End If
End Sub

Private Sub LimparCampos()
          
10        LimparArray mbitFoto
          
20        mstrFlag = ""
30        blnRemoveuFoto = False
          
40        txtCodigoInterno.Text = ""
50        txtApelido.Text = ""
60        txtNomeJogador.Text = ""
70        sscClube.Text = ""
80        sscPosicao.Text = ""
          'txtLocalDeNascimento.Text = ""
90        txtNomePai.Text = ""
100       txtNomeMae.Text = ""
110       txtUsuarioAlteracao.Text = ""
120       txtCertidaoNascimento.Text = ""
130       txtCartorio.Text = ""
140       txtIdentidade.Text = ""
150       txtOrgao.Text = ""
160       txtNomeEscola.Text = ""
170       txtEnderecoEscola.Text = ""
180       sscCidadeEscola.Text = ""
190       txtBairroEscola.Text = ""
200       txtRedeSocialEscola.Text = ""
210       txtEnderecoAtleta.Text = ""
220       sscCidadeEnderecoAtleta.Text = ""
230       txtBairroEnderecoAtleta.Text = ""
240       txtTelCel1.Text = ""
250       txtTelCel2.Text = ""
260       txtEmail.Text = ""
270       txtFacebookAtleta.Text = ""
280       txtInstagramAtleta.Text = ""
290       txtNumCamisa.Text = ""
300       txtUsuarioCadastro.Text = ""

          
310       optMasculino.Value = vbChecked
320       optFeminino.Value = vbUnchecked
          
330       chkWpp1.Value = vbUnchecked
340       chkwpp2.Value = vbUnchecked
          
350       dtcDataUltimaAlteracao.DateValue = Nothing
360       dtcDataNascimento.DateValue = Nothing
370       dtcDataCadastro.DateValue = Nothing
          
380       sscUfEscola.Text = ""
390       sscUfEnderecoAtleta.Text = ""
400       mstrFoto = ""
410       imgFotoJogador.Picture = Nothing
          
420       modBDCombo_SelecionarCartecoriaJogador sscCartegoria
430       modBDCombo_SelecionarEstados sscUfEscola
440       modBDCombo_SelecionarEstados sscUfEnderecoAtleta
450       modBDCombo_SelecionarEquipePorCodigo sscClube
460       modBDCombo_SelecionarPosicoesAtleta sscPosicao
          
      '    modBDCombo_SelecionarCidades sscCidadeEnderecoAtleta
      '    modBDCombo_SelecionarCidades sscCidadeEscola
          
End Sub

Private Sub HabilitarCampos(blnhabilitar As Boolean)
          
10        txtCodigoInterno.Locked = blnhabilitar
20        txtApelido.Locked = Not blnhabilitar
30        txtNomeJogador.Locked = Not blnhabilitar
          
          'txtLocalDeNascimento.Locked = Not blnhabilitar
40        txtNomePai.Locked = Not blnhabilitar
50        txtNomeMae.Locked = Not blnhabilitar
60        txtUsuarioAlteracao.Locked = True
70        txtCertidaoNascimento.Enabled = blnhabilitar
80        txtCartorio.Locked = Not blnhabilitar
90        txtIdentidade.Enabled = blnhabilitar
100       txtOrgao.Locked = Not blnhabilitar
110       txtNomeEscola.Locked = Not blnhabilitar
120       txtEnderecoEscola.Locked = Not blnhabilitar
130       sscCidadeEscola.Enabled = blnhabilitar
140       txtBairroEscola.Locked = Not blnhabilitar
150       txtRedeSocialEscola.Locked = Not blnhabilitar
160       txtEnderecoAtleta.Locked = Not blnhabilitar
170       sscCidadeEnderecoAtleta.Enabled = blnhabilitar
180       txtBairroEnderecoAtleta.Locked = Not blnhabilitar
190       txtTelCel1.Locked = Not blnhabilitar
200       txtTelCel2.Locked = Not blnhabilitar
210       txtEmail.Locked = Not blnhabilitar
220       txtFacebookAtleta.Locked = Not blnhabilitar
230       txtInstagramAtleta.Locked = Not blnhabilitar
240       txtNumCamisa.Locked = Not blnhabilitar
          
250       chkWpp1.Enabled = blnhabilitar
260       chkwpp2.Enabled = blnhabilitar
          
270       optMasculino.Enabled = blnhabilitar
280       optFeminino.Enabled = blnhabilitar
          
290       dtcDataNascimento.Enabled = blnhabilitar
          
300       sscClube.Enabled = blnhabilitar
310       sscPosicao.Enabled = blnhabilitar
320       sscCartegoria.Enabled = blnhabilitar
330       sscUfEscola.Enabled = blnhabilitar
340       sscUfEnderecoAtleta.Enabled = blnhabilitar
          
350       cmdAdicionar.Enabled = blnhabilitar
360       cmdRemover.Enabled = blnhabilitar
End Sub

Private Sub HabilitarTBBotoes(blnNovo As Boolean, blnAlterar As Boolean, blnProcurar As Boolean, blnAbandonar As Boolean, blnGravar As Boolean, blnSair As Boolean, blnImprimir As Boolean, blnExcluir As Boolean)

10        tbBotoes.Buttons("cmdNovo").Enabled = blnNovo
20        tbBotoes.Buttons("cmdAlterar").Enabled = blnAlterar
30        tbBotoes.Buttons("cmdProcurar").Enabled = blnProcurar
40        tbBotoes.Buttons("cmdLimpar").Enabled = blnAbandonar
50        tbBotoes.Buttons("cmdGravar").Enabled = blnGravar
60        tbBotoes.Buttons("cmdSair").Enabled = blnSair
70        tbBotoes.Buttons("cmdimprimir").Enabled = blnImprimir
80        tbBotoes.Buttons("cmdExcluir").Enabled = blnExcluir
          
End Sub


Private Sub sscUfEnderecoAtleta_LostFocus()
10        Call modBDCombo_SelecionarCidades(sscCidadeEnderecoAtleta, , sscUfEnderecoAtleta.Columns("chcodigo").Value)
End Sub

Private Sub sscUfEscola_LostFocus()
10        Call modBDCombo_SelecionarCidades(sscCidadeEscola, , sscUfEscola.Columns("chcodigo").Value)
End Sub

Private Sub tbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
10        If Not (Button.Enabled) Then Exit Sub
20        Select Case Button.Key
              
              Case "cmdNovo":
30                If RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 2) = True Then
40                    mstrFlag = "I"
50                    Call HabilitarCampos(True)
60                    Call HabilitarTBBotoes(False, False, False, True, True, False, False, False)
70                    txtApelido.SetFocus
80                Else
90                    MsgBox "Permissão requerida!" & vbCrLf & "-> Permissão Nº2" & vbCrLf & vbCrLf & "Entre em contato com o administrador para liberar a permissão!", vbOKOnly + vbExclamation, "Permissão negada!"
100               End If
              
110           Case "cmdAlterar":
                  
120               If RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 3) = True Then
130                   mstrFlag = "A"
140                   Call HabilitarCampos(True)
150                   Call HabilitarTBBotoes(False, False, False, True, True, False, False, False)
160                   txtApelido.SetFocus
170               Else
180                   MsgBox "Permissão requerida!" & vbCrLf & "-> Permissão Nº3" & vbCrLf & vbCrLf & "Entre em contato com o administrador para liberar a permissão!", vbOKOnly + vbExclamation, "Permissão negada!"
190               End If
                  
200           Case "cmdExcluir":
                  
210               If lblInativo.Visible = False Then
220                   If MsgBox("Deseja INATIVAR o jogador no sistema?" & vbCrLf & vbCrLf & "O jogador será removido temporáriamente do sistema mas ainda poderá ser reativado.", vbYesNo + vbInformation, "Atenção!") = vbYes Then
230                       Call modJogador_ApagarJogadorPorCodigo(Val(txtCodigoInterno.Text), 1)
240                       Call CarregarJogador(Val(txtCodigoInterno.Text))
250                   End If
260               Else
270                   If MsgBox("Deseja EXCLUIR o jogador no sistema?" & vbCrLf & vbCrLf & "O jogador será removido permanentemente do sistema e NÃO poderá ser reativado.", vbYesNo + vbInformation, "Atenção!") = vbYes Then
280                       Call modJogador_ApagarJogadorPorCodigo(Val(txtCodigoInterno.Text), 2)
290                       mstrFlag = ""
300                       Call LimparCampos
310                       Call HabilitarCampos(False)
320                       Call HabilitarTBBotoes(True, False, True, False, False, True, True, False)
330                       txtCodigoInterno.SetFocus
340                   End If
350               End If
                  
              
360           Case "cmdLimpar":
370               mstrFlag = ""
380               Call LimparCampos
390               Call HabilitarCampos(False)
400               Call HabilitarTBBotoes(True, False, True, False, False, True, True, False)
410                txtCodigoInterno.SetFocus
              
420           Case "cmdGravar"
430               If VerificarCampos Then
440                   GravarJogador
450                   CarregarJogador Val(txtCodigoInterno.Text)
460                   mstrFlag = ""
470               Else
480                   Exit Sub
490               End If
500               Call HabilitarCampos(False)
510               Call HabilitarTBBotoes(False, True, True, True, False, False, True, True)
520               txtCodigoInterno.SetFocus
                  
530           Case "cmdProcurar"
540               If tbBotoes.Buttons("cmdProcurar").Caption = "F4 - Procurar" Then
                      Dim ObjRelatorioJogador As ClsRelJogador
550                   Set ObjRelatorioJogador = New ClsRelJogador
                      
560                   If Not gSMConexao Is Nothing Then
570                       If gSMConexao.EstadoConexaoBD = adStateOpen Then
                              
580                           ObjRelatorioJogador.Show gSMConexao, "ProFut - Relatório de Jogador", vbModal, Me, True
590                           txtCodigoInterno.Text = ObjRelatorioJogador.ID
600                           txtCodigoInterno_KeyDown vbKeyReturn, 0
610                           Exit Sub
620                       Else
630                           gSMConexao.conectar
640                       End If
650                   End If
660               Else
670                   If MsgBox("Deseja REATIVAR o jogador no sistema?" & vbCrLf & vbCrLf & "O jogador será ativado novamente no sistema.", vbYesNo + vbInformation, "Atenção!") = vbYes Then
680                       Call modJogador_ApagarJogadorPorCodigo(Val(txtCodigoInterno.Text), 3)
690                       Call CarregarJogador(Val(txtCodigoInterno.Text))
700                   End If
710               End If
              
720           Case "cmdimprimir"
              
730               If RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 4) = True Then
740                   frmOpcaoImpressao.Show vbModal, Me
750                   Select Case mlngOpcao
                          Case 1
760                           ImprimirFicha
770                       Case 2
780                           ImprimirCarteirinha
790                   End Select
800               Else
810                   MsgBox "Permissão requerida!" & vbCrLf & "-> Permissão Nº4" & vbCrLf & vbCrLf & "Entre em contato com o administrador para liberar a permissão!", vbOKOnly + vbExclamation, "Permissão negada!"
820               End If
830           Case "cmdSair"
840               Unload Me
              
850       End Select
End Sub

Private Sub txtCertidaoNascimento_KeyPress(KeyAscii As Integer)
10        TextBoxSomenteNumeros txtCertidaoNascimento.Text, KeyAscii, False, False
End Sub

Private Sub txtCodigoInterno_KeyDown(KeyCode As Integer, Shift As Integer)
10        If KeyCode = vbKeyReturn Then
20            CarregarJogador Val(txtCodigoInterno.Text)
30        End If
End Sub

Private Sub txtCodigoInterno_KeyPress(KeyAscii As Integer)
10        TextBoxSomenteNumeros txtCodigoInterno.Text, KeyAscii, False, False
End Sub

Private Sub txtEmail_LostFocus()
10        EmailValido txtEmail.Text
End Sub

Private Sub txtApelido_Change()
10        If KeyCode = vbKeyReturn Then
              'chamo procurar jogador aqui
20        End If
End Sub

Private Sub txtApelido_KeyPress(KeyAscii As Integer)
    'TextBoxSomenteNumeros txtApelido.Text, KeyAscii, False, False
End Sub

Private Sub txtIdentidade_KeyPress(KeyAscii As Integer)
10        TextBoxSomenteNumeros txtIdentidade.Text, KeyAscii, False, False
End Sub

Private Sub txtTelCel1_KeyPress(KeyAscii As Integer)
10        TextBoxSomenteNumeros txtTelCel1.Text, KeyAscii, False, False
End Sub

Private Sub txtTelCel1_Validate(Cancel As Boolean)
10        FormataTelefone txtTelCel1.Text, True
End Sub

Private Sub txtTelCel2_KeyPress(KeyAscii As Integer)
10        TextBoxSomenteNumeros txtTelCel2.Text, KeyAscii, False, False
End Sub


Private Function VerificarCampos()
10    On Error GoTo Erro
          Dim blnContinua As Boolean
          Dim strMensagem As String
          
20        blnContinua = True
      '------------INFORMAÇÕES BASICAS------------
30        If txtApelido.Text = "" Then
40            strMensagem = strMensagem & "-> Número de inscrição não preenchido." & vbCrLf
50            blnContinua = False
60        End If
          
70        If txtNomeJogador.Text = "" Then
80            strMensagem = strMensagem & "-> Nome do jogador não preenchido." & vbCrLf
90            blnContinua = False
100       End If
          
110       If sscClube.Text = "" Then
120           strMensagem = strMensagem & "-> Equipe não informada ou inválida." & vbCrLf
130           blnContinua = False
140       End If
          
150       If sscPosicao.Text = "" Then
160           strMensagem = strMensagem & "-> Posição não informada ou inválida." & vbCrLf
170           blnContinua = False
180       End If
          
190       If txtNomePai.Text = "" Then
200           strMensagem = strMensagem & "-> Nome do pai não informado." & vbCrLf
210           blnContinua = False
220       End If
          
230       If txtNomeMae.Text = "" Then
240           strMensagem = strMensagem & "-> Nome da mãe não informado." & vbCrLf
250           blnContinua = False
260       End If
          
270       If txtNomeMae.Text = "" Then
280           strMensagem = strMensagem & "-> Nome da mãe não informado." & vbCrLf
290           blnContinua = False
300       End If
      '-----------------Documentos do Atleta------------------
      '    If txtCertidaoNascimento.Text = "" And txtIdentidade.Text = "" Then
      '        strMensagem = strMensagem & "-> Documento do jogador não informado." & vbCrLf
      '        blnContinua = False
      '    End If
      '
      '    If txtCertidaoNascimento.Text <> "" And txtCartorio.Text = "" Then
      '        strMensagem = strMensagem & "-> Cartório responsável pela certidão não informado." & vbCrLf
      '        blnContinua = False
      '    End If
      '
      '    If txtIdentidade.Text <> "" And txtOrgao.Text = "" Then
      '        strMensagem = strMensagem & "-> Órgão expedidor não informado." & vbCrLf
      '        blnContinua = False
      '    End If
      '
      '----------------Informações escolares-------------------
      'OBRIGATÓRIO SOMENTE MEDIANTE AO PREENCHIMENTO DO NOME DA ESCOLA
310       If txtNomeEscola <> "" Then
          
320           If txtEnderecoEscola.Text = "" Then
330               strMensagem = strMensagem & "-> Endereço da escola não informado." & vbCrLf
340               blnContinua = False
350           End If
                  
360           If sscUfEnderecoAtleta.Text = "" Then
370               strMensagem = strMensagem & "-> Estado da escola não informado." & vbCrLf
380               blnContinua = False
390           End If
                          
400           If sscCidadeEscola.Text = "" Then
410               strMensagem = strMensagem & "-> Cidade da escola não informado." & vbCrLf
420               blnContinua = False
430           End If
              
440           If txtBairroEscola.Text = "" Then
450               strMensagem = strMensagem & "-> Bairro da escola não informado." & vbCrLf
460               blnContinua = False
470           End If
480       End If
          
      '-------------Endereço/Contato do Atleta------------------
490       If txtEnderecoAtleta.Text = "" Then
500           strMensagem = strMensagem & "-> Endereço do atleta não informado." & vbCrLf
510           blnContinua = False
520       End If
          
530       If sscUfEnderecoAtleta.Text = "" Then
540           strMensagem = strMensagem & "-> Estado do atleta não informado." & vbCrLf
550           blnContinua = False
560       End If
          
570       If sscCidadeEnderecoAtleta.Text = "" Then
580           strMensagem = strMensagem & "-> Cidade do atleta não informado." & vbCrLf
590           blnContinua = False
600       End If
          
610       If txtBairroEnderecoAtleta.Text = "" Then
620           strMensagem = strMensagem & "-> Bairro do atleta não informado." & vbCrLf
630           blnContinua = False
640       End If
          
650       If txtTelCel1.Text = "" And txtTelCel2.Text = "" Then
660           strMensagem = strMensagem & "-> O atleta precisa ter pelo menos um telefone de contato." & vbCrLf
670           blnContinua = False
680       End If
          
          
690       If Not blnContinua Then
              
700           MsgBox "O jogador não pode ser gravado pois possuí as seguintes pendências: " & vbCrLf & strMensagem, vbOKOnly + vbInformation, "Atenção!"
              
710       End If
          
720       VerificarCampos = blnContinua

730   Exit Function
Erro:
740      Call MsgBox("Erro no módulo: " & "frmCadastroDeJogador" & vbCrLf & "No Procedimento: " & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Function

Public Sub SalvaImagem(ByRef f() As Byte, File As String)
          Dim b() As Byte
          Dim ff  As Long
          Dim n   As Long
          
10        On Error GoTo ErrHandler
20        ff = FreeFile
30        Open File For Binary Access Read As ff
40        n = LOF(ff)
50        If n Then
60           ReDim b(1 To n) As Byte
70           Get ff, , b()
80        End If
90        Close ff
100       f() = b()
110       Exit Sub
          
ErrHandler:
120       MsgBox "ERROR: " & Err.Description
End Sub

Private Sub GravarJogador()
10    On Error GoTo Erro
            
          Dim udtJogador As TypJogador
          Dim binIMG() As Byte
          
          'COLOCO A IMAGEM EM CÓDIGO BINÁRIO
20        If mstrFoto <> "" Then
30            SalvaImagem binIMG(), mstrFoto
40        End If
          
50        With udtJogador
          
60            .lngCodigo = IIf(txtCodigoInterno.Text <> "", Val(txtCodigoInterno.Text), 0)
70            .strApelido = txtApelido.Text
80            .strNomeAtleta = txtNomeJogador.Text
90            .lngCartegoria = sscCartegoria.Columns("chcodigo").Value
100           .lngEquipe = sscClube.Columns(1).Value
110           .lngPosicao = sscPosicao.Columns(1).Value
              '.strLocalNascimento = txtLocalDeNascimento.Text
120           .strNomePai = txtNomePai.Text
130           .strNomeMae = txtNomeMae.Text
140           .strCertidaoNascimento = txtCertidaoNascimento.Text
150           .strCartorio = txtCartorio.Text
160           .strIdentidade = txtIdentidade.Text
170           .strOrgaoIdentidade = txtOrgao.Text
180           .datDataNascimento = dtcDataNascimento.DateValue
190           If txtNomeEscola.Text <> "" Then
200               .strEscola = txtNomeEscola.Text
210               .lngEstadoEscola = sscUfEscola.Columns(1).Value
220               .strCidadeEscola = sscCidadeEscola.Text
230               .strBairroEscola = txtBairroEscola.Text
240               .strFacebookEscola = txtRedeSocialEscola.Text
250           End If
260           .lngEstado = sscUfEnderecoAtleta.Columns("chcodigo").Value
270           .strEndereco = txtEnderecoAtleta.Text
280           .strCidade = sscCidadeEnderecoAtleta.Text
290           .strBairro = txtBairroEnderecoAtleta.Text
300           .strTelefone1 = txtTelCel1.Text
310           .strTelefone2 = txtTelCel2.Text
320           .blnWpp1 = IIf(chkWpp1.Value = vbChecked, True, False)
330           .blnWpp2 = IIf(chkwpp2.Value = vbChecked, True, False)
340           .strEmailContato = txtEmail.Text
350           .strFacebook = txtFacebookAtleta
360           .strInstagram = txtInstagramAtleta.Text
370           .strEnderecoImagem() = IIf(mstrFoto <> "", binIMG(), mbitFoto())
380           .lngSexo = IIf(optMasculino.Value = True, 1, 2)
390           .lngNumeroCamisa = Val(txtNumCamisa.Text)
400           .blnTemImagem = IIf(blnRemoveuFoto = True, False, True)
410       End With
          
          
420       If mstrFlag = "I" Then
430           Call modJogador_AdicionarJogador(udtJogador)
             
440       ElseIf mstrFlag = "A" Then
450           Call modJogador_AlterarJogador(udtJogador)
460       End If
          
          
470       txtCodigoInterno.Text = udtJogador.lngCodigo
          


480   Exit Sub
Erro:
490       Call MsgBox("Erro no módulo: " & "frmCadastroDeJogador" & vbCrLf & "No Procedimento: " & "GravarJogador" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Private Sub CarregarJogador(lngCodigo As Long)
10    On Error GoTo Erro
          Dim objRsJogador As Recordset
          Dim binIMG() As Byte
20        Set objRsJogador = New Recordset
            
30        Call LimparCampos
40        modJogador_SelecionarJogadorPorCodigo lngCodigo, objRsJogador
          
          
50        If Not objRsJogador Is Nothing Then
60            If Not objRsJogador.EOF And Not objRsJogador.BOF Then
              
              
70                If RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 12) = False Then
80                    If RetornaClubePorUsuario(gSMConexao.CodigoUsuario) <> NZ(objRsJogador!EQUIPE_IN) Then
90                        MsgBox "Você não tem permissão para visualizar o jogador pois ele não pertence a sua equipe.", vbOKOnly + vbExclamation, "Atenção!"
100                       Exit Sub
110                   End If
120               End If
                  
130               txtCodigoInterno.Text = NZ(objRsJogador!ID_JOGADOR_IN)
140               txtApelido.Text = NS(objRsJogador!APELIDO_VC)
150               txtNomeJogador.Text = NS(objRsJogador!NOMEATLETA_VC)
160               txtNumCamisa.Text = NZ(objRsJogador!NUMEROCAMISA_IN)
170               txtNomePai.Text = NS(objRsJogador!NOMEPAI_VC)
180               txtNomeMae.Text = NS(objRsJogador!STRNOMEMAE_VC)
190               txtUsuarioAlteracao.Text = NS(objRsJogador!USUARIOULTIMAALTERACAO_VC)
200               txtUsuarioCadastro.Text = NS(objRsJogador!USUARIOCADASTRO_VC)
210               txtCertidaoNascimento.Text = NS(objRsJogador!CERTIDAONASCIMENTO_VC)
220               txtCartorio.Text = NS(objRsJogador!CARTORIO_VC)
230               txtIdentidade.Text = NS(objRsJogador!IDENTIDADE_VC)
240               txtOrgao.Text = NS(objRsJogador!ORGAOIDENTIDADE_VC)
250               txtNomeEscola.Text = NS(objRsJogador!ESCOLA_VC)
260               txtEnderecoEscola.Text = NS(objRsJogador!ENDERECOESCOLA_VC)
270               sscCidadeEscola.Text = NS(objRsJogador!CIDADEESCOLA_VC)
280               txtBairroEscola.Text = NS(objRsJogador!BAIRROESCOLA_VC)
290               txtRedeSocialEscola.Text = NS(objRsJogador!REDESOCIALESCOLA_VC)
300               txtEnderecoAtleta.Text = NS(objRsJogador!ENDERECO_VC)
310               sscCidadeEnderecoAtleta.Text = NS(objRsJogador!CIDADE_VC)
320               txtBairroEnderecoAtleta.Text = NS(objRsJogador!BAIRRO_VC)
330               txtTelCel1.Text = NS(objRsJogador!TELCEL1_VC)
340               txtTelCel2.Text = NS(objRsJogador!TELCEL2_VC)
350               txtEmail.Text = NS(objRsJogador!EMAIL_VC)
360               txtFacebookAtleta.Text = NS(objRsJogador!FACEBOOK_VC)
370               txtInstagramAtleta.Text = NS(objRsJogador!INSTAGRAM_VC)
                  
380               If NZ(objRsJogador!SEXO_IN) = 1 Then
390                   optMasculino.Value = True
400                   optFeminino.Value = False
410               Else
420                   optMasculino.Value = False
430                   optFeminino.Value = True
440               End If
                  
450               chkWpp1.Value = IIf(NB(objRsJogador!WPP1_BT), vbChecked, vbUnchecked)
460               chkwpp2.Value = IIf(NB(objRsJogador!WPP2_BT), vbChecked, vbUnchecked)
                  
470               dtcDataCadastro.DateValue = ND(objRsJogador!DATACADASTRO_DT)
480               dtcDataUltimaAlteracao.DateValue = ND(objRsJogador!DATAULTIMAALTERACAO_DT)
490               dtcDataNascimento.DateValue = ND(objRsJogador!DATANASCIMENTO_DT)
                  
                  
500               lblInativo.Visible = NB(objRsJogador!EXCLUIDO_BT)
                  
510               If NB(objRsJogador!EXCLUIDO_BT) = True Then
520                   tbBotoes.Buttons("cmdExcluir").Caption = "F5 - Excluir"
530                   tbBotoes.Buttons("cmdProcurar").Caption = "F4 - Reativar"
540                   tbBotoes.Buttons("cmdProcurar").Image = 2
550               Else
560                   tbBotoes.Buttons("cmdExcluir").Caption = "F5 - Inativar"
570                   tbBotoes.Buttons("cmdProcurar").Caption = "F4 - Procurar"
580                   tbBotoes.Buttons("cmdProcurar").Image = 4
590               End If
                  
      '            mstrFoto = NS(objRsJogador!ENDERECOIMAGEM_VC)
      '            If mstrFoto <> "" Then
      '                imgFotoJogador.Picture = Nothing
      '                imgFotoJogador.Stretch = True
      '                On Error Resume Next
      '                imgFotoJogador.Picture = LoadPicture(mstrFoto)
      '                On Error GoTo Erro
      '            End If

      '--------------------------------------------------------------------------------------
600               If Not IsNull(objRsJogador!ENDERECOIMAGEM_VC) Then
                      
610                   mbitFoto() = objRsJogador!ENDERECOIMAGEM_VC
620                   binIMG() = objRsJogador!ENDERECOIMAGEM_VC
630                   If Val(binIMG(1)) <> 0 Then
640                       imgFotoJogador.Picture = Nothing
650                       imgFotoJogador.Stretch = True
660                       On Error Resume Next

                          Dim b()  As Byte
                          Dim ff   As Long
                          Dim Arquivo As String
                      
                          'On Error GoTo ErrHandler
                          'Call GetRandomArquivoName(Arquivo)
670                       Arquivo = "tempimg.bmp"
680                       ff = FreeFile
690                       Open Arquivo For Binary Access Write As ff
700                       b() = binIMG()
710                       Put ff, , b()
720                       Close ff
730                       Erase b
740                       imgFotoJogador.Picture = LoadPicture(Arquivo)
                          'Set GetImageFromField = LoadPicture(Arquivo)
750                       Kill Arquivo
760                       End If
770                   End If
      '--------------------------------------------------------------------------------------
                  
780               modBDCombo_SelecionarCartecoriaJogador sscCartegoria, NZ(objRsJogador!CARTEGORIA_IN)
790               modBDCombo_SelecionarEquipePorCodigo sscClube, NZ(objRsJogador!EQUIPE_IN)
800               modBDCombo_SelecionarPosicoesAtleta sscPosicao, NZ(objRsJogador!POSICAO_IN)
810               modBDCombo_SelecionarEstados sscUfEscola, NZ(objRsJogador!ESTADOESCOLA_IN)
820               modBDCombo_SelecionarEstados sscUfEnderecoAtleta, NZ(objRsJogador!Estado_IN)
                  
830               mstrFlag = ""
840               Call HabilitarCampos(False)
850               Call HabilitarTBBotoes(False, True, True, True, False, True, True, True)
                  
860           Else
870               MsgBox "Jogador não encontrado ou código inválido.", vbOKOnly + vbInformation, "Atenção!"
880           End If
890       Else
900           MsgBox "Jogador não encontrado ou código inválido.", vbOKOnly + vbInformation, "Atenção!"
910       End If

920   Exit Sub
Erro:
930      Call MsgBox("Erro no módulo: " & "frmCadastroDeJogador" & vbCrLf & "No Procedimento: " & "CarregarJogador" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Private Sub txtTelCel2_Validate(Cancel As Boolean)
10        FormataTelefone txtTelCel2.Text, True
End Sub

Private Sub LimparArray(arr As Variant)
    Dim i As Long
    

    
'    Do While i < Len(arr)
'
'        arr(i) = Empty
'
'        i = i + 1
'    Loop
End Sub
