VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Object = "{B074BC93-5A5B-11CE-98BD-0000C0E6B88E}#2.0#0"; "sstabs32.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmCadastroDeJogadorV2 
   Caption         =   "ProFut - Cadastro De Jogador"
   ClientHeight    =   6375
   ClientLeft      =   4515
   ClientTop       =   2355
   ClientWidth     =   10815
   Icon            =   "frmCadastroDeJogadorV2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   10815
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
      Tabs(0).Pages(0).Ctl(0)=   "fraFoto"
      Tabs(0).Pages(0).Ctl(1)=   "fraDadosCadastrais"
      Tabs(0).Pages(0).Ctl(2)=   "fraInfoSistema"
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
      Tabs(1).Pages(0).Ctl(0)=   "fraInfoEscolares"
      Tabs(1).Pages(0).Ctl(1)=   "fraDocumentos"
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
         TabIndex        =   70
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
            TabIndex        =   25
            Top             =   1560
            Width           =   4845
         End
         Begin VB.TextBox txtCidadeEnderecoAtleta 
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
            TabIndex        =   24
            Top             =   1560
            Width           =   5025
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
            TabIndex        =   28
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
            TabIndex        =   26
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
            TabIndex        =   30
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
            TabIndex        =   31
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
            TabIndex        =   32
            Top             =   4110
            Width           =   9915
         End
         Begin VB.CheckBox chkWpp1 
            Height          =   195
            Left            =   2550
            Picture         =   "frmCadastroDeJogadorV2.frx":038A
            TabIndex        =   27
            Top             =   2310
            Width           =   255
         End
         Begin VB.CheckBox chkwpp2 
            Height          =   195
            Left            =   7710
            TabIndex        =   29
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
            TabIndex        =   72
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
            TabIndex        =   71
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
         Begin VB.Label Label31 
            Caption         =   "Bairro"
            Height          =   285
            Left            =   5430
            TabIndex        =   81
            Top             =   1350
            Width           =   3885
         End
         Begin VB.Label Label13 
            Caption         =   "Cidade"
            Height          =   285
            Left            =   360
            TabIndex        =   80
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label15 
            Caption         =   "Endereço "
            Height          =   285
            Left            =   360
            TabIndex        =   79
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label14 
            Caption         =   "UF"
            Height          =   285
            Left            =   9360
            TabIndex        =   78
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Telefone/Celular 2"
            Height          =   285
            Left            =   5400
            TabIndex        =   77
            Top             =   2010
            Width           =   1815
         End
         Begin VB.Label Label17 
            Caption         =   "Telefone/Celular"
            Height          =   285
            Left            =   360
            TabIndex        =   76
            Top             =   1980
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "E-mail Contato"
            Height          =   285
            Left            =   360
            TabIndex        =   75
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label Label19 
            Caption         =   "Facebook"
            Height          =   285
            Left            =   360
            TabIndex        =   74
            Top             =   3270
            Width           =   1815
         End
         Begin VB.Label Label20 
            Caption         =   "Instagram"
            Height          =   285
            Left            =   360
            TabIndex        =   73
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
         TabIndex        =   65
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
            TabIndex        =   12
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
            TabIndex        =   13
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
            TabIndex        =   15
            Top             =   1050
            Width           =   5205
         End
         Begin EditLib.fpMask txtIdentidade 
            Height          =   405
            Left            =   60
            TabIndex        =   14
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
            TabIndex        =   69
            Top             =   210
            Width           =   4125
         End
         Begin VB.Label Label9 
            Caption         =   "Cartório Responsável"
            Height          =   285
            Left            =   5490
            TabIndex        =   68
            Top             =   210
            Width           =   3765
         End
         Begin VB.Label Label10 
            Caption         =   "Identidade(RG)"
            Height          =   285
            Left            =   90
            TabIndex        =   67
            Top             =   840
            Width           =   4125
         End
         Begin VB.Label Label11 
            Caption         =   "Órgão Expedidor"
            Height          =   285
            Left            =   5490
            TabIndex        =   66
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
         TabIndex        =   58
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
            TabIndex        =   16
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
            TabIndex        =   17
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
         Begin VB.TextBox txtCidadeEscola 
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
            Left            =   1050
            MaxLength       =   128
            TabIndex        =   19
            Top             =   1980
            Width           =   4545
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
            TabIndex        =   18
            Top             =   1980
            Width           =   945
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
         Begin VB.Label Label21 
            Caption         =   "Nome Instituição"
            Height          =   285
            Left            =   60
            TabIndex        =   64
            Top             =   510
            Width           =   1485
         End
         Begin VB.Label Label22 
            Caption         =   "Endereço"
            Height          =   285
            Left            =   60
            TabIndex        =   63
            Top             =   1110
            Width           =   5745
         End
         Begin VB.Label Label23 
            Caption         =   "Bairro"
            Height          =   285
            Left            =   5640
            TabIndex        =   62
            Top             =   1770
            Width           =   4695
         End
         Begin VB.Label Label24 
            Caption         =   "Cidade"
            Height          =   285
            Left            =   1050
            TabIndex        =   61
            Top             =   1740
            Width           =   2235
         End
         Begin VB.Label Label25 
            Caption         =   "Rede Social"
            Height          =   285
            Left            =   60
            TabIndex        =   60
            Top             =   2400
            Width           =   1785
         End
         Begin VB.Label Label28 
            Caption         =   "UF"
            Height          =   285
            Left            =   120
            TabIndex        =   59
            Top             =   1770
            Width           =   615
         End
      End
      Begin VB.Frame fraInfoSistema 
         Caption         =   "Informações Sistema"
         Height          =   1545
         Left            =   3960
         TabIndex        =   46
         Top             =   4110
         Width           =   6825
         Begin VB.TextBox txtUsuarioCadastro 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   405
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   420
            Width           =   4785
         End
         Begin VB.TextBox txtUsuarioAlteracao 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   405
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   1080
            Width           =   4785
         End
         Begin SSCalendarWidgets_A.SSDateCombo dtcDataUltimaAlteracao 
            Height          =   405
            Left            =   4920
            TabIndex        =   48
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
            TabIndex        =   52
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
            TabIndex        =   54
            Top             =   210
            Width           =   1965
         End
         Begin VB.Label Label1 
            Caption         =   "Data Cadastro"
            Height          =   285
            Left            =   4920
            TabIndex        =   53
            Top             =   210
            Width           =   1575
         End
         Begin VB.Label Label26 
            Caption         =   "Usuário Ultima Alteração"
            Height          =   285
            Left            =   60
            TabIndex        =   50
            Top             =   870
            Width           =   1965
         End
         Begin VB.Label Label27 
            Caption         =   "Data Ultima alteração"
            Height          =   285
            Left            =   4920
            TabIndex        =   49
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
         TabIndex        =   37
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
            TabIndex        =   10
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
            TabIndex        =   11
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
            Width           =   6675
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
            TabIndex        =   8
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
            TabIndex        =   9
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
            TabIndex        =   7
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
         Begin VB.Label Label8 
            Caption         =   "Equipe"
            Height          =   285
            Left            =   60
            TabIndex        =   57
            Top             =   1680
            Width           =   525
         End
         Begin VB.Label Label6 
            Caption         =   "Número da camisa"
            Height          =   285
            Left            =   4140
            TabIndex        =   56
            Top             =   390
            Width           =   1365
         End
         Begin VB.Label Label4 
            Caption         =   "Nome do Pai"
            Height          =   285
            Left            =   60
            TabIndex        =   45
            Top             =   2340
            Width           =   1665
         End
         Begin VB.Label Label5 
            Caption         =   "Nome da Mãe"
            Height          =   285
            Left            =   60
            TabIndex        =   44
            Top             =   2940
            Width           =   1725
         End
         Begin VB.Label Label30 
            Caption         =   "Sexo"
            Height          =   285
            Left            =   5640
            TabIndex        =   43
            Top             =   390
            Width           =   885
         End
         Begin VB.Label Label7 
            Caption         =   "Data de Nascimento"
            Height          =   285
            Left            =   4920
            TabIndex        =   42
            Top             =   1680
            Width           =   1545
         End
         Begin VB.Label Label29 
            Caption         =   "Cartegoria"
            Height          =   285
            Left            =   2700
            TabIndex        =   41
            Top             =   1680
            Width           =   1365
         End
         Begin VB.Label Label 
            Caption         =   "Código"
            Height          =   285
            Left            =   120
            TabIndex        =   40
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Apelido 
            Caption         =   "Apelido do Jogador"
            Height          =   285
            Left            =   1080
            TabIndex        =   39
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label2 
            Caption         =   "Nome Completo"
            Height          =   285
            Left            =   90
            TabIndex        =   38
            Top             =   1050
            Width           =   1905
         End
      End
      Begin VB.Frame fraFoto 
         Height          =   5325
         Left            =   30
         TabIndex        =   33
         Top             =   330
         Width           =   3915
         Begin Threed.SSCommand cmdRemover 
            Height          =   330
            Left            =   1950
            TabIndex        =   34
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
            TabIndex        =   35
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
            TabIndex        =   36
            Top             =   180
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   7594
            _Version        =   196609
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
      Left            =   1395
      TabIndex        =   55
      Top             =   5775
      Width           =   9360
      _ExtentX        =   16510
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
            Enabled         =   0   'False
            Caption         =   "F6 - Abandonar"
            Key             =   "cmdLimpar"
            Object.ToolTipText     =   "Limpar dados da tela"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "F7-Gravar"
            Key             =   "cmdGravar"
            Object.ToolTipText     =   "Gravar Alterações"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "F8 - Imprimir"
            Key             =   "cmdimprimir"
            Object.ToolTipText     =   "Impirimir carteirinha ou Ficha do jogador"
            ImageIndex      =   8
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

Dim mbitFoto() As Byte

Public Property Let DiretorioFotoJogador(strDiretorio As String)
    mstrFoto = strDiretorio
End Property
Public Property Let OpcaoImpressao(lngOpcao As Long)
    mlngOpcao = lngOpcao
End Property

Private Sub cmdAdicionar_Click()
    On Error GoTo Erro
    
    frmAdicionarFotoJogador.Show vbModal
  
    If mstrFoto <> "" Then
        imgFotoJogador.Picture = Nothing
        imgFotoJogador.Stretch = True
        imgFotoJogador.Picture = LoadPicture(mstrFoto)
    End If
    
    Exit Sub
Erro:
 Call MsgBox("Erro no módulo: " & "frmCadastroDeJogador" & vbCrLf & "No Procedimento: " & "cmdAdicionar_Click" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Private Sub ImprimirFicha()
On Error GoTo Erro
      
    If (Val(txtCodigoInterno.Text)) = 0 Then Exit Sub
    
   rptFichaJogador.Parametros1 = mbitFoto()
   rptFichaJogador.Parametros111 = (Val(txtCodigoInterno.Text))
   rptFichaJogador.Parametros2 = (sscClube.Columns("chdescricao").Value)
   rptFichaJogador.Parametros3 = (txtNomeJogador.Text)
   rptFichaJogador.Parametros4 = (dtcDataNascimento.DateValue)
   rptFichaJogador.Parametros5 = ""
   rptFichaJogador.Parametros6 = txtCertidaoNascimento.Text
   rptFichaJogador.Parametros7 = txtCartorio.Text
   rptFichaJogador.Parametros8 = txtIdentidade.Text
   rptFichaJogador.Parametros9 = txtOrgao.Text
   rptFichaJogador.Parametros10 = txtNomePai.Text
   rptFichaJogador.Parametros11 = txtNomeMae.Text
   rptFichaJogador.Parametros12 = txtEnderecoAtleta.Text
   rptFichaJogador.Parametros13 = txtBairroEnderecoAtleta.Text
   rptFichaJogador.Parametros14 = txtCidadeEnderecoAtleta.Text
   rptFichaJogador.Parametros15 = txtFacebookAtleta.Text
   rptFichaJogador.Parametros16 = txtNomeEscola.Text
   rptFichaJogador.Parametros17 = txtEnderecoEscola.Text
   rptFichaJogador.Parametros18 = txtBairroEscola.Text
   rptFichaJogador.Parametros19 = txtCidadeEscola.Text
   rptFichaJogador.Parametros20 = ""
   rptFichaJogador.Parametros21 = txtRedeSocialEscola.Text
   
    rptFichaJogador.Show vbModal, Me

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmCadastroDeJogadorV2" & vbCrLf & "No Procedimento: " & "ImprimirFicha" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Private Sub ImprimirCarteirinha()
On Error GoTo Erro
    
    If (Val(txtCodigoInterno.Text)) = 0 Then Exit Sub
    
    rptCarteirinha.Codigo = (Val(txtCodigoInterno.Text))
    rptCarteirinha.Apelido = (txtApelido.Text)
    rptCarteirinha.Nome = (txtNomeJogador.Text)
    rptCarteirinha.Camisa = (Val(txtCodigoInterno.Text))
    rptCarteirinha.Equipe = (sscClube.Columns("chdescricao").Value)
    rptCarteirinha.Cartegoria = (sscCartegoria.Columns("chdescricao").Value)
    rptCarteirinha.Nascimento = (dtcDataNascimento.DateValue)
    rptCarteirinha.Mae = (txtNomePai.Text)
    rptCarteirinha.Pai = (txtNomeMae.Text)
    rptCarteirinha.Foto = mbitFoto()
    rptCarteirinha.Show vbModal, Me
Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmCadastroDeJogadorV2" & vbCrLf & "No Procedimento: " & "ImprimirCarteirinha" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Private Sub Command1_Click()
    rptFichaJogador.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2:  tbBotoes.Buttons("cmdNovo").Value = tbrPressed
        Case vbKeyF3:  tbBotoes.Buttons("cmdAlterar").Value = tbrPressed
        Case vbKeyF4:  tbBotoes.Buttons("cmdProcurar").Value = tbrPressed
        Case vbKeyF6:  tbBotoes.Buttons("cmdLimpar").Value = tbrPressed
        Case vbKeyF7:  tbBotoes.Buttons("cmdGravar").Value = tbrPressed
        Case vbKeyF8:  tbBotoes.Buttons("cmdimprimir").Value = tbrPressed
        Case vbKeyF10: tbBotoes.Buttons("cmdSair").Value = tbrPressed
   End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Erro

    tbBotoes.Buttons("cmdNovo").Value = tbrUnpressed
    tbBotoes.Buttons("cmdAlterar").Value = tbrUnpressed
    tbBotoes.Buttons("cmdProcurar").Value = tbrUnpressed
    tbBotoes.Buttons("cmdLimpar").Value = tbrUnpressed
    tbBotoes.Buttons("cmdGravar").Value = tbrUnpressed
    tbBotoes.Buttons("cmdimprimir").Value = tbrUnpressed
    tbBotoes.Buttons("cmdSair").Value = tbrUnpressed
  
    Select Case KeyCode
        Case vbKeyF2:  If tbBotoes.Buttons("cmdNovo").Enabled Then Call tbBotoes_ButtonClick(tbBotoes.Buttons("cmdNovo"))
        Case vbKeyF3:  If tbBotoes.Buttons("cmdAlterar").Enabled Then Call tbBotoes_ButtonClick(tbBotoes.Buttons("cmdAlterar"))
        Case vbKeyF4:  If tbBotoes.Buttons("cmdProcurar").Enabled Then Call tbBotoes_ButtonClick(tbBotoes.Buttons("cmdProcurar"))
        Case vbKeyF6:  If tbBotoes.Buttons("cmdLimpar").Enabled Then Call tbBotoes_ButtonClick(tbBotoes.Buttons("cmdLimpar"))
        Case vbKeyF7:  If tbBotoes.Buttons("cmdGravar").Enabled Then Call tbBotoes_ButtonClick(tbBotoes.Buttons("cmdGravar"))
        Case vbKeyF8:  If tbBotoes.Buttons("cmdimprimir").Enabled Then Call tbBotoes_ButtonClick(tbBotoes.Buttons("cmdimprimir"))
        Case vbKeyF10: If tbBotoes.Buttons("cmdSair").Enabled Then Call tbBotoes_ButtonClick(tbBotoes.Buttons("cmdSair"))
        
        Case vbKeyRight
            On Error Resume Next
            tabPrincipal.ActiveTab = tabPrincipal.ActiveTab + 1
            On Error GoTo Erro
        Case vbKeyLeft
             On Error Resume Next
            tabPrincipal.ActiveTab = tabPrincipal.ActiveTab - 1
            On Error GoTo Erro
    End Select
    
Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmCadastroDeJogadorV2" & vbCrLf & "No Procedimento: " & "Form_KeyUp" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub
Private Sub cmdRemover_Click()
    imgFotoJogador.Picture = Nothing
    mstrFoto = ""
End Sub

Private Sub Form_Load()
    
    mstrFlag = ""
    
    
    Call LimparCampos
    Call HabilitarCampos(False)
    
End Sub

Private Sub LimparCampos()
    
    LimparArray mbitFoto
    
    mstrFlag = ""
    
    txtCodigoInterno.Text = ""
    txtApelido.Text = ""
    txtNomeJogador.Text = ""
    sscClube.Text = ""
    'txtLocalDeNascimento.Text = ""
    txtNomePai.Text = ""
    txtNomeMae.Text = ""
    txtUsuarioAlteracao.Text = ""
    txtCertidaoNascimento.Text = ""
    txtCartorio.Text = ""
    txtIdentidade.Text = ""
    txtOrgao.Text = ""
    txtNomeEscola.Text = ""
    txtEnderecoEscola.Text = ""
    txtCidadeEscola.Text = ""
    txtBairroEscola.Text = ""
    txtRedeSocialEscola.Text = ""
    txtEnderecoAtleta.Text = ""
    txtCidadeEnderecoAtleta.Text = ""
    txtBairroEnderecoAtleta.Text = ""
    txtTelCel1.Text = ""
    txtTelCel2.Text = ""
    txtEmail.Text = ""
    txtFacebookAtleta.Text = ""
    txtInstagramAtleta.Text = ""
    txtNumCamisa.Text = ""
    
    optMasculino.Value = vbChecked
    optFeminino.Value = vbUnchecked
    
    chkWpp1.Value = vbUnchecked
    chkwpp2.Value = vbUnchecked
    
    dtcDataUltimaAlteracao.DateValue = Nothing
    dtcDataNascimento.DateValue = Nothing
    
    sscUfEscola.Text = ""
    sscUfEnderecoAtleta.Text = ""
    mstrFoto = ""
    imgFotoJogador.Picture = Nothing
    
    modBDCombo_SelecionarCartecoriaJogador sscCartegoria
    modBDCombo_SelecionarEstados sscUfEscola
    modBDCombo_SelecionarEstados sscUfEnderecoAtleta
    modBDCombo_SelecionarEquipePorCodigo sscClube
    
End Sub

Private Sub HabilitarCampos(blnhabilitar As Boolean)
    
    txtCodigoInterno.Locked = blnhabilitar
    txtApelido.Locked = Not blnhabilitar
    txtNomeJogador.Locked = Not blnhabilitar
    
    'txtLocalDeNascimento.Locked = Not blnhabilitar
    txtNomePai.Locked = Not blnhabilitar
    txtNomeMae.Locked = Not blnhabilitar
    txtUsuarioAlteracao.Locked = True
    txtCertidaoNascimento.Enabled = blnhabilitar
    txtCartorio.Locked = Not blnhabilitar
    txtIdentidade.Enabled = blnhabilitar
    txtOrgao.Locked = Not blnhabilitar
    txtNomeEscola.Locked = Not blnhabilitar
    txtEnderecoEscola.Locked = Not blnhabilitar
    txtCidadeEscola.Locked = Not blnhabilitar
    txtBairroEscola.Locked = Not blnhabilitar
    txtRedeSocialEscola.Locked = Not blnhabilitar
    txtEnderecoAtleta.Locked = Not blnhabilitar
    txtCidadeEnderecoAtleta.Locked = Not blnhabilitar
    txtBairroEnderecoAtleta.Locked = Not blnhabilitar
    txtTelCel1.Locked = Not blnhabilitar
    txtTelCel2.Locked = Not blnhabilitar
    txtEmail.Locked = Not blnhabilitar
    txtFacebookAtleta.Locked = Not blnhabilitar
    txtInstagramAtleta.Locked = Not blnhabilitar
    txtNumCamisa.Locked = Not blnhabilitar
    
    chkWpp1.Enabled = blnhabilitar
    chkwpp2.Enabled = blnhabilitar
    
    optMasculino.Enabled = blnhabilitar
    optFeminino.Enabled = blnhabilitar
    
    dtcDataNascimento.Enabled = blnhabilitar
    
    sscClube.Enabled = blnhabilitar
    sscCartegoria.Enabled = blnhabilitar
    sscUfEscola.Enabled = blnhabilitar
    sscUfEnderecoAtleta.Enabled = blnhabilitar
    
    cmdAdicionar.Enabled = blnhabilitar
    cmdRemover.Enabled = blnhabilitar
End Sub

Private Sub HabilitarTBBotoes(blnNovo As Boolean, blnAlterar As Boolean, blnProcurar As Boolean, blnAbandonar As Boolean, blnGravar As Boolean, blnSair As Boolean, blnImprimir As Boolean)

    tbBotoes.Buttons("cmdNovo").Enabled = blnNovo
    tbBotoes.Buttons("cmdAlterar").Enabled = blnAlterar
    tbBotoes.Buttons("cmdProcurar").Enabled = blnProcurar
    tbBotoes.Buttons("cmdLimpar").Enabled = blnAbandonar
    tbBotoes.Buttons("cmdGravar").Enabled = blnGravar
    tbBotoes.Buttons("cmdSair").Enabled = blnSair
    tbBotoes.Buttons("cmdimprimir").Enabled = blnImprimir
    
End Sub


Private Sub fraInfoBasicas_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub tbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Not (Button.Enabled) Then Exit Sub
    Select Case Button.Key
        
        Case "cmdNovo":
            mstrFlag = "I"
            Call HabilitarCampos(True)
            Call HabilitarTBBotoes(False, False, False, True, True, False, False)
            txtApelido.SetFocus
        
        Case "cmdAlterar":
            mstrFlag = "A"
            Call HabilitarCampos(True)
            Call HabilitarTBBotoes(False, False, False, True, True, False, False)
            txtApelido.SetFocus
        
        Case "cmdLimpar":
            mstrFlag = ""
            Call LimparCampos
            Call HabilitarCampos(False)
            Call HabilitarTBBotoes(True, False, True, False, False, True, True)
             txtCodigoInterno.SetFocus
        
        Case "cmdGravar"
            If VerificarCampos Then
                GravarJogador
                CarregarJogador Val(txtCodigoInterno.Text)
                mstrFlag = ""
            Else: Exit Sub
            End If
            Call HabilitarCampos(False)
            Call HabilitarTBBotoes(False, True, True, True, False, False, True)
            txtCodigoInterno.SetFocus
            
        Case "cmdProcurar"
            Dim ObjRelatorioJogador As ClsRelJogador
            Set ObjRelatorioJogador = New ClsRelJogador
            
            If Not gSMConexao Is Nothing Then
                If gSMConexao.EstadoConexaoBD = adStateOpen Then
                    
                    ObjRelatorioJogador.Show gSMConexao, "ProFut - Relatório de Jogador", vbModal, Me, True
                    txtCodigoInterno.Text = ObjRelatorioJogador.ID
                    txtCodigoInterno_KeyDown vbKeyReturn, 0
                    Exit Sub
                Else
                    gSMConexao.conectar
                End If
            End If
        
        Case "cmdimprimir"
            frmOpcaoImpressao.Show vbModal, Me
            Select Case mlngOpcao
                Case 1
                    ImprimirFicha
                Case 2
                    ImprimirCarteirinha
            End Select
        Case "cmdSair"
            Unload Me
        
    End Select
End Sub

Private Sub txtCertidaoNascimento_KeyPress(KeyAscii As Integer)
    TextBoxSomenteNumeros txtCertidaoNascimento.Text, KeyAscii, False, False
End Sub

Private Sub txtCodigoInterno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        CarregarJogador Val(txtCodigoInterno.Text)
    End If
End Sub

Private Sub txtCodigoInterno_KeyPress(KeyAscii As Integer)
    TextBoxSomenteNumeros txtCodigoInterno.Text, KeyAscii, False, False
End Sub

Private Sub txtEmail_LostFocus()
    EmailValido txtEmail.Text
End Sub

Private Sub txtApelido_Change()
    If KeyCode = vbKeyReturn Then
        'chamo procurar jogador aqui
    End If
End Sub

Private Sub txtApelido_KeyPress(KeyAscii As Integer)
    'TextBoxSomenteNumeros txtApelido.Text, KeyAscii, False, False
End Sub

Private Sub txtIdentidade_KeyPress(KeyAscii As Integer)
    TextBoxSomenteNumeros txtIdentidade.Text, KeyAscii, False, False
End Sub

Private Sub txtTelCel1_KeyPress(KeyAscii As Integer)
    TextBoxSomenteNumeros txtTelCel1.Text, KeyAscii, False, False
End Sub

Private Sub txtTelCel1_Validate(Cancel As Boolean)
    FormataTelefone txtTelCel1.Text, True
End Sub

Private Sub txtTelCel2_KeyPress(KeyAscii As Integer)
    TextBoxSomenteNumeros txtTelCel2.Text, KeyAscii, False, False
End Sub


Private Function VerificarCampos()
On Error GoTo Erro
    Dim blnContinua As Boolean
    Dim strMensagem As String
    
    blnContinua = True
'------------INFORMAÇÕES BASICAS------------
    If txtApelido.Text = "" Then
        strMensagem = strMensagem & "-> Número de inscrição não preenchido." & vbCrLf
        blnContinua = False
    End If
    
    If txtNomeJogador.Text = "" Then
        strMensagem = strMensagem & "-> Nome do jogador não preenchido." & vbCrLf
        blnContinua = False
    End If
    
    If sscClube.Text = "" Then
        strMensagem = strMensagem & "-> Equipe não informada ou inválida." & vbCrLf
        blnContinua = False
    End If
    
    If txtNomePai.Text = "" Then
        strMensagem = strMensagem & "-> Nome do pai não informado." & vbCrLf
        blnContinua = False
    End If
    
    If txtNomeMae.Text = "" Then
        strMensagem = strMensagem & "-> Nome da mãe não informado." & vbCrLf
        blnContinua = False
    End If
    
    If txtNomeMae.Text = "" Then
        strMensagem = strMensagem & "-> Nome da mãe não informado." & vbCrLf
        blnContinua = False
    End If
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
    If txtNomeEscola <> "" Then
    
        If txtEnderecoEscola.Text = "" Then
            strMensagem = strMensagem & "-> Endereço da escola não informado." & vbCrLf
            blnContinua = False
        End If
            
        If sscUfEnderecoAtleta.Text = "" Then
            strMensagem = strMensagem & "-> Estado da escola não informado." & vbCrLf
            blnContinua = False
        End If
                    
        If txtCidadeEscola.Text = "" Then
            strMensagem = strMensagem & "-> Cidade da escola não informado." & vbCrLf
            blnContinua = False
        End If
        
        If txtBairroEscola.Text = "" Then
            strMensagem = strMensagem & "-> Bairro da escola não informado." & vbCrLf
            blnContinua = False
        End If
    End If
    
'-------------Endereço/Contato do Atleta------------------
    If txtEnderecoAtleta.Text = "" Then
        strMensagem = strMensagem & "-> Endereço do atleta não informado." & vbCrLf
        blnContinua = False
    End If
    
    If sscUfEnderecoAtleta.Text = "" Then
        strMensagem = strMensagem & "-> Estado do atleta não informado." & vbCrLf
        blnContinua = False
    End If
    
    If txtCidadeEnderecoAtleta.Text = "" Then
        strMensagem = strMensagem & "-> Cidade do atleta não informado." & vbCrLf
        blnContinua = False
    End If
    
    If txtBairroEnderecoAtleta.Text = "" Then
        strMensagem = strMensagem & "-> Bairro do atleta não informado." & vbCrLf
        blnContinua = False
    End If
    
    If txtTelCel1.Text = "" And txtTelCel2.Text = "" Then
        strMensagem = strMensagem & "-> O atleta precisa ter pelo menos um telefone de contato." & vbCrLf
        blnContinua = False
    End If
    
    
    If Not blnContinua Then
        
        MsgBox "O jogador não pode ser gravado pois possuí as seguintes pendências: " & vbCrLf & strMensagem, vbOKOnly + vbInformation, "Atenção!"
        
    End If
    
    VerificarCampos = blnContinua

Exit Function
Erro:
   Call MsgBox("Erro no módulo: " & "frmCadastroDeJogador" & vbCrLf & "No Procedimento: " & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Function

Public Sub SalvaImagem(ByRef f() As Byte, File As String)
    Dim b() As Byte
    Dim ff  As Long
    Dim n   As Long
    
    On Error GoTo ErrHandler
    ff = FreeFile
    Open File For Binary Access Read As ff
    n = LOF(ff)
    If n Then
       ReDim b(1 To n) As Byte
       Get ff, , b()
    End If
    Close ff
    f() = b()
    Exit Sub
    
ErrHandler:
    MsgBox "ERROR: " & Err.Description
End Sub

Private Sub GravarJogador()
On Error GoTo Erro
      
    Dim udtJogador As TypJogador
    Dim binIMG() As Byte
    
    'COLOCO A IMAGEM EM CÓDIGO BINÁRIO
    If mstrFoto <> "" Then
        SalvaImagem binIMG(), mstrFoto
    End If
    
    With udtJogador
    
        .lngCodigo = IIf(txtCodigoInterno.Text <> "", Val(txtCodigoInterno.Text), 0)
        .strApelido = txtApelido.Text
        .strNomeAtleta = txtNomeJogador.Text
        .lngCartegoria = sscCartegoria.Columns("chcodigo").Value
        .lngEquipe = sscClube.Columns(1).Value
        '.strLocalNascimento = txtLocalDeNascimento.Text
        .strNomePai = txtNomePai.Text
        .strNomeMae = txtNomeMae.Text
        .strCertidaoNascimento = txtCertidaoNascimento.Text
        .strCartorio = txtCartorio.Text
        .strIdentidade = txtIdentidade.Text
        .strOrgaoIdentidade = txtOrgao.Text
        .datDataNascimento = dtcDataNascimento.DateValue
        If txtNomeEscola.Text <> "" Then
            .strEscola = txtNomeEscola.Text
            .lngEstadoEscola = sscUfEscola.Columns(1).Value
            .strCidadeEscola = txtCidadeEscola.Text
            .strBairroEscola = txtBairroEscola.Text
            .strFacebookEscola = txtRedeSocialEscola.Text
        End If
        .lngEstado = sscUfEnderecoAtleta.Columns("chcodigo").Value
        .strEndereco = txtEnderecoAtleta.Text
        .strCidade = txtCidadeEnderecoAtleta.Text
        .strBairro = txtBairroEnderecoAtleta.Text
        .strTelefone1 = txtTelCel1.Text
        .strTelefone2 = txtTelCel2.Text
        .blnWpp1 = IIf(chkWpp1.Value = vbChecked, True, False)
        .blnWpp2 = IIf(chkwpp2.Value = vbChecked, True, False)
        .strEmailContato = txtEmail.Text
        .strFacebook = txtFacebookAtleta
        .strInstagram = txtInstagramAtleta.Text
        .strEnderecoImagem() = IIf(mstrFoto <> "", binIMG(), 0)
        .lngSexo = IIf(optMasculino.Value = True, 1, 2)
        .lngNumeroCamisa = Val(txtNumCamisa.Text)
    End With
    
    
    If mstrFlag = "I" Then
        Call modJogador_AdicionarJogador(udtJogador)
       
    ElseIf mstrFlag = "A" Then
        Call modJogador_AlterarJogador(udtJogador)
    End If
    
    
    txtCodigoInterno.Text = udtJogador.lngCodigo
    


Exit Sub
Erro:
    Call MsgBox("Erro no módulo: " & "frmCadastroDeJogador" & vbCrLf & "No Procedimento: " & "GravarJogador" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Private Sub CarregarJogador(lngCodigo As Long)
On Error GoTo Erro
    Dim objRsJogador As Recordset
    Dim binIMG() As Byte
    Set objRsJogador = New Recordset
      
    Call LimparCampos
    modJogador_SelecionarJogadorPorCodigo lngCodigo, objRsJogador
    
    If Not objRsJogador Is Nothing Then
        If Not objRsJogador.EOF And Not objRsJogador.BOF Then
            txtCodigoInterno.Text = NZ(objRsJogador!ID_JOGADOR_IN)
            txtApelido.Text = NS(objRsJogador!APELIDO_VC)
            txtNomeJogador.Text = NS(objRsJogador!NOMEATLETA_VC)
            txtNumCamisa.Text = NZ(objRsJogador!NUMEROCAMISA_IN)
            txtNomePai.Text = NS(objRsJogador!NOMEPAI_VC)
            txtNomeMae.Text = NS(objRsJogador!STRNOMEMAE_VC)
            txtUsuarioAlteracao.Text = NS(objRsJogador!USUARIOULTIMAALTERACAO_VC)
            txtUsuarioCadastro.Text = NS(objRsJogador!USUARIOCADASTRO_VC)
            txtCertidaoNascimento.Text = NS(objRsJogador!CERTIDAONASCIMENTO_VC)
            txtCartorio.Text = NS(objRsJogador!CARTORIO_VC)
            txtIdentidade.Text = NS(objRsJogador!IDENTIDADE_VC)
            txtOrgao.Text = NS(objRsJogador!ORGAOIDENTIDADE_VC)
            txtNomeEscola.Text = NS(objRsJogador!ESCOLA_VC)
            txtEnderecoEscola.Text = NS(objRsJogador!ENDERECOESCOLA_VC)
            txtCidadeEscola.Text = NS(objRsJogador!CIDADEESCOLA_VC)
            txtBairroEscola.Text = NS(objRsJogador!BAIRROESCOLA_VC)
            txtRedeSocialEscola.Text = NS(objRsJogador!REDESOCIALESCOLA_VC)
            txtEnderecoAtleta.Text = NS(objRsJogador!ENDERECO_VC)
            txtCidadeEnderecoAtleta.Text = NS(objRsJogador!CIDADE_VC)
            txtBairroEnderecoAtleta.Text = NS(objRsJogador!BAIRRO_VC)
            txtTelCel1.Text = NS(objRsJogador!TELCEL1_VC)
            txtTelCel2.Text = NS(objRsJogador!TELCEL2_VC)
            txtEmail.Text = NS(objRsJogador!EMAIL_VC)
            txtFacebookAtleta.Text = NS(objRsJogador!FACEBOOK_VC)
            txtInstagramAtleta.Text = NS(objRsJogador!INSTAGRAM_VC)
            
            If NZ(objRsJogador!SEXO_IN) = 1 Then
                optMasculino.Value = True
                optFeminino.Value = False
            Else
                optMasculino.Value = False
                optFeminino.Value = True
            End If
            
            chkWpp1.Value = IIf(NB(objRsJogador!WPP1_BT), vbChecked, vbUnchecked)
            chkwpp2.Value = IIf(NB(objRsJogador!WPP2_BT), vbChecked, vbUnchecked)
            
            dtcDataCadastro.DateValue = ND(objRsJogador!DATACADASTRO_DT)
            dtcDataUltimaAlteracao.DateValue = ND(objRsJogador!DATAULTIMAALTERACAO_DT)
            dtcDataNascimento.DateValue = ND(objRsJogador!DATANASCIMENTO_DT)
            
            
'            mstrFoto = NS(objRsJogador!ENDERECOIMAGEM_VC)
'            If mstrFoto <> "" Then
'                imgFotoJogador.Picture = Nothing
'                imgFotoJogador.Stretch = True
'                On Error Resume Next
'                imgFotoJogador.Picture = LoadPicture(mstrFoto)
'                On Error GoTo Erro
'            End If

'--------------------------------------------------------------------------------------
            If Not IsNull(objRsJogador!ENDERECOIMAGEM_VC) Then
                
                mbitFoto() = objRsJogador!ENDERECOIMAGEM_VC
                binIMG() = objRsJogador!ENDERECOIMAGEM_VC
                If Val(binIMG(1)) <> 0 Then
                    imgFotoJogador.Picture = Nothing
                    imgFotoJogador.Stretch = True
                    On Error Resume Next

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
                    imgFotoJogador.Picture = LoadPicture(Arquivo)
                    'Set GetImageFromField = LoadPicture(Arquivo)
                    Kill Arquivo
                    End If
                End If
'--------------------------------------------------------------------------------------
            
            modBDCombo_SelecionarCartecoriaJogador sscCartegoria, NZ(objRsJogador!CARTEGORIA_IN)
            modBDCombo_SelecionarEquipePorCodigo sscClube, NZ(objRsJogador!EQUIPE_IN)
            modBDCombo_SelecionarEstados sscUfEscola, NZ(objRsJogador!ESTADOESCOLA_IN)
            modBDCombo_SelecionarEstados sscUfEnderecoAtleta, NZ(objRsJogador!ESTADO_IN)
            
            mstrFlag = ""
            Call HabilitarCampos(False)
            Call HabilitarTBBotoes(False, True, True, True, False, True, True)
            
        Else
            MsgBox "Jogador não encontrado ou código inválido.", vbOKOnly + vbInformation, "Atenção!"
        End If
    Else
        MsgBox "Jogador não encontrado ou código inválido.", vbOKOnly + vbInformation, "Atenção!"
    End If

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmCadastroDeJogador" & vbCrLf & "No Procedimento: " & "CarregarJogador" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

End Sub

Private Sub txtTelCel2_Validate(Cancel As Boolean)
    FormataTelefone txtTelCel2.Text, True
End Sub

Private Sub LimparArray(arr As Variant)
    Dim i As Integer
    
    Do While i < Len(arr)
    
        arr(i) = Empty
        
        i = i + 1
    Loop
End Sub
