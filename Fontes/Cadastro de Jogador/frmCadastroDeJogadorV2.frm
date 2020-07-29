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
            TabIndex        =   23
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
            TabIndex        =   24
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
            TabIndex        =   84
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
         Begin VB.TextBox txtTelefoneEscola 
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
            Left            =   8520
            MaxLength       =   11
            TabIndex        =   18
            Top             =   720
            Width           =   2145
         End
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
            Width           =   8415
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
            TabIndex        =   19
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
            TabIndex        =   21
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
            TabIndex        =   22
            Top             =   2610
            Width           =   10575
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscUfEscola 
            Height          =   390
            Left            =   90
            TabIndex        =   20
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
            TabIndex        =   83
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
         Begin VB.Label Label33 
            Caption         =   "Telefone/Celular"
            Height          =   285
            Left            =   8520
            TabIndex        =   86
            Top             =   480
            Width           =   1815
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
            TabIndex        =   85
            Top             =   1050
            Width           =   645
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
               TabIndex        =   82
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
      TabIndex        =   55
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
   rptFichaJogador.Parametros14 = sscCidadeEnderecoAtleta.Text
   rptFichaJogador.Parametros15 = txtFacebookAtleta.Text
   rptFichaJogador.Parametros16 = txtNomeEscola.Text
   rptFichaJogador.Parametros17 = txtEnderecoEscola.Text
   rptFichaJogador.Parametros18 = txtBairroEscola.Text
   rptFichaJogador.Parametros19 = sscCidadeEscola.Text
   rptFichaJogador.Parametros20 = txtTelefoneEscola.Text
   rptFichaJogador.Parametros21 = txtRedeSocialEscola.Text
   
    rptFichaJogador.Show vbModal, Me

Exit Sub
Erro:
   Call MsgBox("Erro no módulo: " & "frmCadastroDeJogadorV2" & vbCrLf & "No Procedimento: " & "ImprimirFicha" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")

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
310       txtTelefoneEscola.Text = ""

          
320       optMasculino.Value = vbChecked
330       optFeminino.Value = vbUnchecked
          
340       chkWpp1.Value = vbUnchecked
350       chkwpp2.Value = vbUnchecked
          
360       dtcDataUltimaAlteracao.DateValue = Nothing
370       dtcDataNascimento.DateValue = Nothing
380       dtcDataCadastro.DateValue = Nothing
          
390       sscUfEscola.Text = ""
400       sscUfEnderecoAtleta.Text = ""
410       mstrFoto = ""
420       imgFotoJogador.Picture = Nothing
          
430       modBDCombo_SelecionarCartecoriaJogador sscCartegoria
440       modBDCombo_SelecionarEstados sscUfEscola
450       modBDCombo_SelecionarEstados sscUfEnderecoAtleta
460       modBDCombo_SelecionarEquipePorCodigo sscClube
470       modBDCombo_SelecionarPosicoesAtleta sscPosicao
          
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
250       txtTelefoneEscola.Locked = Not blnhabilitar
          
260       chkWpp1.Enabled = blnhabilitar
270       chkwpp2.Enabled = blnhabilitar
          
280       optMasculino.Enabled = blnhabilitar
290       optFeminino.Enabled = blnhabilitar
          
300       dtcDataNascimento.Enabled = blnhabilitar
          
310       sscClube.Enabled = blnhabilitar
320       sscPosicao.Enabled = blnhabilitar
330       sscCartegoria.Enabled = blnhabilitar
340       sscUfEscola.Enabled = blnhabilitar
350       sscUfEnderecoAtleta.Enabled = blnhabilitar
          
360       cmdAdicionar.Enabled = blnhabilitar
370       cmdRemover.Enabled = blnhabilitar
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

      Dim lngClube As Long

10        If Not (Button.Enabled) Then Exit Sub
20        Select Case Button.Key
              
              Case "cmdNovo":
30                If RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 2) = True Then
40                    mstrFlag = "I"
50                    Call HabilitarCampos(True)
60                    Call HabilitarTBBotoes(False, False, False, True, True, False, False, False)
70                    lngClube = NZ(RetornaClubePorUsuario(gSMConexao.CodigoUsuario))
80                    If lngClube <> 0 Then
90                        If RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 13) = False Then
100                           modBDCombo_SelecionarEquipePorCodigo sscClube, lngClube
110                           sscClube.Enabled = False
120                       End If
130                   End If
140                   txtApelido.SetFocus
150               Else
160                   MsgBox "Permissão requerida!" & vbCrLf & "-> Permissão Nº2" & vbCrLf & vbCrLf & "Entre em contato com o administrador para liberar a permissão!", vbOKOnly + vbExclamation, "Permissão negada!"
170               End If
              
180           Case "cmdAlterar":
                  
190               If RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 3) = True Then
200                   mstrFlag = "A"
210                   Call HabilitarCampos(True)
220                   Call HabilitarTBBotoes(False, False, False, True, True, False, False, False)
230                   lngClube = NZ(RetornaClubePorUsuario(gSMConexao.CodigoUsuario))
240                   If lngClube <> 0 Then
250                       If RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 13) = False Then
260                           modBDCombo_SelecionarEquipePorCodigo sscClube, lngClube
270                           sscClube.Enabled = False
280                       End If
290                   End If
300                   txtApelido.SetFocus
310               Else
320                   MsgBox "Permissão requerida!" & vbCrLf & "-> Permissão Nº3" & vbCrLf & vbCrLf & "Entre em contato com o administrador para liberar a permissão!", vbOKOnly + vbExclamation, "Permissão negada!"
330               End If
                  
340           Case "cmdExcluir":
                  
350               If lblInativo.Visible = False Then
360                   If MsgBox("Deseja INATIVAR o jogador no sistema?" & vbCrLf & vbCrLf & "O jogador será removido temporáriamente do sistema mas ainda poderá ser reativado.", vbYesNo + vbInformation, "Atenção!") = vbYes Then
370                       Call modJogador_ApagarJogadorPorCodigo(Val(txtCodigoInterno.Text), 1)
380                       Call CarregarJogador(Val(txtCodigoInterno.Text))
390                   End If
400               Else
410                   If MsgBox("Deseja EXCLUIR o jogador no sistema?" & vbCrLf & vbCrLf & "O jogador será removido permanentemente do sistema e NÃO poderá ser reativado.", vbYesNo + vbInformation, "Atenção!") = vbYes Then
420                       Call modJogador_ApagarJogadorPorCodigo(Val(txtCodigoInterno.Text), 2)
430                       mstrFlag = ""
440                       Call LimparCampos
450                       Call HabilitarCampos(False)
460                       Call HabilitarTBBotoes(True, False, True, False, False, True, True, False)
470                       txtCodigoInterno.SetFocus
480                   End If
490               End If
                  
              
500           Case "cmdLimpar":
510               mstrFlag = ""
520               Call LimparCampos
530               Call HabilitarCampos(False)
540               Call HabilitarTBBotoes(True, False, True, False, False, True, True, False)
550                txtCodigoInterno.SetFocus
              
560           Case "cmdGravar"
570               If VerificarCampos Then
580                   GravarJogador
590                   CarregarJogador Val(txtCodigoInterno.Text)
600                   mstrFlag = ""
610               Else
620                   Exit Sub
630               End If
640               Call HabilitarCampos(False)
650               Call HabilitarTBBotoes(False, True, True, True, False, False, True, True)
660               txtCodigoInterno.SetFocus
                  
670           Case "cmdProcurar"
680               If tbBotoes.Buttons("cmdProcurar").Caption = "F4 - Procurar" Then
                      Dim ObjRelatorioJogador As ClsRelJogador
690                   Set ObjRelatorioJogador = New ClsRelJogador
                      
700                   If Not gSMConexao Is Nothing Then
710                       If gSMConexao.EstadoConexaoBD = adStateOpen Then
                              
720                           ObjRelatorioJogador.Show gSMConexao, "ProFut - Relatório de Jogador", vbModal, Me, True
730                           txtCodigoInterno.Text = ObjRelatorioJogador.ID
740                           txtCodigoInterno_KeyDown vbKeyReturn, 0
750                           Exit Sub
760                       Else
770                           gSMConexao.conectar
780                       End If
790                   End If
800               Else
810                   If MsgBox("Deseja REATIVAR o jogador no sistema?" & vbCrLf & vbCrLf & "O jogador será ativado novamente no sistema.", vbYesNo + vbInformation, "Atenção!") = vbYes Then
820                       Call modJogador_ApagarJogadorPorCodigo(Val(txtCodigoInterno.Text), 3)
830                       Call CarregarJogador(Val(txtCodigoInterno.Text))
840                   End If
850               End If
              
860           Case "cmdimprimir"
              
870               If RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 4) = True Then
880                   frmOpcaoImpressao.Show vbModal, Me
890                   Select Case mlngOpcao
                          Case 1
900                           ImprimirFicha
910                       Case 2
920                           ImprimirCarteirinha
930                   End Select
940               Else
950                   MsgBox "Permissão requerida!" & vbCrLf & "-> Permissão Nº4" & vbCrLf & vbCrLf & "Entre em contato com o administrador para liberar a permissão!", vbOKOnly + vbExclamation, "Permissão negada!"
960               End If
970           Case "cmdSair"
980               Unload Me
              
990       End Select
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
              
480           If txtTelefoneEscola.Text = "" Then
490               strMensagem = strMensagem & "-> Telefone da escola não informado." & vbCrLf
500               blnContinua = False
510           End If
520       End If
          
      '-------------Endereço/Contato do Atleta------------------
530       If txtEnderecoAtleta.Text = "" Then
540           strMensagem = strMensagem & "-> Endereço do atleta não informado." & vbCrLf
550           blnContinua = False
560       End If
          
570       If sscUfEnderecoAtleta.Text = "" Then
580           strMensagem = strMensagem & "-> Estado do atleta não informado." & vbCrLf
590           blnContinua = False
600       End If
          
610       If sscCidadeEnderecoAtleta.Text = "" Then
620           strMensagem = strMensagem & "-> Cidade do atleta não informado." & vbCrLf
630           blnContinua = False
640       End If
          
650       If txtBairroEnderecoAtleta.Text = "" Then
660           strMensagem = strMensagem & "-> Bairro do atleta não informado." & vbCrLf
670           blnContinua = False
680       End If
          
690       If txtTelCel1.Text = "" And txtTelCel2.Text = "" Then
700           strMensagem = strMensagem & "-> O atleta precisa ter pelo menos um telefone de contato." & vbCrLf
710           blnContinua = False
720       End If
          
          
730       If Not blnContinua Then
              
740           MsgBox "O jogador não pode ser gravado pois possuí as seguintes pendências: " & vbCrLf & strMensagem, vbOKOnly + vbInformation, "Atenção!"
              
750       End If
          
760       VerificarCampos = blnContinua

770   Exit Function
Erro:
780      Call MsgBox("Erro no módulo: " & "frmCadastroDeJogador" & vbCrLf & "No Procedimento: " & "VerificarCampos" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
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
240               .strEnderecoEscola = txtEnderecoEscola.Text
250               .strFacebookEscola = txtRedeSocialEscola.Text
260               .strTelefoneEscola = txtTelefoneEscola.Text
270           End If
280           .lngEstado = sscUfEnderecoAtleta.Columns("chcodigo").Value
290           .strEndereco = txtEnderecoAtleta.Text
300           .strCidade = sscCidadeEnderecoAtleta.Text
310           .strBairro = txtBairroEnderecoAtleta.Text
320           .strTelefone1 = txtTelCel1.Text
330           .strTelefone2 = txtTelCel2.Text
340           .blnWpp1 = IIf(chkWpp1.Value = vbChecked, True, False)
350           .blnWpp2 = IIf(chkwpp2.Value = vbChecked, True, False)
360           .strEmailContato = txtEmail.Text
370           .strFacebook = txtFacebookAtleta
380           .strInstagram = txtInstagramAtleta.Text
390           .strEnderecoImagem() = IIf(mstrFoto <> "", binIMG(), mbitFoto())
400           .lngSexo = IIf(optMasculino.Value = True, 1, 2)
410           .lngNumeroCamisa = Val(txtNumCamisa.Text)
420           .blnTemImagem = IIf(blnRemoveuFoto = True, False, True)
430       End With
          
          
440       If mstrFlag = "I" Then
450           Call modJogador_AdicionarJogador(udtJogador)
             
460       ElseIf mstrFlag = "A" Then
470           Call modJogador_AlterarJogador(udtJogador)
480       End If
          
          
490       txtCodigoInterno.Text = udtJogador.lngCodigo
          


500   Exit Sub
Erro:
510       Call MsgBox("Erro no módulo: " & "frmCadastroDeJogador" & vbCrLf & "No Procedimento: " & "GravarJogador" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
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
        
        
            If RetornaAcessoPorUsuarioEPermissao(gSMConexao.CodigoUsuario, 12) = False Then
                If RetornaClubePorUsuario(gSMConexao.CodigoUsuario) <> NZ(objRsJogador!EQUIPE_IN) Then
                    MsgBox "Você não tem permissão para visualizar o jogador pois ele não pertence a sua equipe.", vbOKOnly + vbExclamation, "Atenção!"
                    Exit Sub
                End If
            End If
            
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
            sscCidadeEscola.Text = NS(objRsJogador!CIDADEESCOLA_VC)
            txtBairroEscola.Text = NS(objRsJogador!BAIRROESCOLA_VC)
            txtRedeSocialEscola.Text = NS(objRsJogador!REDESOCIALESCOLA_VC)
            txtEnderecoAtleta.Text = NS(objRsJogador!ENDERECO_VC)
            sscCidadeEnderecoAtleta.Text = NS(objRsJogador!CIDADE_VC)
            txtBairroEnderecoAtleta.Text = NS(objRsJogador!BAIRRO_VC)
            txtTelCel1.Text = NS(objRsJogador!TELCEL1_VC)
            txtTelCel2.Text = NS(objRsJogador!TELCEL2_VC)
            txtEmail.Text = NS(objRsJogador!EMAIL_VC)
            txtFacebookAtleta.Text = NS(objRsJogador!FACEBOOK_VC)
            txtInstagramAtleta.Text = NS(objRsJogador!INSTAGRAM_VC)
            txtTelefoneEscola.Text = NS(objRsJogador!TELEFONEESCOLA_VC)
            
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
            
            
            lblInativo.Visible = NB(objRsJogador!EXCLUIDO_BT)
            
            If NB(objRsJogador!EXCLUIDO_BT) = True Then
                tbBotoes.Buttons("cmdExcluir").Caption = "F5 - Excluir"
                tbBotoes.Buttons("cmdProcurar").Caption = "F4 - Reativar"
                tbBotoes.Buttons("cmdProcurar").Image = 2
            Else
                tbBotoes.Buttons("cmdExcluir").Caption = "F5 - Inativar"
                tbBotoes.Buttons("cmdProcurar").Caption = "F4 - Procurar"
                tbBotoes.Buttons("cmdProcurar").Image = 4
            End If
            
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
            modBDCombo_SelecionarPosicoesAtleta sscPosicao, NZ(objRsJogador!POSICAO_IN)
            modBDCombo_SelecionarEstados sscUfEscola, NZ(objRsJogador!ESTADOESCOLA_IN)
            modBDCombo_SelecionarEstados sscUfEnderecoAtleta, NZ(objRsJogador!Estado_IN)
            
            mstrFlag = ""
            Call HabilitarCampos(False)
            Call HabilitarTBBotoes(False, True, True, True, False, True, True, True)
            
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

Private Sub txtTelefoneEscola_KeyPress(KeyAscii As Integer)
    TextBoxSomenteNumeros txtTelefoneEscola.Text, KeyAscii, False, False
End Sub
