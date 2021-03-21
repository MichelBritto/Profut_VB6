VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmCadastroDeJogador 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProFut - Cadastro de Jogador"
   ClientHeight    =   8040
   ClientLeft      =   3990
   ClientTop       =   1305
   ClientWidth     =   11835
   Icon            =   "frmCadastroDeJogador.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin VB.Frame fraPrincipal 
      Height          =   9615
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   20415
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
         Left            =   120
         MaxLength       =   128
         TabIndex        =   75
         Top             =   4830
         Width           =   5355
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
         Height          =   3615
         Left            =   60
         TabIndex        =   49
         Top             =   5940
         Width           =   10155
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
            TabIndex        =   30
            Top             =   2280
            Width           =   10005
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
            TabIndex        =   28
            Top             =   1650
            Width           =   4335
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
            Left            =   5400
            MaxLength       =   128
            TabIndex        =   29
            Top             =   1650
            Width           =   4635
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
            TabIndex        =   26
            Top             =   990
            Width           =   10005
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
            TabIndex        =   25
            Top             =   390
            Width           =   10005
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscUfEscola 
            Height          =   390
            Left            =   90
            TabIndex        =   27
            Top             =   1650
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
         Begin VB.Label Label28 
            Caption         =   "UF"
            Height          =   285
            Left            =   120
            TabIndex        =   66
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label25 
            Caption         =   "Rede Social"
            Height          =   285
            Left            =   60
            TabIndex        =   61
            Top             =   2070
            Width           =   1785
         End
         Begin VB.Label Label24 
            Caption         =   "Cidade"
            Height          =   285
            Left            =   1050
            TabIndex        =   60
            Top             =   1410
            Width           =   2265
         End
         Begin VB.Label Label23 
            Caption         =   "Bairro"
            Height          =   285
            Left            =   5400
            TabIndex        =   59
            Top             =   1440
            Width           =   4695
         End
         Begin VB.Label Label22 
            Caption         =   "Endereço"
            Height          =   285
            Left            =   60
            TabIndex        =   58
            Top             =   780
            Width           =   5745
         End
         Begin VB.Label Label21 
            Caption         =   "Nome Instituição"
            Height          =   285
            Left            =   60
            TabIndex        =   57
            Top             =   180
            Width           =   1485
         End
      End
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
         Height          =   5145
         Left            =   10260
         TabIndex        =   44
         Top             =   4410
         Width           =   10095
         Begin VB.PictureBox wpp2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   7680
            Picture         =   "frmCadastroDeJogador.frx":038A
            ScaleHeight     =   345
            ScaleWidth      =   315
            TabIndex        =   56
            Top             =   1740
            Width           =   315
         End
         Begin VB.PictureBox wpp1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   2550
            Picture         =   "frmCadastroDeJogador.frx":0840
            ScaleHeight     =   345
            ScaleWidth      =   315
            TabIndex        =   55
            Top             =   1740
            Width           =   315
         End
         Begin VB.CheckBox chkwpp2 
            Height          =   195
            Left            =   7440
            TabIndex        =   21
            Top             =   1830
            Width           =   225
         End
         Begin VB.CheckBox chkWpp1 
            Height          =   195
            Left            =   2280
            Picture         =   "frmCadastroDeJogador.frx":0CF6
            TabIndex        =   19
            Top             =   1800
            Width           =   255
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
            Left            =   90
            MaxLength       =   128
            TabIndex        =   24
            Top             =   3600
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
            Left            =   90
            MaxLength       =   128
            TabIndex        =   23
            Top             =   2970
            Width           =   9915
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
            Left            =   90
            MaxLength       =   128
            TabIndex        =   22
            Top             =   2340
            Width           =   9915
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
            TabIndex        =   18
            Top             =   1710
            Width           =   2145
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
            Left            =   5160
            MaxLength       =   11
            TabIndex        =   20
            Top             =   1710
            Width           =   2235
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
            Left            =   90
            MaxLength       =   128
            TabIndex        =   14
            Top             =   420
            Width           =   8895
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
            Left            =   90
            MaxLength       =   128
            TabIndex        =   16
            Top             =   1050
            Width           =   5025
         End
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
            Left            =   5160
            MaxLength       =   128
            TabIndex        =   17
            Top             =   1050
            Width           =   4845
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscUfEnderecoAtleta 
            Height          =   390
            Left            =   9060
            TabIndex        =   15
            Top             =   420
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
         Begin VB.Label Label20 
            Caption         =   "Instagram"
            Height          =   285
            Left            =   90
            TabIndex        =   54
            Top             =   3390
            Width           =   1815
         End
         Begin VB.Label Label19 
            Caption         =   "Facebook"
            Height          =   285
            Left            =   90
            TabIndex        =   53
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "E-mail Contato"
            Height          =   285
            Left            =   90
            TabIndex        =   52
            Top             =   2130
            Width           =   1815
         End
         Begin VB.Label Label17 
            Caption         =   "Telefone/Celular"
            Height          =   285
            Left            =   90
            TabIndex        =   51
            Top             =   1470
            Width           =   1815
         End
         Begin VB.Label Label16 
            Caption         =   "Telefone/Celular 2"
            Height          =   285
            Left            =   5160
            TabIndex        =   50
            Top             =   1500
            Width           =   1815
         End
         Begin VB.Label Label14 
            Caption         =   "UF"
            Height          =   285
            Left            =   9090
            TabIndex        =   48
            Top             =   210
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "Endereço "
            Height          =   285
            Left            =   90
            TabIndex        =   47
            Top             =   210
            Width           =   1815
         End
         Begin VB.Label Label13 
            Caption         =   "Cidade"
            Height          =   285
            Left            =   90
            TabIndex        =   46
            Top             =   810
            Width           =   1815
         End
         Begin VB.Label Label12 
            Caption         =   "Bairro"
            Height          =   285
            Left            =   5160
            TabIndex        =   45
            Top             =   840
            Width           =   3885
         End
      End
      Begin VB.Frame fraInfoBasicas 
         Caption         =   "Informações Básicas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4305
         Left            =   3840
         TabIndex        =   31
         Top             =   120
         Width           =   16515
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
            Height          =   405
            Left            =   120
            MaxLength       =   128
            TabIndex        =   77
            Top             =   1800
            Width           =   6855
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
            Left            =   120
            MaxLength       =   128
            TabIndex        =   76
            Top             =   2400
            Width           =   6855
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
            Left            =   15840
            TabIndex        =   6
            Top             =   510
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
            Left            =   15330
            TabIndex        =   5
            Top             =   510
            Value           =   -1  'True
            Width           =   555
         End
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
            Left            =   120
            MaxLength       =   20
            TabIndex        =   7
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Frame fraInfoSistema 
            Caption         =   "Informações Sistema"
            Height          =   1215
            Left            =   12420
            TabIndex        =   62
            Top             =   3030
            Width           =   4035
            Begin VB.TextBox txtUsuarioAlteracao 
               Appearance      =   0  'Flat
               Height          =   405
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   63
               Top             =   510
               Width           =   1995
            End
            Begin SSCalendarWidgets_A.SSDateCombo dtcDataUltimaAlteracao 
               Height          =   405
               Left            =   2160
               TabIndex        =   71
               Top             =   510
               Width           =   1815
               _Version        =   65543
               _ExtentX        =   3201
               _ExtentY        =   714
               _StockProps     =   93
               Format          =   "DD/MM/YY"
               BevelType       =   0
            End
            Begin VB.Label Label27 
               Caption         =   "Data Ultima alteração"
               Height          =   285
               Left            =   2160
               TabIndex        =   65
               Top             =   300
               Width           =   1575
            End
            Begin VB.Label Label26 
               Caption         =   "Usuário Ultima Alteração"
               Height          =   285
               Left            =   120
               TabIndex        =   64
               Top             =   300
               Width           =   1965
            End
         End
         Begin VB.TextBox txtLocalDeNascimento 
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
            Left            =   5220
            MaxLength       =   128
            TabIndex        =   9
            Top             =   1080
            Width           =   9315
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
            Left            =   1110
            MaxLength       =   20
            TabIndex        =   2
            Top             =   420
            Width           =   4125
         End
         Begin SSCalendarWidgets_A.SSDateCombo dtcDataNascimento 
            Height          =   405
            Left            =   14610
            TabIndex        =   10
            Top             =   1080
            Width           =   1815
            _Version        =   65543
            _ExtentX        =   3201
            _ExtentY        =   714
            _StockProps     =   93
            Format          =   "DD/MM/YY"
            BevelType       =   0
         End
         Begin VB.TextBox txtCodigoInterno 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   120
            MaxLength       =   8
            TabIndex        =   1
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
            Left            =   5280
            MaxLength       =   128
            TabIndex        =   3
            Top             =   420
            Width           =   6855
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscCartegoria 
            Height          =   390
            Left            =   12180
            TabIndex        =   4
            Top             =   420
            Width           =   3045
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
            _ExtentX        =   5371
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
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscClube 
            Height          =   390
            Left            =   1530
            TabIndex        =   8
            Top             =   1080
            Width           =   3645
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
            _ExtentX        =   6429
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
         Begin VB.Label Label4 
            Caption         =   "Nome do Pai"
            Height          =   285
            Left            =   120
            TabIndex        =   79
            Top             =   1590
            Width           =   1785
         End
         Begin VB.Label Label5 
            Caption         =   "Nome da Mãe"
            Height          =   285
            Left            =   120
            TabIndex        =   78
            Top             =   2190
            Width           =   1845
         End
         Begin VB.Label Label30 
            Caption         =   "Sexo"
            Height          =   285
            Left            =   15330
            TabIndex        =   74
            Top             =   210
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Número da camisa"
            Height          =   285
            Left            =   120
            TabIndex        =   73
            Top             =   870
            Width           =   1365
         End
         Begin VB.Label Label29 
            Caption         =   "Cartegoria"
            Height          =   285
            Left            =   12180
            TabIndex        =   70
            Top             =   210
            Width           =   2745
         End
         Begin VB.Label Label7 
            Caption         =   "Data de Nascimento"
            Height          =   285
            Left            =   14610
            TabIndex        =   38
            Top             =   870
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "Local de Nascimento"
            Height          =   285
            Left            =   5220
            TabIndex        =   37
            Top             =   870
            Width           =   3855
         End
         Begin VB.Label Label3 
            Caption         =   "Equipe"
            Height          =   285
            Left            =   1530
            TabIndex        =   36
            Top             =   870
            Width           =   1485
         End
         Begin VB.Label Label2 
            Caption         =   "Nome Completo"
            Height          =   285
            Left            =   5280
            TabIndex        =   35
            Top             =   210
            Width           =   2085
         End
         Begin VB.Label Apelido 
            Caption         =   "Apelido do Jogador"
            Height          =   285
            Left            =   1110
            TabIndex        =   34
            Top             =   210
            Width           =   1485
         End
         Begin VB.Label Label 
            Caption         =   "Código"
            Height          =   285
            Left            =   150
            TabIndex        =   33
            Top             =   210
            Width           =   855
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
         Left            =   60
         TabIndex        =   32
         Top             =   4410
         Width           =   10155
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
            TabIndex        =   13
            Top             =   1050
            Width           =   4575
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
            TabIndex        =   11
            Top             =   420
            Width           =   4575
         End
         Begin EditLib.fpMask txtIdentidade 
            Height          =   405
            Left            =   60
            TabIndex        =   12
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
         Begin VB.Label Label11 
            Caption         =   "Órgão Expedidor"
            Height          =   285
            Left            =   5490
            TabIndex        =   43
            Top             =   840
            Width           =   3795
         End
         Begin VB.Label Label10 
            Caption         =   "Identidade(RG)"
            Height          =   285
            Left            =   90
            TabIndex        =   42
            Top             =   840
            Width           =   4125
         End
         Begin VB.Label Label9 
            Caption         =   "Cartório Responsável"
            Height          =   285
            Left            =   5490
            TabIndex        =   41
            Top             =   210
            Width           =   3765
         End
         Begin VB.Label Label8 
            Caption         =   "Certidão de Nascimento"
            Height          =   285
            Left            =   90
            TabIndex        =   40
            Top             =   210
            Width           =   4125
         End
      End
      Begin Threed.SSCommand cmdRemover 
         Height          =   330
         Left            =   1920
         TabIndex        =   67
         Top             =   4020
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
         Picture         =   "frmCadastroDeJogador.frx":140A
         Caption         =   "Remover Foto"
         ButtonStyle     =   3
         PictureAlignment=   1
      End
      Begin Threed.SSCommand cmdAdicionar 
         Height          =   330
         Left            =   60
         TabIndex        =   68
         Top             =   4020
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
         Picture         =   "frmCadastroDeJogador.frx":172C
         Caption         =   "        Adicionar Foto"
         ButtonStyle     =   3
         PictureAlignment=   1
      End
      Begin Threed.SSFrame SSFrame 
         Height          =   3765
         Index           =   1
         Left            =   60
         TabIndex        =   72
         Top             =   210
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   6641
         _Version        =   196609
         Begin VB.Image imgFotoJogador 
            Height          =   3675
            Left            =   30
            Stretch         =   -1  'True
            Top             =   30
            Width           =   3690
         End
      End
   End
   Begin MSComctlLib.Toolbar tbBotoes 
      Height          =   570
      Left            =   12465
      TabIndex        =   39
      Top             =   9705
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   1005
      ButtonWidth     =   2355
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
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
            Caption         =   "F10-Sair"
            Key             =   "cmdSair"
            Object.ToolTipText     =   "Sair da tela"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   180
      Top             =   9600
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
            Picture         =   "frmCadastroDeJogador.frx":1E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogador.frx":23D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogador.frx":2972
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogador.frx":2F0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogador.frx":34A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogador.frx":3A40
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogador.frx":3FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogador.frx":4574
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogador.frx":4B0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroDeJogador.frx":50A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   69
      Top             =   7800
      Width           =   11835
      _ExtentX        =   20876
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
            Object.Width           =   13520
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
   Begin MSComDlg.CommonDialog dlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCadastroDeJogador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstrFlag As String
Dim mstrFoto As String

Public Property Let DiretorioFotoJogador(strDiretorio As String)
    mstrFoto = strDiretorio
End Property

Private Sub cmdAdicionar_Click()
    On Error GoTo Erro
    
    frmAdicionarFotoJogador.Show vbModal
  
    If mstrFoto <> "" Then
        imgFotoJogador.Picture = Nothing
        imgFotoJogador.Stretch = True
        imgFotoJogador.Picture = LoadPicture(mstrFoto)
    End If
    
'
'              modJogador_AdicionarAlterarFotoJogador udtJogador.lngCodigo
'
'    Call FileCopy(txtFoto.Text, "C:\Program Files\TesteDirPadrao")
    Exit Sub
Erro:
 Call MsgBox("Erro no módulo: " & "frmCadastroDeJogador" & vbCrLf & "No Procedimento: " & "cmdAdicionar_Click" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Private Sub cmdRemover_Click()
    imgFotoJogador.Picture = Nothing
    mstrFoto = ""
End Sub

Private Sub Form_Load()
    Sta.Panels(1).Text = gSMConexao.LoginUsuario
    Sta.Panels(2).Text = gSMConexao.NomeBaseDados
    
    mstrFlag = ""
    
    Call LimparCampos
    Call HabilitarCampos(False)
    
End Sub

Private Sub LimparCampos()
    
    mstrFlag = ""
    
    txtCodigoInterno.Text = ""
    txtApelido.Text = ""
    txtNomeJogador.Text = ""
    sscClube.Text = ""
    txtLocalDeNascimento.Text = ""
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
    
    txtLocalDeNascimento.Locked = Not blnhabilitar
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
    
    dtcDataUltimaAlteracao.AllowEdit = False
    dtcDataNascimento.AllowEdit = blnhabilitar
    
    sscClube.Enabled = blnhabilitar
    sscUfEscola.Enabled = blnhabilitar
    sscUfEnderecoAtleta.Enabled = blnhabilitar
    
    cmdAdicionar.Enabled = blnhabilitar
    cmdRemover.Enabled = blnhabilitar
End Sub

Private Sub HabilitarTBBotoes(blnNovo As Boolean, blnAlterar As Boolean, blnProcurar As Boolean, blnAbandonar As Boolean, blnGravar As Boolean, blnSair As Boolean)

    tbBotoes.Buttons("cmdNovo").Enabled = blnNovo
    tbBotoes.Buttons("cmdAlterar").Enabled = blnAlterar
    tbBotoes.Buttons("cmdProcurar").Enabled = blnProcurar
    tbBotoes.Buttons("cmdLimpar").Enabled = blnAbandonar
    tbBotoes.Buttons("cmdGravar").Enabled = blnGravar
    tbBotoes.Buttons("cmdSair").Enabled = blnSair
    
End Sub


Private Sub tbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Not (Button.Enabled) Then Exit Sub
    Select Case Button.Key
        
        Case "cmdNovo":
            mstrFlag = "I"
            Call HabilitarCampos(True)
            Call HabilitarTBBotoes(False, False, False, True, True, False)
            txtApelido.SetFocus
        
        Case "cmdAlterar":
            mstrFlag = "A"
            Call HabilitarCampos(True)
            Call HabilitarTBBotoes(False, False, False, True, True, False)
            txtApelido.SetFocus
        
        Case "cmdLimpar":
            mstrFlag = ""
            Call LimparCampos
            Call HabilitarCampos(False)
            Call HabilitarTBBotoes(True, False, True, False, False, True)
             txtCodigoInterno.SetFocus
        
        Case "cmdGravar"
            If VerificarCampos Then
                GravarJogador
                CarregarJogador txtCodigoInterno.Text
                mstrFlag = ""
            Else: Exit Sub
            End If
            Call HabilitarCampos(False)
            Call HabilitarTBBotoes(False, True, True, True, False, False)
            txtCodigoInterno.SetFocus
            
        Case "cmdProcurar"
        Dim ObjRelatorioJogador As ClsRelJogador
        Set ObjRelatorioJogador = New ClsRelJogador
        
        If Not gSMConexao Is Nothing Then
            If gSMConexao.EstadoConexaoBD = adStateOpen Then
                
                ObjRelatorioJogador.Show gSMConexao, "ProFut - Relatório de Jogador"
                Exit Sub
            Else
                gSMConexao.conectar
            End If
        End If
        
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

Private Sub GravarJogador()
On Error GoTo Erro
      
    Dim udtJogador As TypJogador
    
    With udtJogador
    
        .lngCodigo = IIf(txtCodigoInterno.Text <> "", Val(txtCodigoInterno.Text), 0)
        .strApelido = txtApelido.Text
        .strNomeAtleta = txtNomeJogador.Text
        .lngCartegoria = sscCartegoria.Columns("chcodigo").Value
        .lngEquipe = sscClube.Columns(1).Value
        .strLocalNascimento = txtLocalDeNascimento.Text
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
        .strEnderecoImagem = mstrFoto
        .lngSexo = IIf(optMasculino.Value = vbChecked, 1, 2)
        .lngNumeroCamisa = txtNumCamisa.Text
    End With
    
    
    If mstrFlag = "I" Then
        Call modJogador_AdicionarJogador(udtJogador)
       
    ElseIf mstrFlag = "A" Then
        Call modJogador_AlterarJogador(udtJogador)
    End If
    
    
    txtCodigoInterno.Text = udtJogador.lngCodigo
    
'
'    Dim x As Scripting.FileSystemObject
'
'    Set x = CreateObject("scripting.filesystemobject")
'
'    x.GetFile(mstrFoto).Copy "C:\Users\FALCO\Desktop\VisualizadorVB", False
'
'
'    Call FileCopy(mstrFoto, "C:\Users\FALCO\Desktop\VisualizadorVB")


Exit Sub
Erro:
    Call MsgBox("Erro no módulo: " & "frmCadastroDeJogador" & vbCrLf & "No Procedimento: " & "GravarJogador" & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Número: " & Err.Number & vbCrLf & "Na linha: " & Erl & vbCrLf & "Entre em contato com o suporte e mostre esta mensagem!", vbOKOnly + vbCritical, "Atenção!")
End Sub

Private Sub CarregarJogador(lngCodigo As Long)
On Error GoTo Erro
    Dim objRsJogador As Recordset
    Set objRsJogador = New Recordset
      
    Call LimparCampos
    modJogador_SelecionarJogadorPorCodigo lngCodigo, objRsJogador
    
    If Not objRsJogador Is Nothing Then
        If Not objRsJogador.EOF And Not objRsJogador.BOF Then
            txtCodigoInterno.Text = NZ(objRsJogador!ID_JOGADOR_IN)
            txtApelido.Text = NS(objRsJogador!APELIDO_VC)
            txtNomeJogador.Text = NS(objRsJogador!NOMEATLETA_VC)
            'sscClube.Text = NS(objRsJogador!EQUIPE_IN)
            txtLocalDeNascimento.Text = NS(objRsJogador!LOCALNASCIMENTO_VC)
            txtNomePai.Text = NS(objRsJogador!NOMEPAI_VC)
            txtNomeMae.Text = NS(objRsJogador!STRNOMEMAE_VC)
            txtUsuarioAlteracao.Text = NS(objRsJogador!USUARIOULTIMAALTERACAO_VC)
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
            
            chkWpp1.Value = IIf(NB(objRsJogador!WPP1_BT), vbChecked, vbUnchecked)
            chkwpp2.Value = IIf(NB(objRsJogador!WPP2_BT), vbChecked, vbUnchecked)
            
            dtcDataUltimaAlteracao.DateValue = NS(objRsJogador!DATAULTIMAALTERACAO_DT)
            dtcDataNascimento.DateValue = ND(objRsJogador!DATANASCIMENTO_DT)
            
            mstrFoto = NS(objRsJogador!ENDERECOIMAGEM_VC)
            If mstrFoto <> "" Then
                imgFotoJogador.Picture = Nothing
                imgFotoJogador.Stretch = True
                On Error Resume Next
                imgFotoJogador.Picture = LoadPicture(mstrFoto)
                On Error GoTo Erro
            End If
            
            modBDCombo_SelecionarCartecoriaJogador sscCartegoria, NZ(objRsJogador!CARTEGORIA_IN)
            modBDCombo_SelecionarEquipePorCodigo sscClube, NZ(objRsJogador!EQUIPE_IN)
            modBDCombo_SelecionarEstados sscUfEscola, NZ(objRsJogador!ESTADOESCOLA_IN)
            modBDCombo_SelecionarEstados sscUfEnderecoAtleta, NZ(objRsJogador!Estado_IN)
            
            mstrFlag = ""
            Call HabilitarCampos(False)
            Call HabilitarTBBotoes(False, True, True, True, False, True)
            
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
