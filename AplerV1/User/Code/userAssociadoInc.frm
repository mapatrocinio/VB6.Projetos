VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserAssociadoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de associado"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   6180
      Left            =   8430
      ScaleHeight     =   6180
      ScaleWidth      =   1860
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4605
         Left            =   90
         ScaleHeight     =   4545
         ScaleWidth      =   1605
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1665
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&X"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   3570
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   2730
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5985
      Left            =   120
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   10557
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "&Cadastro"
      TabPicture(0)   =   "userAssociadoInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&End. res."
      TabPicture(1)   =   "userAssociadoInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Titular"
      TabPicture(2)   =   "userAssociadoInc.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "End. &com./cob."
      TabPicture(3)   =   "userAssociadoInc.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame4"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "&Profissão"
      TabPicture(4)   =   "userAssociadoInc.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "grdProfissao"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Convênio"
      TabPicture(5)   =   "userAssociadoInc.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "grdConvenio"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Depe&ndente"
      TabPicture(6)   =   "userAssociadoInc.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "grdDependente"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "&Linha"
      TabPicture(7)   =   "userAssociadoInc.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "grdLinha"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).ControlCount=   1
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4755
         Left            =   -74910
         TabIndex        =   106
         Top             =   390
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   4515
            Index           =   3
            Left            =   120
            ScaleHeight     =   4515
            ScaleWidth      =   7575
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.ComboBox cboEmpresa 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   32
               Top             =   720
               Width           =   6105
            End
            Begin VB.ComboBox cboCaptador 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   1440
               Width           =   6105
            End
            Begin VB.ComboBox cboOrigem 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   33
               Top             =   1080
               Width           =   6105
            End
            Begin VB.TextBox txtContrato 
               Height          =   285
               Left            =   5220
               MaxLength       =   30
               TabIndex        =   29
               Text            =   "txtContrato"
               Top             =   90
               Width           =   2175
            End
            Begin MSMask.MaskEdBox mskDtInicio 
               Height          =   255
               Left            =   1320
               TabIndex        =   30
               Top             =   420
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDtFim 
               Height          =   255
               Left            =   5220
               TabIndex        =   31
               Top             =   420
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskMatricula 
               Height          =   255
               Left            =   1320
               TabIndex        =   28
               Top             =   120
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,###;($#,###)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Empresa"
               Height          =   195
               Index           =   37
               Left            =   60
               TabIndex        =   119
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Captador"
               Height          =   195
               Index           =   36
               Left            =   60
               TabIndex        =   113
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Origem"
               Height          =   195
               Index           =   35
               Left            =   60
               TabIndex        =   112
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Dt. Fim"
               Height          =   255
               Index           =   2
               Left            =   3960
               TabIndex        =   111
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Dt. Início"
               Height          =   255
               Index           =   0
               Left            =   60
               TabIndex        =   110
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Matricula 
               Caption         =   "Matrícula"
               Height          =   195
               Index           =   35
               Left            =   60
               TabIndex        =   109
               Top             =   90
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Contrato"
               Height          =   195
               Index           =   34
               Left            =   3960
               TabIndex        =   108
               Top             =   90
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comercial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2475
         Left            =   -74910
         TabIndex        =   97
         Top             =   2490
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2175
            Index           =   5
            Left            =   120
            ScaleHeight     =   2175
            ScaleWidth      =   7575
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   180
            Width           =   7575
            Begin VB.TextBox txtTelefoneCom1 
               Height          =   285
               Left            =   1320
               MaxLength       =   30
               TabIndex        =   49
               Text            =   "txtTelefoneCom1"
               Top             =   1740
               Width           =   2175
            End
            Begin VB.TextBox txtTelefoneCom2 
               Height          =   285
               Left            =   5220
               MaxLength       =   30
               TabIndex        =   50
               Text            =   "txtTelefoneCom2"
               Top             =   1740
               Width           =   2175
            End
            Begin VB.TextBox txtEstadoCom 
               Height          =   285
               Left            =   1320
               MaxLength       =   2
               TabIndex        =   45
               Text            =   "txtEstadoCom"
               Top             =   750
               Width           =   435
            End
            Begin VB.TextBox txtComplementoCom 
               Height          =   285
               Left            =   5220
               MaxLength       =   30
               TabIndex        =   44
               Text            =   "txtComplementoCom"
               Top             =   420
               Width           =   2175
            End
            Begin VB.TextBox txtNumeroCom 
               Height          =   285
               Left            =   1320
               MaxLength       =   15
               TabIndex        =   43
               Text            =   "txtNumeroCom"
               Top             =   420
               Width           =   2175
            End
            Begin VB.TextBox txtRuaCom 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   42
               Text            =   "txtRuaCom"
               Top             =   90
               Width           =   6075
            End
            Begin VB.TextBox txtBairroCom 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   47
               Text            =   "txtBairroCom"
               Top             =   1080
               Width           =   6075
            End
            Begin VB.TextBox txtCidadeCom 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   48
               Text            =   "txtCidadeCom"
               Top             =   1410
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskCepCom 
               Height          =   255
               Left            =   5220
               TabIndex        =   46
               Top             =   750
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##.###-###"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Telefone com. 1"
               Height          =   195
               Index           =   41
               Left            =   60
               TabIndex        =   115
               Top             =   1740
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Telefone com. 2"
               Height          =   195
               Index           =   40
               Left            =   3960
               TabIndex        =   114
               Top             =   1740
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Estado"
               Height          =   195
               Index           =   30
               Left            =   60
               TabIndex        =   105
               Top             =   750
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Complemento"
               Height          =   195
               Index           =   29
               Left            =   3960
               TabIndex        =   104
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Número"
               Height          =   195
               Index           =   26
               Left            =   60
               TabIndex        =   103
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Rua"
               Height          =   195
               Index           =   25
               Left            =   60
               TabIndex        =   102
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Bairro"
               Height          =   195
               Index           =   22
               Left            =   60
               TabIndex        =   101
               Top             =   1125
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cidade"
               Height          =   195
               Index           =   21
               Left            =   60
               TabIndex        =   100
               Top             =   1455
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cep"
               Height          =   195
               Index           =   20
               Left            =   3960
               TabIndex        =   99
               Top             =   750
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cobrança"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -74910
         TabIndex        =   88
         Top             =   390
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   1815
            Index           =   4
            Left            =   120
            ScaleHeight     =   1815
            ScaleWidth      =   7575
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   180
            Width           =   7575
            Begin VB.TextBox txtCidadeCob 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   41
               Text            =   "txtCidadeCob"
               Top             =   1410
               Width           =   6075
            End
            Begin VB.TextBox txtBairroCob 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   40
               Text            =   "txtBairroCob"
               Top             =   1080
               Width           =   6075
            End
            Begin VB.TextBox txtRuaCob 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   35
               Text            =   "txtRuaCob"
               Top             =   90
               Width           =   6075
            End
            Begin VB.TextBox txtNumeroCob 
               Height          =   285
               Left            =   1320
               MaxLength       =   15
               TabIndex        =   36
               Text            =   "txtNumeroCob"
               Top             =   420
               Width           =   2175
            End
            Begin VB.TextBox txtComplementoCob 
               Height          =   285
               Left            =   5220
               MaxLength       =   30
               TabIndex        =   37
               Text            =   "txtComplementoCob"
               Top             =   420
               Width           =   2175
            End
            Begin VB.TextBox txtEstadoCob 
               Height          =   285
               Left            =   1320
               MaxLength       =   2
               TabIndex        =   38
               Text            =   "txtEstadoCob"
               Top             =   750
               Width           =   435
            End
            Begin MSMask.MaskEdBox mskCepCob 
               Height          =   255
               Left            =   5220
               TabIndex        =   39
               Top             =   750
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##.###-###"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Cep"
               Height          =   195
               Index           =   19
               Left            =   3960
               TabIndex        =   96
               Top             =   750
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cidade"
               Height          =   195
               Index           =   18
               Left            =   60
               TabIndex        =   95
               Top             =   1455
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Bairro"
               Height          =   195
               Index           =   17
               Left            =   60
               TabIndex        =   94
               Top             =   1125
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Rua"
               Height          =   195
               Index           =   15
               Left            =   60
               TabIndex        =   93
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Número"
               Height          =   195
               Index           =   14
               Left            =   60
               TabIndex        =   92
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Complemento"
               Height          =   195
               Index           =   13
               Left            =   3960
               TabIndex        =   91
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Estado"
               Height          =   195
               Index           =   12
               Left            =   60
               TabIndex        =   90
               Top             =   750
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4755
         Left            =   -74880
         TabIndex        =   79
         Top             =   360
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   4515
            Index           =   2
            Left            =   120
            ScaleHeight     =   4515
            ScaleWidth      =   7575
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtEspecial 
               Height          =   285
               Left            =   1320
               MaxLength       =   20
               TabIndex        =   27
               Text            =   "txtEspecial"
               Top             =   1740
               Width           =   2175
            End
            Begin VB.TextBox txtEstadoRes 
               Height          =   285
               Left            =   1320
               MaxLength       =   2
               TabIndex        =   23
               Text            =   "txtEstadoRes"
               Top             =   750
               Width           =   435
            End
            Begin VB.TextBox txtComplementoRes 
               Height          =   285
               Left            =   5220
               MaxLength       =   30
               TabIndex        =   22
               Text            =   "txtComplementoRes"
               Top             =   420
               Width           =   2175
            End
            Begin VB.TextBox txtNumeroRes 
               Height          =   285
               Left            =   1320
               MaxLength       =   15
               TabIndex        =   21
               Text            =   "txtNumeroRes"
               Top             =   420
               Width           =   2175
            End
            Begin VB.TextBox txtRuaRes 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   20
               Text            =   "txtRuaRes"
               Top             =   90
               Width           =   6075
            End
            Begin VB.TextBox txtBairroRes 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   25
               Text            =   "txtBairroRes"
               Top             =   1080
               Width           =   6075
            End
            Begin VB.TextBox txtCidadeRes 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   26
               Text            =   "txtCidadeRes"
               Top             =   1410
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskCepRes 
               Height          =   255
               Left            =   5220
               TabIndex        =   24
               Top             =   750
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##.###-###"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Especial"
               Height          =   195
               Index           =   23
               Left            =   60
               TabIndex        =   120
               Top             =   1740
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Estado"
               Height          =   195
               Index           =   9
               Left            =   60
               TabIndex        =   87
               Top             =   750
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Complemento"
               Height          =   195
               Index           =   8
               Left            =   3960
               TabIndex        =   86
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Número"
               Height          =   195
               Index           =   7
               Left            =   60
               TabIndex        =   85
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Rua"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   84
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Bairro"
               Height          =   195
               Index           =   2
               Left            =   60
               TabIndex        =   83
               Top             =   1125
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cidade"
               Height          =   195
               Index           =   16
               Left            =   60
               TabIndex        =   82
               Top             =   1455
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cep"
               Height          =   195
               Index           =   3
               Left            =   3960
               TabIndex        =   81
               Top             =   750
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   120
         TabIndex        =   63
         Top             =   360
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   855
            Index           =   1
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   7575
            TabIndex        =   116
            Top             =   4650
            Width           =   7575
            Begin VB.ComboBox cboGrauParentesco 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   360
               Width           =   6105
            End
            Begin MSMask.MaskEdBox mskMatriculaDep 
               Height          =   255
               Left            =   1320
               TabIndex        =   18
               Top             =   60
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,###;($#,###)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Grau Parentesco"
               Height          =   195
               Index           =   33
               Left            =   60
               TabIndex        =   118
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Matricula 
               Caption         =   "Matrícula"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   117
               Top             =   30
               Width           =   1215
            End
         End
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   4485
            Index           =   0
            Left            =   120
            ScaleHeight     =   4485
            ScaleWidth      =   7575
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtNaturalidade 
               Height          =   285
               Left            =   6690
               MaxLength       =   2
               TabIndex        =   8
               Text            =   "txtNaturalidade"
               Top             =   1320
               Width           =   705
            End
            Begin VB.TextBox txtOrgaoEmissor 
               Height          =   285
               Left            =   1320
               MaxLength       =   20
               TabIndex        =   6
               Text            =   "txtOrgaoEmissor"
               Top             =   990
               Width           =   2325
            End
            Begin VB.TextBox txtIdentidade 
               Height          =   285
               Left            =   4830
               MaxLength       =   20
               TabIndex        =   5
               Text            =   "txtIdentidade"
               Top             =   630
               Width           =   2565
            End
            Begin VB.ComboBox cboEstadoCivil 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   1320
               Width           =   3765
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1350
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   77
               TabStop         =   0   'False
               Top             =   4050
               Width           =   2235
               Begin VB.OptionButton optExcluido 
                  Caption         =   "Sim"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   16
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1065
               End
               Begin VB.OptionButton optExcluido 
                  Caption         =   "Não"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   17
                  Top             =   0
                  Width           =   1095
               End
            End
            Begin VB.ComboBox cboValorPlano 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   2040
               Width           =   6105
            End
            Begin VB.TextBox txtNomeMae 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   11
               Text            =   "txtNomeMae"
               Top             =   2400
               Width           =   6075
            End
            Begin VB.TextBox txtObservacao 
               Height          =   615
               Left            =   1320
               MaxLength       =   255
               MultiLine       =   -1  'True
               TabIndex        =   15
               Text            =   "userAssociadoInc.frx":00E0
               Top             =   3390
               Width           =   6075
            End
            Begin VB.TextBox txtEmail 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   14
               Text            =   "txtEmail"
               Top             =   3060
               Width           =   6075
            End
            Begin VB.TextBox txtCelular 
               Height          =   285
               Left            =   5220
               MaxLength       =   30
               TabIndex        =   13
               Text            =   "txtCelular"
               Top             =   2730
               Width           =   2175
            End
            Begin VB.TextBox txtTelefoneRes 
               Height          =   285
               Left            =   1320
               MaxLength       =   30
               TabIndex        =   12
               Text            =   "txtTelefoneRes"
               Top             =   2730
               Width           =   2175
            End
            Begin VB.ComboBox cboTipoSocio 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   1680
               Width           =   6105
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   5190
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   360
               Width           =   2235
               Begin VB.OptionButton optSexo 
                  Caption         =   "Feminino"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   3
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton optSexo 
                  Caption         =   "Masculino"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   2
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1065
               End
            End
            Begin VB.TextBox txtNome 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   0
               Text            =   "txtNome"
               Top             =   75
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskDtNascimento 
               Height          =   255
               Left            =   1320
               TabIndex        =   4
               Top             =   690
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskCpf 
               Height          =   255
               Left            =   1320
               TabIndex        =   1
               Top             =   390
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   14
               Mask            =   "###.###.###-##"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Naturalidade"
               Height          =   195
               Index           =   43
               Left            =   5520
               TabIndex        =   124
               Top             =   1365
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Órg. Emissor"
               Height          =   195
               Index           =   42
               Left            =   60
               TabIndex        =   123
               Top             =   1035
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Identidade"
               Height          =   195
               Index           =   39
               Left            =   3870
               TabIndex        =   122
               Top             =   675
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Estado Cívil"
               Height          =   195
               Index           =   38
               Left            =   60
               TabIndex        =   121
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Excluido"
               Height          =   315
               Index           =   11
               Left            =   120
               TabIndex        =   78
               Top             =   4080
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Plano"
               Height          =   195
               Index           =   10
               Left            =   60
               TabIndex        =   76
               Top             =   2040
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Nome da mãe"
               Height          =   195
               Index           =   6
               Left            =   60
               TabIndex        =   75
               Top             =   2445
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Observação"
               Height          =   195
               Index           =   32
               Left            =   60
               TabIndex        =   74
               Top             =   3420
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "E-mail"
               Height          =   195
               Index           =   31
               Left            =   60
               TabIndex        =   73
               Top             =   3060
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Celular"
               Height          =   195
               Index           =   28
               Left            =   3960
               TabIndex        =   72
               Top             =   2730
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Telefone res."
               Height          =   195
               Index           =   27
               Left            =   60
               TabIndex        =   71
               Top             =   2730
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "CPF"
               Height          =   195
               Index           =   4
               Left            =   60
               TabIndex        =   70
               Top             =   390
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Dt. Nascimento"
               Height          =   255
               Index           =   1
               Left            =   60
               TabIndex        =   69
               Top             =   690
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Tipo Sócio"
               Height          =   195
               Index           =   24
               Left            =   60
               TabIndex        =   68
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Sexo"
               Height          =   315
               Index           =   5
               Left            =   3870
               TabIndex        =   66
               Top             =   390
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Nome"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   65
               Top             =   120
               Width           =   1215
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdProfissao 
         Height          =   4725
         Left            =   -74910
         OleObjectBlob   =   "userAssociadoInc.frx":00EE
         TabIndex        =   51
         Top             =   390
         Width           =   7905
      End
      Begin TrueDBGrid60.TDBGrid grdConvenio 
         Height          =   4725
         Left            =   -74910
         OleObjectBlob   =   "userAssociadoInc.frx":4648
         TabIndex        =   52
         Top             =   390
         Width           =   7905
      End
      Begin TrueDBGrid60.TDBGrid grdDependente 
         Height          =   4725
         Left            =   -74910
         OleObjectBlob   =   "userAssociadoInc.frx":94D5
         TabIndex        =   53
         Top             =   390
         Width           =   7905
      End
      Begin TrueDBGrid60.TDBGrid grdLinha 
         Height          =   4725
         Left            =   -74910
         OleObjectBlob   =   "userAssociadoInc.frx":E7F4
         TabIndex        =   54
         Top             =   390
         Width           =   4665
      End
   End
End
Attribute VB_Name = "frmUserAssociadoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean
Public blnPrimeiraVez           As Boolean
Public strNomeAssociado         As String
Public strIcAssociado           As String
  'T - TITULAR
  'D - DEPENDENTE

Public lngPKID                  As Long
Public lngASSOCIADOTITULARID    As Long

Dim PROF_COLUNASMATRIZ          As Long
Dim PROF_LINHASMATRIZ           As Long
Private PROF_Matriz()           As String

Dim CONV_COLUNASMATRIZ          As Long
Dim CONV_LINHASMATRIZ           As Long
Private CONV_Matriz()           As String

Dim DEP_COLUNASMATRIZ           As Long
Dim DEP_LINHASMATRIZ            As Long
Private DEP_Matriz()            As String

Dim LIN_COLUNASMATRIZ           As Long
Dim LIN_LINHASMATRIZ            As Long
Private LIN_Matriz()            As String

Private Sub TratarCampos()
  On Error GoTo trata
  If strIcAssociado = "T" Then
    'Titular
    'Associao
    pictrava(0).Visible = True
    'Dependente
    pictrava(1).Visible = False
    'Titular
    tabDetalhes.TabVisible(2) = True
    'End com e cob
    tabDetalhes.TabVisible(3) = True
    'Dependentes
    tabDetalhes.TabVisible(6) = True
    'Linha
    tabDetalhes.TabVisible(7) = True
  Else
    'Dependente
    'Associao
    pictrava(0).Visible = True
    'Dependente
    pictrava(1).Visible = True
    'Titular
    tabDetalhes.TabVisible(2) = False
    'End com e cob
    tabDetalhes.TabVisible(3) = False
    'Dependentes
    tabDetalhes.TabVisible(6) = False
    'Linha
    tabDetalhes.TabVisible(7) = False
    'Caption form
    Me.Caption = Me.Caption & " - Titular (" & strNomeAssociado & ")"
  End If
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserAssociadoInc.TratarCampos]", _
            Err.Description
End Sub

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Associado
  'Dados cadastrais
  LimparCampoTexto txtNome
  LimparCampoMask mskCpf
  optSexo(0).Value = False
  optSexo(1).Value = False
  LimparCampoMask mskDtNascimento
  LimparCampoTexto txtIdentidade
  LimparCampoTexto txtOrgaoEmissor
  LimparCampoCombo cboEstadoCivil
  LimparCampoTexto txtNaturalidade
  LimparCampoTexto txtEspecial
  LimparCampoCombo cboTipoSocio
  LimparCampoCombo cboValorPlano
  LimparCampoTexto txtNomeMae
  LimparCampoTexto txtTelefoneRes
  LimparCampoTexto txtCelular
  LimparCampoTexto txtEmail
  LimparCampoTexto txtObservacao
  optExcluido(0).Value = False
  optExcluido(1).Value = False
  LimparCampoMask mskMatriculaDep
  LimparCampoCombo cboGrauParentesco
  'Endereço res
  LimparCampoTexto txtRuaRes
  LimparCampoTexto txtNumeroRes
  LimparCampoTexto txtComplementoRes
  LimparCampoTexto txtEstadoRes
  LimparCampoMask mskCepRes
  LimparCampoTexto txtBairroRes
  LimparCampoTexto txtCidadeRes
  'Titular
  LimparCampoMask mskMatricula
  LimparCampoTexto txtContrato
  LimparCampoMask mskDtInicio
  LimparCampoMask mskDtFim
  LimparCampoCombo cboOrigem
  LimparCampoCombo cboCaptador
  LimparCampoCombo cboEmpresa
  'Endereço Cobrança
  LimparCampoTexto txtRuaCob
  LimparCampoTexto txtNumeroCob
  LimparCampoTexto txtComplementoCob
  LimparCampoTexto txtEstadoCob
  LimparCampoMask mskCepCob
  LimparCampoTexto txtBairroCob
  LimparCampoTexto txtCidadeCob
  'Endereço Comercial
  LimparCampoTexto txtRuaCom
  LimparCampoTexto txtNumeroCom
  LimparCampoTexto txtComplementoCom
  LimparCampoTexto txtEstadoCom
  LimparCampoMask mskCepCom
  LimparCampoTexto txtBairroCom
  LimparCampoTexto txtCidadeCom
  LimparCampoTexto txtTelefoneCom1
  LimparCampoTexto txtTelefoneCom2
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserAssociadoInc.LimparCampos]", _
            Err.Description
End Sub


Private Sub cboCaptador_LostFocus()
  Pintar_Controle cboCaptador, tpCorContr_Normal
End Sub

Private Sub cboEmpresa_LostFocus()
  Pintar_Controle cboEmpresa, tpCorContr_Normal
End Sub

Private Sub cboEstadoCivil_LostFocus()
  Pintar_Controle cboEstadoCivil, tpCorContr_Normal
End Sub

Private Sub cboGrauParentesco_LostFocus()
  Pintar_Controle cboGrauParentesco, tpCorContr_Normal
End Sub

Private Sub cboOrigem_LostFocus()
  Pintar_Controle cboOrigem, tpCorContr_Normal
End Sub

Private Sub cboTipoSocio_LostFocus()
  Pintar_Controle cboTipoSocio, tpCorContr_Normal
End Sub

Private Sub cboValorPlano_LostFocus()
  Pintar_Controle cboValorPlano, tpCorContr_Normal
End Sub

Private Sub cmdAlterar_Click()
  On Error GoTo trata
  Dim objFormAssociado As Apler.frmUserAssociadoInc
  Select Case tabDetalhes.Tab
  Case 5
    'Convênios
    If Not IsNumeric(grdConvenio.Columns("PKID").Value & "") Then
      MsgBox "Selecione um plano !", vbExclamation, TITULOSISTEMA
      SetarFoco grdConvenio
      Exit Sub
    End If

    frmUserConvAssocInc.lngPKID = grdConvenio.Columns("PKID").Value
    frmUserConvAssocInc.lngASSOCIADOID = lngPKID
    frmUserConvAssocInc.strNomeAssociado = txtNome.Text
    frmUserConvAssocInc.Status = tpStatus_Alterar
    frmUserConvAssocInc.Show vbModal

    If frmUserConvAssocInc.blnRetorno Then
      CONV_MontaMatriz
      grdConvenio.Bookmark = Null
      grdConvenio.ReBind
      grdConvenio.ApproxCount = CONV_LINHASMATRIZ
    End If
    SetarFoco grdConvenio
  Case 6
    'Dependente
    If Not IsNumeric(grdDependente.Columns("PKID").Value & "") Then
      MsgBox "Selecione um dependente !", vbExclamation, TITULOSISTEMA
      SetarFoco grdDependente
      Exit Sub
    End If

    Set objFormAssociado = New Apler.frmUserAssociadoInc
    objFormAssociado.Status = tpStatus_Alterar
    objFormAssociado.lngPKID = grdDependente.Columns("PKID").Value
    objFormAssociado.lngASSOCIADOTITULARID = lngPKID
    objFormAssociado.strIcAssociado = "D"
    objFormAssociado.strNomeAssociado = txtNome.Text
    objFormAssociado.Show vbModal
    If objFormAssociado.blnRetorno Then
      DEP_MontaMatriz
      grdDependente.Bookmark = Null
      grdDependente.ReBind
      grdDependente.ApproxCount = DEP_LINHASMATRIZ
    End If
    Set objFormAssociado = Nothing
    SetarFoco grdDependente
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdCancelar_Click()
  blnFechar = True
  blnRetorno = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdExcluir_Click()
  On Error GoTo trata
  Dim objAssociado      As busApler.clsAssociado
  Dim objConvAssoc      As busApler.clsConvAssoc
  '
  Select Case tabDetalhes.Tab
  Case 5 'Exclusão de Plano
    '
    If Len(Trim(grdConvenio.Columns("PKID").Value & "")) = 0 Then
      MsgBox "Selecione um plano.", vbExclamation, TITULOSISTEMA
      SetarFoco grdConvenio
      Exit Sub
    End If
    '
    Set objConvAssoc = New busApler.clsConvAssoc
    '
    If MsgBox("Confirma exclusão do plano" & grdConvenio.Columns("Plano").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
      SetarFoco grdConvenio
      Exit Sub
    End If
    'OK
    objConvAssoc.ExcluirConvAssoc CLng(grdConvenio.Columns("PKID").Value)
    '
    CONV_MontaMatriz
    grdConvenio.Bookmark = Null
    grdConvenio.ReBind
    grdConvenio.ApproxCount = CONV_LINHASMATRIZ

    Set objConvAssoc = Nothing
    SetarFoco grdConvenio
  Case 6 'Exclusão de associado
    If Not IsNumeric(grdDependente.Columns("PKID").Value & "") Then
      MsgBox "Selecione um associado para exclusão.", vbExclamation, TITULOSISTEMA
      SetarFoco grdDependente
      Exit Sub
    End If
    '
    If MsgBox("ATENÇÃO: A exclusão do associado removerá todas associações de pagamento e convênios." & vbCrLf & "Caso queira você pode apenas alterá-lo e selecionar a opção excluído, isso irá excluílo logicamente, mantendo suas informações na base de dados." & vbCrLf & "Confirma exclusão do associado " & grdDependente.Columns("Nome").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
      SetarFoco grdDependente
      Exit Sub
    End If
    'OK
    Set objAssociado = New busApler.clsAssociado
    objAssociado.ExcluirAssociado CLng(grdDependente.Columns("PKID").Value), _
                                  "D"
    Set objAssociado = Nothing
    '
    DEP_MontaMatriz
    grdDependente.Bookmark = Null
    grdDependente.ReBind
    SetarFoco grdDependente
    
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdIncluir_Click()
  On Error GoTo trata
  Dim objFormAssociado As Apler.frmUserAssociadoInc
  
  '
  Select Case tabDetalhes.Tab
  Case 4
    'Profissão
    frmUserProfAssocInc.lngASSOCIADOID = lngPKID
    frmUserProfAssocInc.strNomeAssociado = txtNome.Text
    frmUserProfAssocInc.Show vbModal

    If frmUserProfAssocInc.blnRetorno Then
      PROF_MontaMatriz
      grdProfissao.Bookmark = Null
      grdProfissao.ReBind
      grdProfissao.ApproxCount = PROF_LINHASMATRIZ
    End If
    SetarFoco grdProfissao
  Case 5
    'Plano Convênio
    frmUserConvAssocInc.Status = tpStatus_Incluir
    frmUserConvAssocInc.lngPKID = 0
    frmUserConvAssocInc.lngASSOCIADOID = lngPKID
    frmUserConvAssocInc.strNomeAssociado = txtNome.Text
    frmUserConvAssocInc.Show vbModal

    If frmUserConvAssocInc.blnRetorno Then
      CONV_MontaMatriz
      grdConvenio.Bookmark = Null
      grdConvenio.ReBind
      grdConvenio.ApproxCount = CONV_LINHASMATRIZ
    End If
    SetarFoco grdConvenio
  Case 6
    'DEPENDENTE
    Set objFormAssociado = New Apler.frmUserAssociadoInc
    objFormAssociado.Status = tpStatus_Incluir
    objFormAssociado.lngPKID = 0
    objFormAssociado.lngASSOCIADOTITULARID = lngPKID
    objFormAssociado.strIcAssociado = "D"
    objFormAssociado.strNomeAssociado = txtNome.Text
    objFormAssociado.Show vbModal
    If objFormAssociado.blnRetorno Then
      DEP_MontaMatriz
      grdDependente.Bookmark = Null
      grdDependente.ReBind
      grdDependente.ApproxCount = DEP_LINHASMATRIZ
    End If
    Set objFormAssociado = Nothing
    SetarFoco grdDependente
  Case 7
    'Linha
    frmUserLinhaAssocInc.lngASSOCIADOID = lngPKID
    frmUserLinhaAssocInc.strNomeAssociado = txtNome.Text
    frmUserLinhaAssocInc.Show vbModal

    If frmUserLinhaAssocInc.blnRetorno Then
      LIN_MontaMatriz
      grdLinha.Bookmark = Null
      grdLinha.ReBind
      grdLinha.ApproxCount = LIN_LINHASMATRIZ
    End If
    SetarFoco grdLinha
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Resume Next
End Sub

Private Sub cmdOK_Click()
  Dim objAssociado              As busApler.clsAssociado
  Dim objDependente             As busApler.clsAssociadoDependente
  Dim objTitular                As busApler.clsAssociadoTitular
  Dim objGeral                  As busApler.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngTIPOSOCIOID            As Long
  Dim lngVALORPLANOID           As Long
  Dim lngORIGEMID               As Long
  Dim lngCAPTADORID             As Long
  Dim lngEMPRESAID              As Long
  Dim lngGRAUPARENTESCOID       As Long
  Dim lngESTADOCIVILID          As Long
  Dim strSexo                   As String
  Dim strExcluido               As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busApler.clsGeral
  Set objAssociado = New busApler.clsAssociado
  'ESTADO CÍVIL
  lngESTADOCIVILID = 0
  strSql = "SELECT PKID FROM ESTADOCIVIL WHERE DESCRICAO = " & Formata_Dados(cboEstadoCivil.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngESTADOCIVILID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'EMPRESA
  lngEMPRESAID = 0
  strSql = "SELECT PKID FROM EMPRESA WHERE NOME = " & Formata_Dados(cboEmpresa.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngEMPRESAID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'TIPO SOCIO
  lngTIPOSOCIOID = 0
  strSql = "SELECT PKID FROM TIPOSOCIO WHERE DESCRICAO = " & Formata_Dados(cboTipoSocio.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngTIPOSOCIOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'VALOR PLANO
  lngVALORPLANOID = 0
  strSql = "SELECT PKID FROM VALORPLANO WHERE DESCRICAO = " & Formata_Dados(cboValorPlano.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngVALORPLANOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  If strIcAssociado = "T" Then
    'Titular
    'ORIGEM
    lngORIGEMID = 0
    strSql = "SELECT PKID FROM ORIGEM WHERE DESCRICAO = " & Formata_Dados(cboOrigem.Text, tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngORIGEMID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    'CAPTADOR
    lngCAPTADORID = 0
    strSql = "SELECT PKID FROM CAPTADOR WHERE NOME = " & Formata_Dados(cboCaptador.Text, tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngCAPTADORID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
  Else
    'Dependente
    'GRAU DE PARENTESCO
    lngGRAUPARENTESCOID = 0
    strSql = "SELECT PKID FROM GRAUPARENTESCO WHERE DESCRICAO = " & Formata_Dados(cboGrauParentesco.Text, tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngGRAUPARENTESCOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  'Status
  If optSexo(0).Value Then
    strSexo = "M"
  ElseIf optSexo(1).Value Then
    strSexo = "F"
  Else
    strSexo = ""
  End If
  'Status
  If optExcluido(0).Value Then
    strExcluido = "S"
  ElseIf optExcluido(1).Value Then
    strExcluido = "N"
  Else
    strExcluido = ""
  End If
  'Validar se funcionário já cadastrado
  strSql = "SELECT * FROM ASSOCIADO " & _
    " WHERE ASSOCIADO.NOME = " & Formata_Dados(txtNome.Text, tpDados_Texto) & _
    " AND ASSOCIADO.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle txtNome, tpCorContr_Erro
    TratarErroPrevisto "Associado já cadastrado"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objAssociado = Nothing
    cmdOk.Enabled = True
    SetarFoco txtNome
    tabDetalhes.Tab = 1
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar Associado
    objAssociado.AlterarAssociado lngPKID, _
                                  lngTIPOSOCIOID, _
                                  lngVALORPLANOID, _
                                  lngESTADOCIVILID, _
                                  txtIdentidade.Text, _
                                  txtOrgaoEmissor.Text, _
                                  txtNaturalidade.Text, _
                                  mskCpf.ClipText, _
                                  txtNome.Text, _
                                  strSexo, _
                                  IIf(mskDtNascimento.ClipText = "", "", mskDtNascimento.Text), _
                                  txtNomeMae.Text, _
                                  txtEspecial.Text, _
                                  txtRuaRes.Text, _
                                  txtNumeroRes.Text, _
                                  txtComplementoRes.Text, _
                                  txtEstadoRes.Text, _
                                  IIf(mskCepRes.ClipText = "", "", mskCepRes.ClipText), _
                                  txtBairroRes.Text, _
                                  txtCidadeRes.Text, _
                                  txtTelefoneRes.Text, _
                                  txtCelular.Text, _
                                  txtEmail.Text, _
                                  strExcluido, _
                                  txtObservacao.Text
    'Verificação
    If strIcAssociado = "T" Then
      'Titular
      Set objTitular = New busApler.clsAssociadoTitular
      objTitular.AlterarTitular lngPKID, _
                                lngORIGEMID, _
                                lngCAPTADORID, _
                                lngEMPRESAID, _
                                mskMatricula.ClipText, _
                                txtContrato.Text, _
                                mskDtInicio.Text, _
                                IIf(mskDtFim.ClipText = "", "", mskDtFim.Text), _
                                txtRuaCom.Text, _
                                txtNumeroCom.Text, _
                                txtComplementoCom.Text, _
                                IIf(mskCepCom.ClipText = "", "", mskCepCom.ClipText), _
                                txtBairroCom.Text, _
                                txtCidadeCom.Text, _
                                txtEstadoCom.Text, _
                                txtTelefoneCom1.Text, _
                                txtTelefoneCom2.Text, _
                                txtRuaCob.Text, _
                                txtNumeroCob.Text, _
                                txtComplementoCob.Text, _
                                IIf(mskCepCob.ClipText = "", "", mskCepCob.ClipText), _
                                txtBairroCob.Text, _
                                txtCidadeCob.Text, _
                                txtEstadoCob.Text
      Set objTitular = Nothing
    Else
      'Dependente
      Set objDependente = New busApler.clsAssociadoDependente
      objDependente.AlterarDependente lngPKID, _
                                      lngGRAUPARENTESCOID, _
                                      mskMatriculaDep.ClipText
      Set objDependente = Nothing
    End If
    '
    blnRetorno = True
    blnFechar = True
    Unload Me
  ElseIf Status = tpStatus_Incluir Then
    'Inserir Associado
    objAssociado.InserirAssociado lngTIPOSOCIOID, _
                                  lngVALORPLANOID, _
                                  lngESTADOCIVILID, _
                                  txtIdentidade.Text, _
                                  txtOrgaoEmissor.Text, _
                                  txtNaturalidade.Text, _
                                  mskCpf.ClipText, _
                                  txtNome.Text, _
                                  strSexo, _
                                  IIf(mskDtNascimento.ClipText = "", "", mskDtNascimento.Text), _
                                  txtNomeMae.Text, _
                                  txtEspecial.Text, _
                                  txtRuaRes.Text, _
                                  txtNumeroRes.Text, _
                                  txtComplementoRes.Text, _
                                  txtEstadoRes.Text, _
                                  IIf(mskCepRes.ClipText = "", "", mskCepRes.ClipText), _
                                  txtBairroRes.Text, _
                                  txtCidadeRes.Text, _
                                  txtTelefoneRes.Text, _
                                  txtCelular.Text, _
                                  txtEmail.Text, _
                                  txtObservacao.Text, _
                                  strIcAssociado

      'ASSOCIADO
      Set objRs = objAssociado.SelecionarAssociadoPeloNome(txtNome.Text)
      If Not objRs.EOF Then
        lngPKID = objRs.Fields("PKID").Value
      End If
      objRs.Close
      Set objRs = Nothing
    If strIcAssociado = "T" Then
      'Titular
      Set objTitular = New busApler.clsAssociadoTitular
      objTitular.InserirTitular lngPKID, _
                                lngORIGEMID, _
                                lngCAPTADORID, _
                                lngEMPRESAID, _
                                mskMatricula.ClipText, _
                                txtContrato.Text, _
                                mskDtInicio.Text, _
                                IIf(mskDtFim.ClipText = "", "", mskDtFim.Text), _
                                txtRuaCom.Text, _
                                txtNumeroCom.Text, _
                                txtComplementoCom.Text, _
                                IIf(mskCepCom.ClipText = "", "", mskCepCom.ClipText), _
                                txtBairroCom.Text, _
                                txtCidadeCom.Text, _
                                txtEstadoCom.Text, _
                                txtTelefoneCom1.Text, _
                                txtTelefoneCom2.Text, _
                                txtRuaCob.Text, _
                                txtNumeroCob.Text, _
                                txtComplementoCob.Text, _
                                IIf(mskCepCob.ClipText = "", "", mskCepCob.ClipText), _
                                txtBairroCob.Text, _
                                txtCidadeCob.Text, _
                                txtEstadoCob.Text
      Set objTitular = Nothing
    Else
      'Dependente
      Set objDependente = New busApler.clsAssociadoDependente
      objDependente.InserirDependente lngPKID, _
                                      lngGRAUPARENTESCOID, _
                                      lngASSOCIADOTITULARID, _
                                      mskMatriculaDep.ClipText
      Set objDependente = Nothing
    End If
    '
    'Selecionar associado pelo nome
    Status = tpStatus_Alterar
    'Reload na tela
    Form_Load
    'Acerta tabs
    tabDetalhes.Tab = 4
    blnRetorno = True
  End If
  Set objAssociado = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  cmdOk.Enabled = True
End Sub


Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCampos = False
  If Not Valida_String(txtNome, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o nome" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(Trim(mskCpf.ClipText)) = 0 Then
    strMsg = strMsg & "Informar o CPF" & vbCrLf
    Pintar_Controle mskCpf, tpCorContr_Erro
    SetarFoco mskCpf
    tabDetalhes.Tab = 0
    blnSetarFocoControle = False
  End If
  If Len(Trim(mskCpf.ClipText)) > 0 Then
    If Not TestaCPF(mskCpf.ClipText) Then
      strMsg = strMsg & "Informar o CPF válido" & vbCrLf
      Pintar_Controle mskCpf, tpCorContr_Erro
      SetarFoco mskCpf
      tabDetalhes.Tab = 0
      blnSetarFocoControle = False
    End If
  End If
  If Not Valida_Option(optSexo, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o sexo" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Data(mskDtNascimento, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de nascimento válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(cboTipoSocio, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Slecionar o tipo do sócio" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(cboValorPlano, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Slecionar o plano" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  'Dependente
  If strIcAssociado = "D" Then
    'Dependente
    If Not Valida_Moeda(mskMatriculaDep, TpNaoObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Preencher a matricula do dependente válida" & vbCrLf
      tabDetalhes.Tab = 0
    End If
    If Not Valida_String(cboGrauParentesco, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Slecionar o grau de parentesco" & vbCrLf
      tabDetalhes.Tab = 0
    End If
  End If
  'Endereço residencial
  If Len(Trim(mskCepRes.ClipText)) > 0 Then
    If Len(Trim(mskCepRes.ClipText)) <> 8 Then
      strMsg = strMsg & "Informar o CEP residencial válido" & vbCrLf
      Pintar_Controle mskCepRes, tpCorContr_Erro
      SetarFoco mskCepRes
      tabDetalhes.Tab = 1
      blnSetarFocoControle = False
    End If
  End If
  'Titulat
  If strIcAssociado = "T" Then
    'Titular
    If Not Valida_Moeda(mskMatricula, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Preencher a matricula válida" & vbCrLf
      tabDetalhes.Tab = 2
    End If
    If Not Valida_String(txtContrato, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Preencher o contrato válido" & vbCrLf
      tabDetalhes.Tab = 2
    End If
    If Not Valida_Data(mskDtInicio, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Preencher a data início do contrato válida" & vbCrLf
      tabDetalhes.Tab = 2
    End If
    If Not Valida_Data(mskDtFim, TpNaoObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Preencher a data fim do contrato válida" & vbCrLf
      tabDetalhes.Tab = 2
    End If
    If Not Valida_String(cboEmpresa, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Slecionar a empresa" & vbCrLf
      tabDetalhes.Tab = 2
    End If
    'End cob
    If Len(Trim(mskCepCob.ClipText)) > 0 Then
      If Len(Trim(mskCepCob.ClipText)) <> 8 Then
        strMsg = strMsg & "Informar o CEP de cobrança válido" & vbCrLf
        Pintar_Controle mskCepCob, tpCorContr_Erro
        SetarFoco mskCepCob
        tabDetalhes.Tab = 3
        blnSetarFocoControle = False
      End If
    End If
    If Len(Trim(txtEstadoCob.Text)) <> 0 Then
      If Len(Trim(txtEstadoCob.Text)) <> 2 Then
        strMsg = strMsg & "Informar o estado do endereço de cobrança com duas posições" & vbCrLf
        Pintar_Controle txtEstadoCob, tpCorContr_Erro
        'SetarFoco txtEstadoCob
        tabDetalhes.Tab = 3
        blnSetarFocoControle = False
      End If
    End If
    'End com
    If Len(Trim(mskCepCom.ClipText)) > 0 Then
      If Len(Trim(mskCepCom.ClipText)) <> 8 Then
        strMsg = strMsg & "Informar o CEP comercial válido" & vbCrLf
        Pintar_Controle mskCepCom, tpCorContr_Erro
        SetarFoco mskCepCom
        tabDetalhes.Tab = 3
        blnSetarFocoControle = False
      End If
    End If
    If Len(Trim(txtEstadoCom.Text)) <> 0 Then
      If Len(Trim(txtEstadoCom.Text)) <> 2 Then
        strMsg = strMsg & "Informar o estado do endereço comercial com duas posições" & vbCrLf
        Pintar_Controle txtEstadoCom, tpCorContr_Erro
        'SetarFoco txtEstadoCom
        tabDetalhes.Tab = 3
        blnSetarFocoControle = False
      End If
    End If
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserAssociadoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserAssociadoInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco txtNome
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserAssociadoInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objAssociado            As busApler.clsAssociado
  Dim objTitular              As busApler.clsAssociadoTitular
  Dim objDependente           As busApler.clsAssociadoDependente
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 6660
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, cmdExcluir, , cmdIncluir, cmdAlterar
  '
  'Limpar Campos
  LimparCampos
  'Tratar campos
  TratarCampos
  
  'Empresa
  strSql = "Select NOME from EMPRESA ORDER BY NOME"
  PreencheCombo cboEmpresa, strSql, False, True
  'Tipo de Sócio
  strSql = "Select DESCRICAO from TIPOSOCIO ORDER BY DESCRICAO"
  PreencheCombo cboTipoSocio, strSql, False, True
  'Valor do Plano
  strSql = "Select DESCRICAO from VALORPLANO ORDER BY DESCRICAO"
  PreencheCombo cboValorPlano, strSql, False, True
  
  'Grau Parentesco
  strSql = "Select DESCRICAO from GRAUPARENTESCO ORDER BY DESCRICAO"
  PreencheCombo cboGrauParentesco, strSql, False, True
  'Origem
  strSql = "Select DESCRICAO from ORIGEM ORDER BY DESCRICAO"
  PreencheCombo cboOrigem, strSql, False, True
  'Captador
  strSql = "Select NOME from CAPTADOR ORDER BY NOME"
  PreencheCombo cboCaptador, strSql, False, True
  'Estado cívil
  strSql = "Select DESCRICAO from ESTADOCIVIL ORDER BY DESCRICAO"
  PreencheCombo cboEstadoCivil, strSql, False, True
  tabDetalhes_Click 1
  If Status = tpStatus_Incluir Then
    '
    optExcluido(0).Value = True
    'Visible
    optExcluido(0).Visible = False
    optExcluido(1).Visible = False
    Label5(11).Visible = False
    '
    tabDetalhes.TabEnabled(4) = False
    tabDetalhes.TabEnabled(5) = False
    tabDetalhes.TabEnabled(6) = False
    tabDetalhes.TabEnabled(7) = False
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    '-----------------------------
    'ASSOCIADO
    '------------------------------
    Set objAssociado = New busApler.clsAssociado
    Set objRs = objAssociado.SelecionarAssociadoPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      'Associado
      'Dados cadastrais
      txtNome.Text = objRs.Fields("NOME").Value & ""
      INCLUIR_VALOR_NO_MASK mskCpf, objRs.Fields("CPF").Value, TpMaskSemMascara
      If objRs.Fields("SEXO").Value & "" = "M" Then
        optSexo(0).Value = True
        optSexo(1).Value = False
      ElseIf objRs.Fields("SEXO").Value & "" = "F" Then
        optSexo(0).Value = False
        optSexo(1).Value = True
      Else
        optSexo(0).Value = False
        optSexo(1).Value = False
      End If
      INCLUIR_VALOR_NO_MASK mskDtNascimento, objRs.Fields("DATANASCIMENTO").Value, TpMaskData
      
      txtIdentidade.Text = objRs.Fields("IDENTIDADE").Value & ""
      txtOrgaoEmissor.Text = objRs.Fields("ORGEMISS").Value & ""
      If objRs.Fields("DESCR_ESTADOCIVIL").Value & "" <> "" Then
        cboEstadoCivil.Text = objRs.Fields("DESCR_ESTADOCIVIL").Value & ""
      End If
      txtNaturalidade.Text = objRs.Fields("NATURALIDADE").Value & ""
      txtEspecial.Text = objRs.Fields("ESPECIAL").Value & ""
      If objRs.Fields("DESCR_TIPOSOCIO").Value & "" <> "" Then
        cboTipoSocio.Text = objRs.Fields("DESCR_TIPOSOCIO").Value & ""
      End If
      If objRs.Fields("DESCR_VALORPLANO").Value & "" <> "" Then
        cboValorPlano.Text = objRs.Fields("DESCR_VALORPLANO").Value & ""
      End If
      txtNomeMae.Text = objRs.Fields("NOMEMAE").Value & ""
      txtTelefoneRes.Text = objRs.Fields("TELEFONERES1").Value & ""
      txtCelular.Text = objRs.Fields("CELULAR").Value & ""
      txtEmail.Text = objRs.Fields("EMAIL").Value & ""
      txtObservacao.Text = objRs.Fields("OBSERVACAO").Value & ""
      If objRs.Fields("EXCLUIDO").Value & "" = "S" Then
        optExcluido(0).Value = True
        optExcluido(1).Value = False
      ElseIf objRs.Fields("EXCLUIDO").Value & "" = "N" Then
        optExcluido(0).Value = False
        optExcluido(1).Value = True
      Else
        optExcluido(0).Value = False
        optExcluido(1).Value = False
      End If
      '
      'Endereço res
      txtRuaRes.Text = objRs.Fields("ENDRUARES").Value & ""
      txtNumeroRes.Text = objRs.Fields("ENDNUMERORES").Value & ""
      txtComplementoRes.Text = objRs.Fields("ENDCOMPLRES").Value & ""
      txtEstadoRes.Text = objRs.Fields("ENDESTADORES").Value & ""
      INCLUIR_VALOR_NO_MASK mskCepRes, objRs.Fields("ENDCEPRES").Value, TpMaskSemMascara
      txtBairroRes.Text = objRs.Fields("ENDBAIRRORES").Value & ""
      txtCidadeRes.Text = objRs.Fields("ENDCIDADERES").Value & ""
    End If
    objRs.Close
    Set objRs = Nothing
    Set objAssociado = Nothing
    If strIcAssociado = "D" Then
      '-----------------------------
      'DEPENDENTE
      '------------------------------
      Set objDependente = New busApler.clsAssociadoDependente
      Set objRs = objDependente.SelecionarDependentePeloPkid(lngPKID)
      '
      If Not objRs.EOF Then
        'DEPENDENTE
        INCLUIR_VALOR_NO_MASK mskMatriculaDep, objRs.Fields("MATRICULADEP").Value, TpMaskLongo
        If objRs.Fields("DESCR_GRAUPARENTESCO").Value & "" <> "" Then
          cboGrauParentesco.Text = objRs.Fields("DESCR_GRAUPARENTESCO").Value & ""
        End If
        '
      End If
      objRs.Close
      Set objRs = Nothing
      Set objDependente = Nothing
    Else
      '-----------------------------
      'TITULAR
      '------------------------------
      Set objTitular = New busApler.clsAssociadoTitular
      Set objRs = objTitular.SelecionarTitularPeloPkid(lngPKID)
      '
      If Not objRs.EOF Then
        'Titular
        INCLUIR_VALOR_NO_MASK mskMatricula, objRs.Fields("MATRICULA").Value, TpMaskLongo
        txtContrato.Text = objRs.Fields("NUMEROCONTRATO").Value & ""
        INCLUIR_VALOR_NO_MASK mskDtInicio, objRs.Fields("DATAINICONTRATO").Value, TpMaskData
        INCLUIR_VALOR_NO_MASK mskDtFim, objRs.Fields("DATAFIMCONTRATO").Value, TpMaskData
        If objRs.Fields("DESCR_ORIGEM").Value & "" <> "" Then
          cboOrigem.Text = objRs.Fields("DESCR_ORIGEM").Value & ""
        End If
        If objRs.Fields("DESCR_CAPTADOR").Value & "" <> "" Then
          cboCaptador.Text = objRs.Fields("DESCR_CAPTADOR").Value & ""
        End If
        If objRs.Fields("DESCR_EMPRESA").Value & "" <> "" Then
          cboEmpresa.Text = objRs.Fields("DESCR_EMPRESA").Value & ""
        End If
        'Endereço Cobrança
        txtRuaCob.Text = objRs.Fields("ENDRUACOB").Value & ""
        txtNumeroCob.Text = objRs.Fields("ENDNUMEROCOB").Value & ""
        txtComplementoCob.Text = objRs.Fields("ENDCOMPLCOB").Value & ""
        txtEstadoCob.Text = objRs.Fields("ENDESTADOCOB").Value & ""
        INCLUIR_VALOR_NO_MASK mskCepCob, objRs.Fields("ENDCEPCOB").Value, TpMaskSemMascara
        txtBairroCob.Text = objRs.Fields("ENDBAIRROCOB").Value & ""
        txtCidadeCob.Text = objRs.Fields("ENDCIDADECOB").Value & ""
        'Endereço Comercial
        txtRuaCom.Text = objRs.Fields("ENDRUACOM").Value & ""
        txtNumeroCom.Text = objRs.Fields("ENDNUMEROCOM").Value & ""
        txtComplementoCom.Text = objRs.Fields("ENDCOMPLCOM").Value & ""
        txtEstadoCom.Text = objRs.Fields("ENDESTADOCOM").Value & ""
        INCLUIR_VALOR_NO_MASK mskCepCom, objRs.Fields("ENDCEPCOM").Value, TpMaskSemMascara
        txtBairroCom.Text = objRs.Fields("ENDBAIRROCOM").Value & ""
        txtCidadeCom.Text = objRs.Fields("ENDCIDADECOM").Value & ""
        txtTelefoneCom1.Text = objRs.Fields("TELEFONECOM1").Value & ""
        txtTelefoneCom2.Text = objRs.Fields("TELEFONECOM2").Value & ""
        '
      End If
      objRs.Close
      Set objRs = Nothing
      Set objTitular = Nothing
    End If
    'Visible
    optExcluido(0).Visible = True
    optExcluido(1).Visible = True
    Label5(11).Visible = True
    '
    tabDetalhes.TabEnabled(4) = True
    tabDetalhes.TabEnabled(5) = True
    tabDetalhes.TabEnabled(6) = True
    tabDetalhes.TabEnabled(7) = True
  End If
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then Cancel = True
End Sub



Private Sub grdConvenio_UnboundReadDataEx( _
     ByVal RowBuf As TrueDBGrid60.RowBuffer, _
    StartLocation As Variant, ByVal Offset As Long, _
    ApproximatePosition As Long)
  ' UnboundReadData is fired by an unbound grid whenever
  ' it requires data for display. This event will fire
  ' when the grid is first shown, when Refresh or ReBind
  ' is used, when the grid is scrolled, and after a
  ' record in the grid is modified and the user commits
  ' the change by moving off of the current row. The
  ' grid fetches data in "chunks", and the number of rows
  ' the grid is asking for is given by RowBuf.RowCount.
  ' RowBuf is the row buffer where you place the data
  ' the bookmarks for the rows that the grid is
  ' requesting to display. It will also hold the number
  ' of rows that were successfully supplied to the grid.
  ' StartLocation is a vrtBookmark which, together with
  ' Offset, specifies the row for the programmer to start
  ' transferring data. A StartLocation of Null indicates
  ' a request for data from BOF or EOF.
  ' Offset specifies the relative position (from
  ' StartLocation) of the row for the programmer to start
  ' transferring data. A positive number indicates a
  ' forward relative position while a negative number
  ' indicates a backward relative position. Regardless
  ' of whether the rows to be read are before or after
  ' StartLocation, rows are always fetched going forward
  ' (this is why there is no ReadPriorRows parameter to
  ' the procedure).
  ' If you page down on the grid, for instance, the new
  ' top row of the grid will have an index greater than
  ' the StartLocation (Offset > 0). If you page up on
  ' the grid, the new index is less than that of
  ' StartLocation, so Offset < 0. If StartLocation is
  ' a vrtBookmark to row N, the grid always asks for row
  ' data in the following order:
  '   (N + Offset), (N + Offset + 1), (N + Offset + 2)...
  ' ApproximatePosition is a value you can set to indicate
  ' the ordinal position of (StartLocation + Offset).
  ' Setting this variable will enhance the ability of the
  ' grid to display its vertical scroll bar accurately.
  ' If the exact ordinal position of the new location is
  ' not known, you can set it to a reasonable,
  ' approximate value, or just ignore this parameter.
    
  On Error GoTo trata
  '
  Dim intColIndex      As Integer
  Dim intJ             As Integer
  Dim intRowsFetched   As Integer
  Dim intI             As Long
  Dim lngNewPosition   As Long
  Dim vrtBookmark      As Variant
  '
  intRowsFetched = 0
  For intI = 0 To RowBuf.RowCount - 1
    ' Get the vrtBookmark of the next available row
    vrtBookmark = GetRelativeBookmarkGeral(StartLocation, _
               Offset + intI, CONV_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, CONV_COLUNASMATRIZ, CONV_LINHASMATRIZ, CONV_Matriz)
    Next intJ
  
    ' Set the vrtBookmark for the row
    RowBuf.Bookmark(intI) = vrtBookmark
  
    ' Increment the count of fetched rows
    intRowsFetched = intRowsFetched + 1
  Next intI
  
' Tell the grid how many rows were fetched
  RowBuf.RowCount = intRowsFetched
  
' Set the approximate scroll bar position. Only
' nonnegative values of IndexFromBookmark() are valid.
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, CONV_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserAssociadoInc.grdGeral_UnboundReadDataEx]"
End Sub

Private Sub grdDependente_UnboundReadDataEx( _
     ByVal RowBuf As TrueDBGrid60.RowBuffer, _
    StartLocation As Variant, ByVal Offset As Long, _
    ApproximatePosition As Long)
  ' UnboundReadData is fired by an unbound grid whenever
  ' it requires data for display. This event will fire
  ' when the grid is first shown, when Refresh or ReBind
  ' is used, when the grid is scrolled, and after a
  ' record in the grid is modified and the user commits
  ' the change by moving off of the current row. The
  ' grid fetches data in "chunks", and the number of rows
  ' the grid is asking for is given by RowBuf.RowCount.
  ' RowBuf is the row buffer where you place the data
  ' the bookmarks for the rows that the grid is
  ' requesting to display. It will also hold the number
  ' of rows that were successfully supplied to the grid.
  ' StartLocation is a vrtBookmark which, together with
  ' Offset, specifies the row for the programmer to start
  ' transferring data. A StartLocation of Null indicates
  ' a request for data from BOF or EOF.
  ' Offset specifies the relative position (from
  ' StartLocation) of the row for the programmer to start
  ' transferring data. A positive number indicates a
  ' forward relative position while a negative number
  ' indicates a backward relative position. Regardless
  ' of whether the rows to be read are before or after
  ' StartLocation, rows are always fetched going forward
  ' (this is why there is no ReadPriorRows parameter to
  ' the procedure).
  ' If you page down on the grid, for instance, the new
  ' top row of the grid will have an index greater than
  ' the StartLocation (Offset > 0). If you page up on
  ' the grid, the new index is less than that of
  ' StartLocation, so Offset < 0. If StartLocation is
  ' a vrtBookmark to row N, the grid always asks for row
  ' data in the following order:
  '   (N + Offset), (N + Offset + 1), (N + Offset + 2)...
  ' ApproximatePosition is a value you can set to indicate
  ' the ordinal position of (StartLocation + Offset).
  ' Setting this variable will enhance the ability of the
  ' grid to display its vertical scroll bar accurately.
  ' If the exact ordinal position of the new location is
  ' not known, you can set it to a reasonable,
  ' approximate value, or just ignore this parameter.
    
  On Error GoTo trata
  '
  Dim intColIndex      As Integer
  Dim intJ             As Integer
  Dim intRowsFetched   As Integer
  Dim intI             As Long
  Dim lngNewPosition   As Long
  Dim vrtBookmark      As Variant
  '
  intRowsFetched = 0
  For intI = 0 To RowBuf.RowCount - 1
    ' Get the vrtBookmark of the next available row
    vrtBookmark = GetRelativeBookmarkGeral(StartLocation, _
               Offset + intI, DEP_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, DEP_COLUNASMATRIZ, DEP_LINHASMATRIZ, DEP_Matriz)
    Next intJ
  
    ' Set the vrtBookmark for the row
    RowBuf.Bookmark(intI) = vrtBookmark
  
    ' Increment the count of fetched rows
    intRowsFetched = intRowsFetched + 1
  Next intI
  
' Tell the grid how many rows were fetched
  RowBuf.RowCount = intRowsFetched
  
' Set the approximate scroll bar position. Only
' nonnegative values of IndexFromBookmark() are valid.
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, DEP_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserAssociadoInc.grdGeral_UnboundReadDataEx]"
End Sub

Private Sub grdLinha_UnboundReadDataEx( _
     ByVal RowBuf As TrueDBGrid60.RowBuffer, _
    StartLocation As Variant, ByVal Offset As Long, _
    ApproximatePosition As Long)
  ' UnboundReadData is fired by an unbound grid whenever
  ' it requires data for display. This event will fire
  ' when the grid is first shown, when Refresh or ReBind
  ' is used, when the grid is scrolled, and after a
  ' record in the grid is modified and the user commits
  ' the change by moving off of the current row. The
  ' grid fetches data in "chunks", and the number of rows
  ' the grid is asking for is given by RowBuf.RowCount.
  ' RowBuf is the row buffer where you place the data
  ' the bookmarks for the rows that the grid is
  ' requesting to display. It will also hold the number
  ' of rows that were successfully supplied to the grid.
  ' StartLocation is a vrtBookmark which, together with
  ' Offset, specifies the row for the programmer to start
  ' transferring data. A StartLocation of Null indicates
  ' a request for data from BOF or EOF.
  ' Offset specifies the relative position (from
  ' StartLocation) of the row for the programmer to start
  ' transferring data. A positive number indicates a
  ' forward relative position while a negative number
  ' indicates a backward relative position. Regardless
  ' of whether the rows to be read are before or after
  ' StartLocation, rows are always fetched going forward
  ' (this is why there is no ReadPriorRows parameter to
  ' the procedure).
  ' If you page down on the grid, for instance, the new
  ' top row of the grid will have an index greater than
  ' the StartLocation (Offset > 0). If you page up on
  ' the grid, the new index is less than that of
  ' StartLocation, so Offset < 0. If StartLocation is
  ' a vrtBookmark to row N, the grid always asks for row
  ' data in the following order:
  '   (N + Offset), (N + Offset + 1), (N + Offset + 2)...
  ' ApproximatePosition is a value you can set to indicate
  ' the ordinal position of (StartLocation + Offset).
  ' Setting this variable will enhance the ability of the
  ' grid to display its vertical scroll bar accurately.
  ' If the exact ordinal position of the new location is
  ' not known, you can set it to a reasonable,
  ' approximate value, or just ignore this parameter.
    
  On Error GoTo trata
  '
  Dim intColIndex      As Integer
  Dim intJ             As Integer
  Dim intRowsFetched   As Integer
  Dim intI             As Long
  Dim lngNewPosition   As Long
  Dim vrtBookmark      As Variant
  '
  intRowsFetched = 0
  For intI = 0 To RowBuf.RowCount - 1
    ' Get the vrtBookmark of the next available row
    vrtBookmark = GetRelativeBookmarkGeral(StartLocation, _
               Offset + intI, LIN_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, LIN_COLUNASMATRIZ, LIN_LINHASMATRIZ, LIN_Matriz)
    Next intJ
  
    ' Set the vrtBookmark for the row
    RowBuf.Bookmark(intI) = vrtBookmark
  
    ' Increment the count of fetched rows
    intRowsFetched = intRowsFetched + 1
  Next intI
  
' Tell the grid how many rows were fetched
  RowBuf.RowCount = intRowsFetched
  
' Set the approximate scroll bar position. Only
' nonnegative values of IndexFromBookmark() are valid.
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, LIN_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserAssociadoInc.grdGeral_UnboundReadDataEx]"
End Sub

Private Sub grdProfissao_UnboundReadDataEx( _
     ByVal RowBuf As TrueDBGrid60.RowBuffer, _
    StartLocation As Variant, ByVal Offset As Long, _
    ApproximatePosition As Long)
  ' UnboundReadData is fired by an unbound grid whenever
  ' it requires data for display. This event will fire
  ' when the grid is first shown, when Refresh or ReBind
  ' is used, when the grid is scrolled, and after a
  ' record in the grid is modified and the user commits
  ' the change by moving off of the current row. The
  ' grid fetches data in "chunks", and the number of rows
  ' the grid is asking for is given by RowBuf.RowCount.
  ' RowBuf is the row buffer where you place the data
  ' the bookmarks for the rows that the grid is
  ' requesting to display. It will also hold the number
  ' of rows that were successfully supplied to the grid.
  ' StartLocation is a vrtBookmark which, together with
  ' Offset, specifies the row for the programmer to start
  ' transferring data. A StartLocation of Null indicates
  ' a request for data from BOF or EOF.
  ' Offset specifies the relative position (from
  ' StartLocation) of the row for the programmer to start
  ' transferring data. A positive number indicates a
  ' forward relative position while a negative number
  ' indicates a backward relative position. Regardless
  ' of whether the rows to be read are before or after
  ' StartLocation, rows are always fetched going forward
  ' (this is why there is no ReadPriorRows parameter to
  ' the procedure).
  ' If you page down on the grid, for instance, the new
  ' top row of the grid will have an index greater than
  ' the StartLocation (Offset > 0). If you page up on
  ' the grid, the new index is less than that of
  ' StartLocation, so Offset < 0. If StartLocation is
  ' a vrtBookmark to row N, the grid always asks for row
  ' data in the following order:
  '   (N + Offset), (N + Offset + 1), (N + Offset + 2)...
  ' ApproximatePosition is a value you can set to indicate
  ' the ordinal position of (StartLocation + Offset).
  ' Setting this variable will enhance the ability of the
  ' grid to display its vertical scroll bar accurately.
  ' If the exact ordinal position of the new location is
  ' not known, you can set it to a reasonable,
  ' approximate value, or just ignore this parameter.
    
  On Error GoTo trata
  '
  Dim intColIndex      As Integer
  Dim intJ             As Integer
  Dim intRowsFetched   As Integer
  Dim intI             As Long
  Dim lngNewPosition   As Long
  Dim vrtBookmark      As Variant
  '
  intRowsFetched = 0
  For intI = 0 To RowBuf.RowCount - 1
    ' Get the vrtBookmark of the next available row
    vrtBookmark = GetRelativeBookmarkGeral(StartLocation, _
               Offset + intI, PROF_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, PROF_COLUNASMATRIZ, PROF_LINHASMATRIZ, PROF_Matriz)
    Next intJ
  
    ' Set the vrtBookmark for the row
    RowBuf.Bookmark(intI) = vrtBookmark
  
    ' Increment the count of fetched rows
    intRowsFetched = intRowsFetched + 1
  Next intI
  
' Tell the grid how many rows were fetched
  RowBuf.RowCount = intRowsFetched
  
' Set the approximate scroll bar position. Only
' nonnegative values of IndexFromBookmark() are valid.
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, PROF_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserAssociadoInc.grdGeral_UnboundReadDataEx]"
End Sub

Private Sub mskCepCob_GotFocus()
  Seleciona_Conteudo_Controle mskCepCob
End Sub
Private Sub mskCepCob_LostFocus()
  Pintar_Controle mskCepCob, tpCorContr_Normal
End Sub

Private Sub mskCepCom_GotFocus()
  Seleciona_Conteudo_Controle mskCepCom
End Sub
Private Sub mskCepCom_LostFocus()
  Pintar_Controle mskCepCom, tpCorContr_Normal
End Sub

Private Sub mskCepRes_GotFocus()
  Seleciona_Conteudo_Controle mskCepRes
End Sub
Private Sub mskCepRes_LostFocus()
  Pintar_Controle mskCepRes, tpCorContr_Normal
End Sub

Private Sub mskCpf_GotFocus()
  Seleciona_Conteudo_Controle mskCpf
End Sub
Private Sub mskCpf_LostFocus()
  Pintar_Controle mskCpf, tpCorContr_Normal
End Sub

Private Sub mskDtFim_GotFocus()
  Seleciona_Conteudo_Controle mskDtFim
End Sub
Private Sub mskDtFim_LostFocus()
  Pintar_Controle mskDtFim, tpCorContr_Normal
End Sub
Private Sub mskDtInicio_GotFocus()
  Seleciona_Conteudo_Controle mskDtInicio
End Sub
Private Sub mskDtInicio_LostFocus()
  Pintar_Controle mskDtInicio, tpCorContr_Normal
End Sub

Private Sub mskDtNascimento_GotFocus()
  Seleciona_Conteudo_Controle mskDtNascimento
End Sub
Private Sub mskDtNascimento_LostFocus()
  Pintar_Controle mskDtNascimento, tpCorContr_Normal
End Sub

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    'Dados cadastrais
    grdProfissao.Enabled = False
    grdConvenio.Enabled = False
    grdDependente.Enabled = False
    grdLinha.Enabled = False
    pictrava(0).Enabled = True
    pictrava(1).Enabled = True
    pictrava(2).Enabled = False
    pictrava(3).Enabled = False
    pictrava(4).Enabled = False
    pictrava(5).Enabled = False
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    SetarFoco txtNome
  Case 1
    'Endereço Residencial
    grdProfissao.Enabled = False
    grdConvenio.Enabled = False
    grdDependente.Enabled = False
    grdLinha.Enabled = False
    pictrava(0).Enabled = False
    pictrava(1).Enabled = False
    pictrava(2).Enabled = True
    pictrava(3).Enabled = False
    pictrava(4).Enabled = False
    pictrava(5).Enabled = False
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    SetarFoco txtRuaRes
  Case 2
    'Titular
    grdProfissao.Enabled = False
    grdConvenio.Enabled = False
    grdDependente.Enabled = False
    grdLinha.Enabled = False
    pictrava(0).Enabled = False
    pictrava(1).Enabled = False
    pictrava(2).Enabled = False
    pictrava(3).Enabled = True
    pictrava(4).Enabled = False
    pictrava(5).Enabled = False
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    SetarFoco mskMatricula
  Case 3
    'Endereço com/cob
    grdProfissao.Enabled = False
    grdConvenio.Enabled = False
    grdDependente.Enabled = False
    grdLinha.Enabled = False
    pictrava(0).Enabled = False
    pictrava(1).Enabled = False
    pictrava(2).Enabled = False
    pictrava(3).Enabled = False
    pictrava(4).Enabled = True
    pictrava(5).Enabled = True
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    SetarFoco txtRuaCob
  Case 4
    'Profissão
    grdProfissao.Enabled = True
    grdConvenio.Enabled = False
    grdDependente.Enabled = False
    grdLinha.Enabled = False
    pictrava(0).Enabled = False
    pictrava(1).Enabled = False
    pictrava(2).Enabled = False
    pictrava(3).Enabled = False
    pictrava(4).Enabled = False
    pictrava(5).Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = False
    '
    'Montar RecordSet
    PROF_COLUNASMATRIZ = grdProfissao.Columns.Count
    PROF_LINHASMATRIZ = 0
    PROF_MontaMatriz
    grdProfissao.Bookmark = Null
    grdProfissao.ReBind
    grdProfissao.ApproxCount = PROF_LINHASMATRIZ
    '
    SetarFoco grdProfissao
  Case 5
    'Convênios
    grdProfissao.Enabled = False
    grdConvenio.Enabled = True
    grdDependente.Enabled = False
    grdLinha.Enabled = False
    pictrava(0).Enabled = False
    pictrava(1).Enabled = False
    pictrava(2).Enabled = False
    pictrava(3).Enabled = False
    pictrava(4).Enabled = False
    pictrava(5).Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = True
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = True
    '
    'Montar RecordSet
    CONV_COLUNASMATRIZ = grdConvenio.Columns.Count
    CONV_LINHASMATRIZ = 0
    CONV_MontaMatriz
    grdConvenio.Bookmark = Null
    grdConvenio.ReBind
    grdConvenio.ApproxCount = CONV_LINHASMATRIZ
    '
    SetarFoco grdConvenio
  Case 6
    'Dependentes
    grdProfissao.Enabled = False
    grdConvenio.Enabled = False
    grdDependente.Enabled = True
    grdLinha.Enabled = False
    pictrava(0).Enabled = False
    pictrava(1).Enabled = False
    pictrava(2).Enabled = False
    pictrava(3).Enabled = False
    pictrava(4).Enabled = False
    pictrava(5).Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = True
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = True
    '
    'Montar RecordSet
    DEP_COLUNASMATRIZ = grdDependente.Columns.Count
    DEP_LINHASMATRIZ = 0
    DEP_MontaMatriz
    grdDependente.Bookmark = Null
    grdDependente.ReBind
    grdDependente.ApproxCount = DEP_LINHASMATRIZ
    '
    SetarFoco grdDependente
  Case 7
    'Linha
    grdProfissao.Enabled = False
    grdConvenio.Enabled = False
    grdDependente.Enabled = False
    grdLinha.Enabled = True
    pictrava(0).Enabled = False
    pictrava(1).Enabled = False
    pictrava(2).Enabled = False
    pictrava(3).Enabled = False
    pictrava(4).Enabled = False
    pictrava(5).Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = False
    '
    'Montar RecordSet
    LIN_COLUNASMATRIZ = grdLinha.Columns.Count
    LIN_LINHASMATRIZ = 0
    LIN_MontaMatriz
    grdLinha.Bookmark = Null
    grdLinha.ReBind
    grdLinha.ApproxCount = LIN_LINHASMATRIZ
    '
    SetarFoco grdLinha
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "Apler.frmUserPlanoInc.tabDetalhes"
  AmpN
End Sub


Private Sub txtBairroCob_GotFocus()
  Seleciona_Conteudo_Controle txtBairroCob
End Sub
Private Sub txtBairroCob_LostFocus()
  Pintar_Controle txtBairroCob, tpCorContr_Normal
End Sub

Private Sub txtBairroCom_GotFocus()
  Seleciona_Conteudo_Controle txtBairroCom
End Sub
Private Sub txtBairroCom_LostFocus()
  Pintar_Controle txtBairroCom, tpCorContr_Normal
End Sub

Private Sub txtBairroRes_GotFocus()
  Seleciona_Conteudo_Controle txtBairroRes
End Sub
Private Sub txtBairroRes_LostFocus()
  Pintar_Controle txtBairroRes, tpCorContr_Normal
End Sub

Private Sub txtCelular_GotFocus()
  Seleciona_Conteudo_Controle txtCelular
End Sub
Private Sub txtCelular_LostFocus()
  Pintar_Controle txtCelular, tpCorContr_Normal
End Sub

Private Sub txtCidadeCob_GotFocus()
  Seleciona_Conteudo_Controle txtCidadeCob
End Sub
Private Sub txtCidadeCob_LostFocus()
  Pintar_Controle txtCidadeCob, tpCorContr_Normal
End Sub

Private Sub txtCidadeCom_GotFocus()
  Seleciona_Conteudo_Controle txtCidadeCom
End Sub
Private Sub txtCidadeCom_LostFocus()
  Pintar_Controle txtCidadeCom, tpCorContr_Normal
End Sub

Private Sub txtCidadeRes_GotFocus()
  Seleciona_Conteudo_Controle txtCidadeRes
End Sub
Private Sub txtCidadeRes_LostFocus()
  Pintar_Controle txtCidadeRes, tpCorContr_Normal
End Sub

Private Sub txtComplementoCob_GotFocus()
  Seleciona_Conteudo_Controle txtComplementoCob
End Sub
Private Sub txtComplementoCob_LostFocus()
  Pintar_Controle txtComplementoCob, tpCorContr_Normal
End Sub

Private Sub txtComplementoCom_GotFocus()
  Seleciona_Conteudo_Controle txtComplementoCom
End Sub
Private Sub txtComplementoCom_LostFocus()
  Pintar_Controle txtComplementoCom, tpCorContr_Normal
End Sub

Private Sub txtComplementoRes_GotFocus()
  Seleciona_Conteudo_Controle txtComplementoRes
End Sub
Private Sub txtComplementoRes_LostFocus()
  Pintar_Controle txtComplementoRes, tpCorContr_Normal
End Sub

Private Sub txtContrato_GotFocus()
  Seleciona_Conteudo_Controle txtContrato
End Sub
Private Sub txtContrato_LostFocus()
  Pintar_Controle txtContrato, tpCorContr_Normal
End Sub

Private Sub txtEmail_GotFocus()
  Seleciona_Conteudo_Controle txtEmail
End Sub
Private Sub txtEmail_LostFocus()
  Pintar_Controle txtEmail, tpCorContr_Normal
End Sub

Private Sub mskMatricula_GotFocus()
  Seleciona_Conteudo_Controle mskMatricula
End Sub
Private Sub mskMatricula_LostFocus()
  Pintar_Controle mskMatricula, tpCorContr_Normal
End Sub

Private Sub mskMatriculaDep_GotFocus()
  Seleciona_Conteudo_Controle mskMatriculaDep
End Sub
Private Sub mskMatriculaDep_LostFocus()
  Pintar_Controle mskMatriculaDep, tpCorContr_Normal
End Sub

Private Sub txtEspecial_GotFocus()
  Seleciona_Conteudo_Controle txtEspecial
End Sub
Private Sub txtEspecial_LostFocus()
  Pintar_Controle txtEspecial, tpCorContr_Normal
End Sub

Private Sub txtEstadoCob_GotFocus()
  Seleciona_Conteudo_Controle txtEstadoCob
End Sub
Private Sub txtEstadoCob_LostFocus()
  Pintar_Controle txtEstadoCob, tpCorContr_Normal
End Sub

Private Sub txtEstadoCom_GotFocus()
  Seleciona_Conteudo_Controle txtEstadoCom
End Sub
Private Sub txtEstadoCom_LostFocus()
  Pintar_Controle txtEstadoCom, tpCorContr_Normal
End Sub

Private Sub txtEstadoRes_GotFocus()
  Seleciona_Conteudo_Controle txtEstadoRes
End Sub
Private Sub txtEstadoRes_LostFocus()
  Pintar_Controle txtEstadoRes, tpCorContr_Normal
End Sub

Private Sub txtIdentidade_GotFocus()
  Seleciona_Conteudo_Controle txtIdentidade
End Sub
Private Sub txtIdentidade_LostFocus()
  Pintar_Controle txtIdentidade, tpCorContr_Normal
End Sub

Private Sub txtNaturalidade_GotFocus()
  Seleciona_Conteudo_Controle txtNaturalidade
End Sub
Private Sub txtNaturalidade_LostFocus()
  Pintar_Controle txtNaturalidade, tpCorContr_Normal
End Sub

Private Sub txtNome_GotFocus()
  Seleciona_Conteudo_Controle txtNome
End Sub
Private Sub txtNome_LostFocus()
  Pintar_Controle txtNome, tpCorContr_Normal
End Sub


Public Sub PROF_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busApler.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busApler.clsGeral
  '
  strSql = "SELECT TAB_PROFASSOC.PKID, PROFISSAO.DESCRICAO " & _
          "FROM TAB_PROFASSOC INNER JOIN PROFISSAO ON PROFISSAO.PKID = TAB_PROFASSOC.PROFISSAOID " & _
          "WHERE TAB_PROFASSOC.ASSOCIADOID = " & Formata_Dados(lngPKID, tpDados_Longo) & _
          " ORDER BY profissao.DESCRICAO"

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    PROF_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim PROF_Matriz(0 To PROF_COLUNASMATRIZ - 1, 0 To PROF_LINHASMATRIZ - 1)
  Else
    ReDim PROF_Matriz(0 To PROF_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To PROF_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To PROF_COLUNASMATRIZ - 1  'varre as colunas
          PROF_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Public Sub LIN_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busApler.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busApler.clsGeral
  '
  strSql = "SELECT LINHA.PKID, LINHA.DESCRICAO " & _
          "FROM LINHA INNER JOIN TAB_TITLINHA ON LINHA.PKID = TAB_TITLINHA.LINHAID " & _
          "WHERE TAB_TITLINHA.TITULARASSOCIADOID = " & Formata_Dados(lngPKID, tpDados_Longo) & _
          " ORDER BY LINHA.DESCRICAO"

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    LIN_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim LIN_Matriz(0 To LIN_COLUNASMATRIZ - 1, 0 To LIN_LINHASMATRIZ - 1)
  Else
    ReDim LIN_Matriz(0 To LIN_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To LIN_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To LIN_COLUNASMATRIZ - 1  'varre as colunas
          LIN_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Public Sub CONV_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busApler.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busApler.clsGeral
  '
  strSql = "SELECT TAB_CONVASSOC.PKID, PLANOCONVENIO.NOME, TAB_CONVASSOC.DATAINICIO, TAB_CONVASSOC.DATATERMINO " & _
          "FROM TAB_CONVASSOC INNER JOIN PLANOCONVENIO ON PLANOCONVENIO.PKID = TAB_CONVASSOC.PLANOCONVENIOID " & _
          "WHERE TAB_CONVASSOC.ASSOCIADOID = " & Formata_Dados(lngPKID, tpDados_Longo) & _
          " ORDER BY PLANOCONVENIO.NOME"

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    CONV_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim CONV_Matriz(0 To CONV_COLUNASMATRIZ - 1, 0 To CONV_LINHASMATRIZ - 1)
  Else
    ReDim CONV_Matriz(0 To CONV_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To CONV_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To CONV_COLUNASMATRIZ - 1  'varre as colunas
          CONV_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Public Sub DEP_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busApler.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busApler.clsGeral
  '
  strSql = "SELECT ASSOCIADO.PKID, ASSOCIADO.NOME, ASSOCIADO.CPF, ASSOCIADO.DATANASCIMENTO, CASE ASSOCIADO.EXCLUIDO WHEN 'S' THEN 'Sim' ELSE 'Não' END " & _
          "FROM ASSOCIADO INNER JOIN DEPENDENTE ON ASSOCIADO.PKID = DEPENDENTE.ASSOCIADOID " & _
          "WHERE DEPENDENTE.TITULARASSOCIADOID = " & Formata_Dados(lngPKID, tpDados_Longo) & _
          " ORDER BY ASSOCIADO.NOME"

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    DEP_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim DEP_Matriz(0 To DEP_COLUNASMATRIZ - 1, 0 To DEP_LINHASMATRIZ - 1)
  Else
    ReDim DEP_Matriz(0 To DEP_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To DEP_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To DEP_COLUNASMATRIZ - 1  'varre as colunas
          DEP_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtNomeMae_GotFocus()
  Seleciona_Conteudo_Controle txtNomeMae
End Sub
Private Sub txtNomeMae_LostFocus()
  Pintar_Controle txtNomeMae, tpCorContr_Normal
End Sub

Private Sub txtNumeroCob_GotFocus()
  Seleciona_Conteudo_Controle txtNumeroCob
End Sub
Private Sub txtNumeroCob_LostFocus()
  Pintar_Controle txtNumeroCob, tpCorContr_Normal
End Sub

Private Sub txtNumeroCom_GotFocus()
  Seleciona_Conteudo_Controle txtNumeroCom
End Sub
Private Sub txtNumeroCom_LostFocus()
  Pintar_Controle txtNumeroCom, tpCorContr_Normal
End Sub

Private Sub txtNumeroRes_GotFocus()
  Seleciona_Conteudo_Controle txtNumeroRes
End Sub
Private Sub txtNumeroRes_LostFocus()
  Pintar_Controle txtNumeroRes, tpCorContr_Normal
End Sub

Private Sub txtObservacao_GotFocus()
  Seleciona_Conteudo_Controle txtObservacao
End Sub
Private Sub txtObservacao_LostFocus()
  Pintar_Controle txtObservacao, tpCorContr_Normal
End Sub

Private Sub txtOrgaoEmissor_GotFocus()
  Seleciona_Conteudo_Controle txtOrgaoEmissor
End Sub
Private Sub txtOrgaoEmissor_LostFocus()
  Pintar_Controle txtOrgaoEmissor, tpCorContr_Normal
End Sub

Private Sub txtRuaCob_GotFocus()
  Seleciona_Conteudo_Controle txtRuaCob
End Sub
Private Sub txtRuaCob_LostFocus()
  Pintar_Controle txtRuaCob, tpCorContr_Normal
End Sub

Private Sub txtRuaCom_GotFocus()
  Seleciona_Conteudo_Controle txtRuaCom
End Sub
Private Sub txtRuaCom_LostFocus()
  Pintar_Controle txtRuaCom, tpCorContr_Normal
End Sub

Private Sub txtRuaRes_GotFocus()
  Seleciona_Conteudo_Controle txtRuaRes
End Sub
Private Sub txtRuaRes_LostFocus()
  Pintar_Controle txtRuaRes, tpCorContr_Normal
End Sub

Private Sub txtTelefoneCom1_GotFocus()
  Seleciona_Conteudo_Controle txtTelefoneCom1
End Sub
Private Sub txtTelefoneCom1_LostFocus()
  Pintar_Controle txtTelefoneCom1, tpCorContr_Normal
End Sub

Private Sub txtTelefoneCom2_GotFocus()
  Seleciona_Conteudo_Controle txtTelefoneCom2
End Sub
Private Sub txtTelefoneCom2_LostFocus()
  Pintar_Controle txtTelefoneCom2, tpCorContr_Normal
End Sub

Private Sub txtTelefoneRes_GotFocus()
  Seleciona_Conteudo_Controle txtTelefoneRes
End Sub
Private Sub txtTelefoneRes_LostFocus()
  Pintar_Controle txtTelefoneRes, tpCorContr_Normal
End Sub

