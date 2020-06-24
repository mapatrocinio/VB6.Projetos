VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserContaCorrente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo de conta corrente"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   6630
      Left            =   9405
      ScaleHeight     =   6630
      ScaleWidth      =   1860
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   6465
         Left            =   120
         ScaleHeight     =   6405
         ScaleWidth      =   1605
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   30
         Width           =   1665
         Begin VB.CommandButton cmdParcela 
            Caption         =   "&U"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Y"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   2730
            Width           =   1335
         End
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "&X"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&V"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdCalculadora 
            Caption         =   "&Z"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   3600
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   5370
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   4485
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   6375
      Left            =   120
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   120
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do conta corrente"
      TabPicture(0)   =   "userContaCorrente.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame12"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame12 
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
         Height          =   5865
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   9015
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   5
            Left            =   240
            ScaleHeight     =   375
            ScaleWidth      =   3315
            TabIndex        =   74
            Top             =   1170
            Width           =   3315
            Begin VB.ComboBox cboTipoPagamento 
               Height          =   315
               Left            =   1260
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   0
               Width           =   2055
            End
            Begin VB.Label Label11 
               Caption         =   "Tipo Pgto."
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   75
               Top             =   30
               Width           =   975
            End
         End
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   5655
            Index           =   4
            Left            =   120
            ScaleHeight     =   5655
            ScaleWidth      =   8775
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   120
            Width           =   8775
            Begin VB.TextBox txtTrocoTotal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4590
               Locked          =   -1  'True
               TabIndex        =   69
               TabStop         =   0   'False
               Text            =   "txtTrocoTotal"
               Top             =   5250
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.CommandButton cmdImprimir 
               Caption         =   "ENTER"
               Height          =   800
               Left            =   7830
               Style           =   1  'Graphical
               TabIndex        =   73
               Top             =   720
               Width           =   800
            End
            Begin VB.TextBox txtRestante 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6990
               Locked          =   -1  'True
               TabIndex        =   70
               TabStop         =   0   'False
               Text            =   "txtRestante"
               Top             =   5280
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox txtLancado 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2220
               Locked          =   -1  'True
               TabIndex        =   68
               TabStop         =   0   'False
               Text            =   "txtLancado"
               Top             =   5250
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.PictureBox picPgtoFatura 
               BorderStyle     =   0  'None
               Height          =   345
               Left            =   60
               ScaleHeight     =   345
               ScaleWidth      =   8595
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   2670
               Width           =   8595
               Begin MSMask.MaskEdBox mskNroParcelas 
                  Height          =   255
                  Left            =   1230
                  TabIndex        =   22
                  Top             =   30
                  Width           =   525
                  _ExtentX        =   926
                  _ExtentY        =   450
                  _Version        =   393216
                  MaxLength       =   2
                  Format          =   "#,##0;($#,##0)"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskData 
                  Height          =   255
                  Index           =   0
                  Left            =   4470
                  TabIndex        =   23
                  Top             =   30
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   450
                  _Version        =   393216
                  AutoTab         =   -1  'True
                  MaxLength       =   10
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin VB.Label Label2 
                  Caption         =   "Nro. Parcelas"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   67
                  Top             =   0
                  Width           =   1125
               End
               Begin VB.Label Label1 
                  Caption         =   "Dt. Primeira Parcela"
                  Height          =   255
                  Left            =   2970
                  TabIndex        =   66
                  Top             =   30
                  Width           =   1425
               End
            End
            Begin VB.PictureBox picListaPgto 
               BorderStyle     =   0  'None
               Height          =   855
               Left            =   150
               ScaleHeight     =   855
               ScaleWidth      =   8505
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   1770
               Width           =   8505
               Begin TrueDBGrid60.TDBGrid grdGeral 
                  Height          =   750
                  Left            =   0
                  OleObjectBlob   =   "userContaCorrente.frx":001C
                  TabIndex        =   24
                  Top             =   0
                  Width           =   8310
               End
            End
            Begin VB.PictureBox picPgtoPenhor 
               BorderStyle     =   0  'None
               Height          =   615
               Left            =   30
               ScaleHeight     =   615
               ScaleWidth      =   8595
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   3120
               Width           =   8595
               Begin VB.TextBox txtDocumentoPenhor 
                  Height          =   285
                  Left            =   4470
                  MaxLength       =   30
                  TabIndex        =   12
                  Top             =   0
                  Width           =   1845
               End
               Begin VB.TextBox txtCliente 
                  Height          =   285
                  Left            =   1230
                  MaxLength       =   50
                  TabIndex        =   11
                  Top             =   0
                  Width           =   1845
               End
               Begin VB.TextBox txtObjeto 
                  Height          =   285
                  Left            =   1230
                  MaxLength       =   50
                  TabIndex        =   13
                  Top             =   330
                  Width           =   7095
               End
               Begin VB.Label Label62 
                  Caption         =   "Documento"
                  Height          =   255
                  Left            =   3390
                  TabIndex        =   63
                  Top             =   30
                  Width           =   825
               End
               Begin VB.Label Label58 
                  Caption         =   "Cliente"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   62
                  Top             =   0
                  Width           =   585
               End
               Begin VB.Label Label57 
                  Caption         =   "Desc. Objeto"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   61
                  Top             =   330
                  Width           =   1215
               End
            End
            Begin VB.PictureBox picPgtoCheque 
               BorderStyle     =   0  'None
               Height          =   615
               Left            =   30
               ScaleHeight     =   615
               ScaleWidth      =   8595
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   3780
               Width           =   8595
               Begin VB.TextBox mskCPF 
                  Height          =   285
                  Left            =   1230
                  MaxLength       =   20
                  TabIndex        =   14
                  Text            =   "mskCPF"
                  Top             =   0
                  Width           =   2115
               End
               Begin VB.TextBox txtConta 
                  Height          =   285
                  Left            =   6810
                  MaxLength       =   15
                  TabIndex        =   18
                  Text            =   "txtConta"
                  Top             =   300
                  Width           =   1695
               End
               Begin VB.TextBox txtAgencia 
                  Height          =   285
                  Left            =   4890
                  MaxLength       =   10
                  TabIndex        =   17
                  Text            =   "txtAgencia"
                  Top             =   330
                  Width           =   1095
               End
               Begin VB.TextBox txtNroCheque 
                  Height          =   285
                  Left            =   4890
                  MaxLength       =   15
                  TabIndex        =   15
                  Text            =   "txtNroCheque"
                  Top             =   0
                  Width           =   1695
               End
               Begin VB.ComboBox cboBanco 
                  Height          =   315
                  Left            =   1230
                  Style           =   2  'Dropdown List
                  TabIndex        =   16
                  Top             =   300
                  Width           =   2475
               End
               Begin VB.Label Label11 
                  Caption         =   "CPF/CNPJ"
                  Height          =   255
                  Index           =   7
                  Left            =   0
                  TabIndex        =   59
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.Label Label11 
                  Caption         =   "Conta"
                  Height          =   255
                  Index           =   3
                  Left            =   6120
                  TabIndex        =   58
                  Top             =   330
                  Width           =   645
               End
               Begin VB.Label Label11 
                  Caption         =   "Agência"
                  Height          =   255
                  Index           =   4
                  Left            =   3840
                  TabIndex        =   57
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.Label Label11 
                  Caption         =   "Nro. Cheque"
                  Height          =   255
                  Index           =   6
                  Left            =   3840
                  TabIndex        =   56
                  Top             =   0
                  Width           =   1455
               End
               Begin VB.Label Label11 
                  Caption         =   "Banco"
                  Height          =   255
                  Index           =   8
                  Left            =   0
                  TabIndex        =   55
                  Top             =   300
                  Width           =   1455
               End
            End
            Begin VB.PictureBox picPgtoCartaoDeb 
               BorderStyle     =   0  'None
               Height          =   345
               Left            =   30
               ScaleHeight     =   345
               ScaleWidth      =   8595
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   4440
               Width           =   8595
               Begin VB.ComboBox cboCartaoDebito 
                  Height          =   315
                  Left            =   1230
                  Style           =   2  'Dropdown List
                  TabIndex        =   19
                  Top             =   0
                  Width           =   2055
               End
               Begin VB.Label Label11 
                  Caption         =   "Débito/Convênio"
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   53
                  Top             =   30
                  Width           =   1245
               End
            End
            Begin VB.PictureBox picPgtoCartaoCred 
               BorderStyle     =   0  'None
               Height          =   345
               Left            =   30
               ScaleHeight     =   345
               ScaleWidth      =   8595
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   4830
               Width           =   8595
               Begin VB.TextBox txtLote 
                  Height          =   285
                  Left            =   4470
                  MaxLength       =   10
                  TabIndex        =   21
                  Text            =   "txtLote"
                  Top             =   0
                  Width           =   1935
               End
               Begin VB.ComboBox cboCartao 
                  Height          =   315
                  Left            =   1230
                  Style           =   2  'Dropdown List
                  TabIndex        =   20
                  Top             =   0
                  Width           =   2055
               End
               Begin VB.Label Label11 
                  Caption         =   "Lote"
                  Height          =   255
                  Index           =   10
                  Left            =   3390
                  TabIndex        =   51
                  Top             =   30
                  Width           =   615
               End
               Begin VB.Label Label11 
                  Caption         =   "Cartão"
                  Height          =   255
                  Index           =   5
                  Left            =   0
                  TabIndex        =   50
                  Top             =   0
                  Width           =   975
               End
            End
            Begin VB.ComboBox cboDebitoCredito 
               Height          =   315
               Left            =   4620
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   1080
               Width           =   2055
            End
            Begin VB.ComboBox cboFuncionario 
               Height          =   315
               Left            =   1380
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   690
               Visible         =   0   'False
               Width           =   4245
            End
            Begin VB.TextBox txtResponsavel 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   4140
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   1
               TabStop         =   0   'False
               Text            =   "txtResponsavel"
               Top             =   60
               Width           =   4455
            End
            Begin VB.PictureBox Picture4 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               ScaleHeight     =   255
               ScaleWidth      =   3015
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   60
               Width           =   3015
               Begin MSMask.MaskEdBox mskData 
                  Height          =   255
                  Index           =   1
                  Left            =   1260
                  TabIndex        =   0
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  _Version        =   393216
                  BackColor       =   14737632
                  AutoTab         =   -1  'True
                  MaxLength       =   16
                  Mask            =   "##/##/#### ##:##"
                  PromptChar      =   "_"
               End
               Begin VB.Label Label4 
                  Caption         =   "Dt./Hr.  Receb."
                  Height          =   255
                  Left            =   0
                  TabIndex        =   40
                  Top             =   0
                  Width           =   1155
               End
            End
            Begin VB.TextBox txtTotalsDesc 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1380
               Locked          =   -1  'True
               TabIndex        =   2
               TabStop         =   0   'False
               Text            =   "txtTotalsDesc"
               Top             =   360
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox txtDesconto 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4140
               Locked          =   -1  'True
               TabIndex        =   3
               TabStop         =   0   'False
               Text            =   "txtDesconto"
               ToolTipText     =   "DESCONTO EM %"
               Top             =   390
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox txtTotalaPagar 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7140
               Locked          =   -1  'True
               TabIndex        =   4
               TabStop         =   0   'False
               Text            =   "txtTotalaPagar"
               Top             =   390
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox txtTroco 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7170
               Locked          =   -1  'True
               TabIndex        =   10
               TabStop         =   0   'False
               Text            =   "txtTroco"
               ToolTipText     =   "DESCONTO EM %"
               Top             =   1410
               Width           =   1455
            End
            Begin VB.PictureBox picTravaGorjeta 
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   3180
               ScaleHeight     =   255
               ScaleWidth      =   3135
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   1440
               Width           =   3135
               Begin MSMask.MaskEdBox mskGorjeta 
                  Height          =   255
                  Left            =   1440
                  TabIndex        =   9
                  Top             =   0
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   450
                  _Version        =   393216
                  BackColor       =   14737632
                  AutoTab         =   -1  'True
                  Format          =   "#,##0.00;($#,##0.00)"
                  PromptChar      =   "_"
               End
               Begin VB.Label Label34 
                  Caption         =   "Gorjeta"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   38
                  Top             =   0
                  Width           =   885
               End
            End
            Begin MSMask.MaskEdBox mskValor 
               Height          =   255
               Left            =   1380
               TabIndex        =   8
               Top             =   1410
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label7 
               Caption         =   "Troco"
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
               Left            =   3930
               TabIndex        =   76
               Top             =   5250
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label6 
               Caption         =   "Restante"
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
               Left            =   6150
               TabIndex        =   72
               Top             =   5280
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.Label Label3 
               Caption         =   "Lançado"
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
               Left            =   1410
               TabIndex        =   71
               Top             =   5250
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.Label Label11 
               Caption         =   "Deb/Cred"
               Height          =   255
               Index           =   2
               Left            =   3540
               TabIndex        =   48
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label11 
               Caption         =   "Funcionário"
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   47
               Top             =   690
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Label17 
               Caption         =   "Valor"
               Height          =   255
               Left            =   150
               TabIndex        =   46
               Top             =   1410
               Width           =   1455
            End
            Begin VB.Label Label5 
               Caption         =   "Responsável"
               Height          =   255
               Left            =   3120
               TabIndex        =   45
               Top             =   60
               Width           =   1035
            End
            Begin VB.Label Label33 
               Caption         =   "Total s/ desc"
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
               TabIndex        =   44
               Top             =   360
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label35 
               Caption         =   "Desc."
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
               Left            =   3120
               TabIndex        =   43
               Top             =   390
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label Label38 
               Caption         =   "Total a pagar"
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
               Left            =   5760
               TabIndex        =   42
               Top             =   390
               Visible         =   0   'False
               Width           =   1245
            End
            Begin VB.Label Label53 
               Caption         =   "Troco"
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
               Left            =   6390
               TabIndex        =   41
               Top             =   1410
               Width           =   975
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserContaCorrente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                 As tpStatus
Public blnRetorno             As Boolean
Public blnFechar              As Boolean
Public strFuncionarioNome     As String
'Informa Qual Grupo irá chamar os Tabs 3 - Fechamento...
Public intGrupo               As Integer
Public strGrupo               As String
Public strNivelAcesso         As String
Private blnFatura             As Boolean
'
Public strStatusLanc          As String
'CC - Conta Corrente
'RC - Recebimento
'RE - Recebimento Empresa
'DP - Depósito
'
Public lngGRID                As Long
Public lngCCID                As Long
Public lngTurnoRecebeId       As Long

Private blnPrimeiraVez        As Boolean
Public blnFecharContaCorrente As Boolean

Public blnChamadaViz          As Boolean

'Default Values
Private Const intTopSuperior      As Integer = 1710
Private Const intLeftSuperior     As Integer = 150
Private Const intTopInferior      As Integer = 4830
Private Const intLeftInferior     As Integer = 150

Private vrCalcTroco               As Currency

Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String

Private Sub CapturaTotais()
  On Error GoTo trata
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim objGeral      As busSisMed.clsGeral
  Dim vrTotGR       As Currency
  Dim vrLancado     As Currency
  Dim vrRestante    As Currency
  Dim vrGorjeta     As Currency
  Dim vrTroco       As Currency
  'Captura total a ser pago pela empresa
  vrLancado = 0
  vrRestante = 0
  vrGorjeta = 0
  vrTroco = 0
  Set objGeral = New busSisMed.clsGeral

  Select Case strStatusLanc
  Case "RC"
    vrTotGR = CCur(IIf(Not IsNumeric(txtTotalaPagar.Text), 0, txtTotalaPagar.Text))
    strSql = "SELECT SUM(case INDDEBITOCREDITO when 'C' then VALOR else (VALOR * (-1)) end) AS VRTOTAL, SUM(VRGORJETA) AS VRGORJETAJAPAGO, SUM(vrTroco) As VRTROCOJAPAGO FROM CONTACORRENTE " & _
      "WHERE CONTACORRENTE.STATUSLANCAMENTO = " & Formata_Dados("RC", tpDados_Texto) & _
      " AND GRID = " & Formata_Dados(lngGRID, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      If IsNumeric(objRs.Fields("VRTOTAL").Value) Then
        vrLancado = objRs.Fields("VRTOTAL").Value
      End If
      If IsNumeric(objRs.Fields("VRGORJETAJAPAGO").Value) Then
        vrGorjeta = objRs.Fields("VRGORJETAJAPAGO").Value
      End If
      If IsNumeric(objRs.Fields("VRTROCOJAPAGO").Value) Then
        vrTroco = objRs.Fields("VRTROCOJAPAGO").Value
      End If
    End If
    objRs.Close
    Set objRs = Nothing
    'Calcula Vr Restante
    vrRestante = vrTotGR + vrGorjeta + vrTroco - vrLancado
  End Select
  'Jogar Valor na tela
  txtLancado.Text = Format(vrLancado, "###,##0.00")
  txtTrocoTotal.Text = Format(vrTroco, "###,##0.00")
  txtRestante.Text = Format(vrRestante, "###,##0.00")
  Set objGeral = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source, _
            Err.Description
End Sub

Private Sub TratarCombos()
  On Error GoTo trata
  Dim strSql    As String
  strSql = "SELECT PRONTUARIO.NOME FROM PRONTUARIO INNER JOIN FUNCIONARIO ON PRONTUARIO.PKID = FUNCIONARIO.PRONTUARIOID WHERE FUNCIONARIO.INDEXCLUIDO = " & Formata_Dados("N", tpDados_Texto) & " ORDER BY PRONTUARIO.NOME"
  PreencheCombo cboFuncionario, strSql, False, True
  strSql = "SELECT NOME FROM CARTAO ORDER BY NOME"
  PreencheCombo cboCartao, strSql, False, True
  strSql = "SELECT NOME FROM CARTAODEBITO ORDER BY NOME"
  PreencheCombo cboCartaoDebito, strSql, False, True
  strSql = "SELECT NOME FROM BANCO ORDER BY NOME"
  PreencheCombo cboBanco, strSql, False, True
  'Combo Tipo de Pagamento
  cboTipoPagamento.Clear

  Select Case strStatusLanc
  Case "RC"
    cboTipoPagamento.AddItem "Cartão de Crédito"
    cboTipoPagamento.AddItem "Cartão de Débito"
    'cboTipoPagamento.AddItem "Cheque"
    cboTipoPagamento.AddItem "Espécie"
    'cboTipoPagamento.AddItem "Penhor"
    'If gbTrabComPgtoFatura = True Then
    '  cboTipoPagamento.AddItem "Fatura"
    'End If
  End Select
  'Combo Tipo
  cboDebitoCredito.Clear
  cboDebitoCredito.AddItem "Crédito"
  cboDebitoCredito.AddItem "Débito"
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source, _
            Err.Description
End Sub

Private Sub TratarTotais()
  On Error GoTo trata
  Dim objGR    As busSisMed.clsGR
  Dim objRs     As ADODB.Recordset
  'Independente do Status, calcular os totais a serem pagos
  Select Case strStatusLanc
  Case "RC"
    'Recebimento para empresa e Recebimento
    Set objGR = New busSisMed.clsGR
    '

    Set objRs = objGR.SelecionarGRTotalProc(lngGRID)
    If Not objRs.EOF Then
      txtTotalaPagar.Text = Format(IIf(Not IsNumeric(objRs.Fields("TOTALAPAGAR").Value), 0, objRs.Fields("TOTALAPAGAR").Value), "###,##0.00")
      txtTotalsDesc.Text = Format(IIf(Not IsNumeric(objRs.Fields("TOTALAPAGAR").Value), 0, objRs.Fields("TOTALAPAGAR").Value), "###,##0.00")
      txtDesconto.Text = Format("0", "###,##0.00")
    Else
      txtTotalaPagar.Text = Format("0", "###,##0.00")
      txtTotalsDesc.Text = Format("0", "###,##0.00")
      txtDesconto.Text = Format("0", "###,##0.00")
    End If
    '
    objRs.Close
    Set objRs = Nothing
    Set objGR = Nothing
  End Select
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source, _
            Err.Description

End Sub
Private Sub TratarCampos()
  On Error GoTo trata
  'Configurações iniciais
  If strStatusLanc = "DE" Then
    'Resgatar cheque apenas no depósito
    cmdCalculadora.Enabled = True
  Else
    cmdCalculadora.Enabled = False
  End If

  If Not gbTrabComTroco Then
    'Caso não trabalhe com troco, Desabilita ContaCorrente em Gorjeta
    picTravaGorjeta.Enabled = False
    mskGorjeta.BackColor = &HE0E0E0
  Else
    picTravaGorjeta.Enabled = True
    mskGorjeta.BackColor = vbWhite
  End If
  'CONFIGURAÇÕES DO SISTEMA
  If gbTrabComChequesBons Then
    If strStatusLanc = "RC" Then
      Label11(7).Enabled = False
      mskCpf.Enabled = False
    End If
  End If
  '
  Select Case strStatusLanc
  Case "RC"
    Label53.Visible = True
    Label34.Visible = True
    Label11(9).Visible = True
    Label35.Visible = True
    Label33.Visible = True
    Label38.Visible = True
    txtTotalaPagar.Visible = True
    cmdParcela.Enabled = True
    '
    txtTroco.Visible = True
    mskGorjeta.Visible = False
    Label34.Visible = False
    cboFuncionario.Visible = True
    txtDesconto.Visible = True
    txtTotalsDesc.Visible = True
    txtRestante.Visible = True
    txtLancado.Visible = True
    txtTrocoTotal.Visible = True
    Label3.Visible = True
    Label6.Visible = True
    Label7.Visible = True
  End Select
  Select Case strStatusLanc & ""
  Case "RC"
    Me.Caption = "Módulo Recebimento - GR " & strFuncionarioNome
  Case Else
    Me.Caption = "Recebimento - Status não definido"
  End Select
  '
  If strStatusLanc = "CC" Then
    cmdImprimir.Visible = True
  Else
    cboDebitoCredito.Enabled = False
    cmdImprimir.Visible = False
  End If
  On Error Resume Next
  cboFuncionario.Text = strFuncionarioNome
  cboFuncionario.Enabled = False
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserContaCorrente.TratarCampos]"
End Sub


Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  '
  LimparCampoCombo cboFuncionario
  LimparCampoMask mskData(1)
  LimparCampoTexto txtResponsavel
  LimparCampoTexto txtTotalsDesc
  LimparCampoMask mskValor
  LimparCampoTexto txtDesconto
  LimparCampoTexto txtTotalaPagar
  LimparCampoMask mskGorjeta
  LimparCampoTexto txtTroco
  LimparCampoCombo cboCartao
  LimparCampoTexto txtLote
  LimparCampoTexto mskCpf
  LimparCampoTexto txtNroCheque
  LimparCampoCombo cboBanco
  LimparCampoTexto txtAgencia
  LimparCampoTexto txtConta
  LimparCampoTexto txtCliente
  LimparCampoTexto txtDocumentoPenhor
  LimparCampoTexto txtObjeto
  LimparCampoMask mskNroParcelas
  LimparCampoMask mskData(0)
  LimparCampoTexto txtRestante
  LimparCampoTexto txtLancado
  LimparCampoTexto txtTrocoTotal
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserContaCorrente.LimparCampos]"
End Sub


Private Sub cboBanco_LostFocus()
  Pintar_Controle cboBanco, tpCorContr_Normal
End Sub

Private Sub cboCartaoDebito_LostFocus()
  Pintar_Controle cboCartaoDebito, tpCorContr_Normal
End Sub

Private Sub cboCartao_LostFocus()
  Pintar_Controle cboCartao, tpCorContr_Normal
End Sub

Private Sub cboFuncionario_LostFocus()
  Pintar_Controle cboFuncionario, tpCorContr_Normal
End Sub
Sub TratarTipoRecebimentoCC(Optional blnHabilita As Boolean = False)
  On Error Resume Next
  'Cartão de Crédito
  If blnHabilita Then
    picPgtoCartaoCred.Top = intTopSuperior
    picPgtoCartaoCred.Left = intLeftSuperior
    picPgtoCartaoCred.Visible = True
  Else
    picPgtoCartaoCred.Top = intTopInferior
    picPgtoCartaoCred.Left = intLeftInferior
    picPgtoCartaoCred.Visible = False
    cboCartao.ListIndex = -1
    txtLote.Text = ""
  End If
End Sub
Sub TratarTipoRecebimentoCD(Optional blnHabilita As Boolean = False)
  On Error Resume Next
  'Cartão de Débito
  If blnHabilita Then
    picPgtoCartaoDeb.Top = intTopSuperior
    picPgtoCartaoDeb.Left = intLeftSuperior
    picPgtoCartaoDeb.Visible = True
  Else
    picPgtoCartaoDeb.Top = intTopInferior
    picPgtoCartaoDeb.Left = intLeftInferior
    picPgtoCartaoDeb.Visible = False
    cboCartaoDebito.ListIndex = -1
  End If
End Sub
Sub TratarTipoRecebimentoCH(Optional blnHabilita As Boolean = False)
  On Error Resume Next
  'Cheque
  If blnHabilita Then
    picPgtoCheque.Top = intTopSuperior
    picPgtoCheque.Left = intLeftSuperior
    picPgtoCheque.Visible = True
  Else
    picPgtoCheque.Top = intTopInferior
    picPgtoCheque.Left = intLeftInferior
    picPgtoCheque.Visible = False
    LimparCampoTexto mskCpf
    LimparCampoTexto txtNroCheque
    cboBanco.ListIndex = -1
    LimparCampoTexto txtAgencia
    LimparCampoTexto txtConta
  End If
End Sub
Sub TratarTipoRecebimentoPH(Optional blnHabilita As Boolean = False)
  On Error Resume Next
  'Penhor
  If blnHabilita Then
    picPgtoPenhor.Top = intTopSuperior
    picPgtoPenhor.Left = intLeftSuperior
    picPgtoPenhor.Visible = True
  Else
    picPgtoPenhor.Top = intTopInferior
    picPgtoPenhor.Left = intLeftInferior
    picPgtoPenhor.Visible = False
    LimparCampoTexto txtCliente
    LimparCampoTexto txtDocumentoPenhor
    LimparCampoTexto txtObjeto
  End If
End Sub
Sub TratarTipoRecebimentoFT(Optional blnHabilita As Boolean = False)
  On Error Resume Next
  'Penhor
  If blnHabilita Then
    picPgtoFatura.Top = intTopSuperior
    picPgtoFatura.Left = intLeftSuperior
    picPgtoFatura.Visible = True
  Else
    picPgtoFatura.Top = intTopInferior
    picPgtoFatura.Left = intLeftInferior
    picPgtoFatura.Visible = False
    LimparCampoMask mskNroParcelas
    LimparCampoMask mskData(0)
  End If
End Sub

Sub TratarTipoPagamento(strTipoPagamento As String)
  On Error Resume Next
  '
  Select Case strTipoPagamento
  Case "Cartão de Crédito"
    TratarTipoRecebimentoCC True
    TratarTipoRecebimentoCD
    TratarTipoRecebimentoCH
    TratarTipoRecebimentoPH
    TratarTipoRecebimentoFT
  Case "Cartão de Débito"
    TratarTipoRecebimentoCC
    TratarTipoRecebimentoCD True
    TratarTipoRecebimentoCH
    TratarTipoRecebimentoPH
    TratarTipoRecebimentoFT
  Case "Cheque"
    TratarTipoRecebimentoCC
    TratarTipoRecebimentoCD
    TratarTipoRecebimentoCH True
    TratarTipoRecebimentoPH
    TratarTipoRecebimentoFT
  Case "Espécie"
    TratarTipoRecebimentoCC
    TratarTipoRecebimentoCD
    TratarTipoRecebimentoCH
    TratarTipoRecebimentoPH
    TratarTipoRecebimentoFT
  Case "Penhor"
    TratarTipoRecebimentoCC
    TratarTipoRecebimentoCD
    TratarTipoRecebimentoCH
    TratarTipoRecebimentoPH True
    TratarTipoRecebimentoFT
  Case "Fatura"
    TratarTipoRecebimentoCC
    TratarTipoRecebimentoCD
    TratarTipoRecebimentoCH
    TratarTipoRecebimentoPH
    TratarTipoRecebimentoFT True
  Case Else
    TratarTipoRecebimentoCC
    TratarTipoRecebimentoCD
    TratarTipoRecebimentoCH
    TratarTipoRecebimentoPH
    TratarTipoRecebimentoFT
  End Select
End Sub
Private Sub cboTipoPagamento_Click()
  TratarTipoPagamento cboTipoPagamento.Text
End Sub

Private Sub cmdAlterar_Click()
  On Error GoTo trata
  If Not IsNumeric(grdGeral.Columns("ID").Value & "") Then
    MsgBox "Selecione um lançamento !", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  ElseIf strStatusLanc <> grdGeral.Columns("StatusLancamento").Value & "" Then
    MsgBox "Selecione um lançamento da mesma origem que " & IIf(strStatusLanc = "CC", "Conta Corrente", IIf(strStatusLanc = "RC", "Recebimento", IIf(strStatusLanc = "RE", "Recebimento Empresa", IIf(strStatusLanc = "DP", "Depósito", IIf(strStatusLanc = "DE", "Despesa", IIf(strStatusLanc = "VD", "Venda", IIf(strStatusLanc = "EX", "Extra", ""))))))), vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If
  lngCCID = grdGeral.Columns("ID").Value
  strStatusLanc = grdGeral.Columns("StatusLancamento").Value
  Status = tpStatus_Alterar
  Form_Load
  SetarFoco cboTipoPagamento
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub cmdCalculadora_Click()
  On Error GoTo trata
  Dim vrValorJaPago   As Currency
  Dim strSql          As String
  Dim vrTotGR        As Currency

  'Demais campos calculadora
  frmUserCalculadora.Status = tpStatus_Consultar
  frmUserCalculadora.txtTotalsDesc.Text = txtTotalsDesc.Text
  frmUserCalculadora.txtDesconto.Text = txtDesconto.Text
  frmUserCalculadora.txtTotalaPagar.Text = txtTotalaPagar.Text
  frmUserCalculadora.txtTotalaPagar.Text = txtTotalaPagar.Text
  frmUserCalculadora.txtFuncionario.Text = strFuncionarioNome
  frmUserCalculadora.Show vbModal
  SetarFoco mskValor

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdCancelar_Click()
  On Error GoTo trata
  Dim vrValorJaPago       As Currency
  Dim vrTotGR             As Currency
  Dim strSql              As String

  Dim objRs               As ADODB.Recordset
  Dim objGeral            As busSisMed.clsGeral

  Select Case strStatusLanc
  Case "DE", "VD", "EX"
'''    'Capturar valor já pago
'''    vrValorJaPago = 0
'''    Set objGeral = New busSisMed.clsGeral
'''    strSql = "SELECT SUM(VALOR) AS VALORJAPAGO, SUM(VRGORJETA) AS VRGORJETAJAPAGO, SUM(VRTROCO) AS VRTROCOJAPAGO " & _
'''      "FROM CONTACORRENTE "
'''
'''    If strStatusLanc = "DE" Then
'''      strSql = strSql & " WHERE DESPESAID = " & Formata_Dados(lngGRID, tpDados_Longo, tpNulo_Aceita)
'''    ElseIf strStatusLanc = "VD" Then
'''      strSql = strSql & " WHERE STATUSLANCAMENTO = " & Formata_Dados(strStatusLanc, tpDados_Texto, tpNulo_Aceita)
'''      strSql = strSql & " AND VENDAID = " & Formata_Dados(lngGRID, tpDados_Longo, tpNulo_Aceita)
'''    ElseIf strStatusLanc = "EX" Then
'''      strSql = strSql & " WHERE STATUSLANCAMENTO = " & Formata_Dados(strStatusLanc, tpDados_Texto, tpNulo_Aceita)
'''      strSql = strSql & " AND EXTRAID = " & Formata_Dados(lngGRID, tpDados_Longo, tpNulo_Aceita)
'''    End If
'''
'''    Set objRs = objGeral.ExecutarSQL(strSql)
'''    If Not objRs.EOF Then
'''      If Not IsNull(objRs.Fields("VALORJAPAGO").Value) Then
'''        vrValorJaPago = objRs.Fields("VALORJAPAGO").Value
'''      End If
'''      If Not IsNull(objRs.Fields("VRGORJETAJAPAGO").Value) Then
'''        If strStatusLanc <> "DE" Then
'''          vrValorJaPago = vrValorJaPago - objRs.Fields("VRGORJETAJAPAGO").Value
'''        End If
'''      End If
'''      If Not IsNull(objRs.Fields("VRTROCOJAPAGO").Value) Then
'''        If strStatusLanc <> "DE" Then
'''          vrValorJaPago = vrValorJaPago - objRs.Fields("VRTROCOJAPAGO").Value
'''        End If
'''      End If
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
'''    Set objGeral = Nothing
'''    'Depende do Tipo
'''    vrTotGR = CCur(IIf(Not IsNumeric(txtTotalaPagar.Text), 0, txtTotalaPagar.Text))
'''    If vrValorJaPago <> vrTotGR Then
'''      'Valor do pagamento < que valor a pagar
'''      TratarErroPrevisto "Pagamento não pode ser diferente do restante. Favor complementá-la."
'''      SetarFoco cboTipoPagamento
'''    Else
'''      'Cancelar Cartão
'''      blnFechar = True
'''      Unload Me
'''    End If
'''  Case "DP"
'''    'Capturar valor já pago
'''    vrValorJaPago = 0
'''    Set objGeral = New busSisMed.clsGeral
'''    strSql = "SELECT * " & _
'''      "FROM CONTACORRENTE "
'''
'''    If strStatusLanc = "DP" Then
'''      strSql = strSql & " WHERE LOCACAOID = " & Formata_Dados(lngGRID, tpDados_Longo, tpNulo_Aceita)
'''    End If
'''
'''    Set objRs = objGeral.ExecutarSQL(strSql)
'''    If objRs.EOF Then
'''      If frmMDI.objForm.blnDeposito = True Then
'''        'Está vindo da entrada
'''        'Depósito ainda não lançado
'''        TratarErroPrevisto "Lançamento de depósito obrigatório para entradas de clientes a pé."
'''        SetarFoco cboTipoPagamento
'''      Else
'''        'está vindo do depósito,
'''        'caso haja mudança emitir boleto
'''        blnFechar = True
'''        Unload Me
'''      End If
'''    Else
'''      If frmMDI.objForm.blnDeposito = False And blnRetorno Then
'''        'emitir comprovante de depósito
'''        IMP_COMPROV_DEPOSITO lngGRID, gsNomeEmpresa, 1, strFuncionarioNome
'''      End If
'''      'ok
'''      blnFechar = True
'''      Unload Me
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
'''    Set objGeral = Nothing
'''    'Depende do Tipo
  Case Else
    'Cancelar Cartão
    blnFechar = True
    Unload Me
  End Select


  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub


Private Sub cmdExcluir_Click()
  Dim objContaCorrente        As busSisMed.clsContaCorrente
  Dim objParcela              As busSisMed.clsParcela
  '
  On Error GoTo trata
  If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
    MsgBox "Selecione um lançamento.", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  ElseIf strStatusLanc <> "RC" And strStatusLanc <> "DE" And strStatusLanc <> grdGeral.Columns("StatusLancamento").Value & "" Then
    MsgBox "Selecione um lançamento da mesma origem que " & IIf(strStatusLanc = "CC", "Conta Corrente", IIf(strStatusLanc = "RC", "Recebimento", IIf(strStatusLanc = "RE", "Recebimento Empresa", IIf(strStatusLanc = "DP", "Depósito", IIf(strStatusLanc = "DE", "Despesa", IIf(strStatusLanc = "VD", "Venda", IIf(strStatusLanc = "EX", "Extra", ""))))))), vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If
  '
  '
  If MsgBox(IIf(grdGeral.Columns("Tipo de Pagamento").Value = "Fatura", "A exclusão da fatura excluirá também suas parcelas." & vbCrLf & vbCrLf, "") & "Confirma exclusão do lançamento [" & grdGeral.Columns("Origem").Value & "] / [" & grdGeral.Columns("Tipo de Pagamento").Value & "] ?", vbYesNo, TITULOSISTEMA) = vbNo Then
    SetarFoco grdGeral
    Exit Sub
  End If
  '--------------------------------
  '--------------------------------
  '
  Set objContaCorrente = New busSisMed.clsContaCorrente
  Set objParcela = New busSisMed.clsParcela
  'OK
  objParcela.ExcluirParcelasDaCC CLng(grdGeral.Columns("ID").Value)
  objContaCorrente.ExcluirContaCorrente strStatusLanc, _
                                        CLng(grdGeral.Columns("ID").Value)

  VerificaMovAposFecha lngGRID
  '
  MontaMatriz
  grdGeral.Bookmark = Null
  grdGeral.ReBind

  Set objContaCorrente = Nothing
  Set objParcela = Nothing
  SetarFoco cboTipoPagamento
  'Fechamento e impressão
  'blnFechar = True
  lngCCID = 0
  Status = tpStatus_Incluir
  Form_Load
  blnRetorno = True
  'Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub


'''Private Sub cmdImprimir_Click()
'''  On Error GoTo TratErro
'''  AmpS
'''  '
'''  'Cabeçalho do report
'''  grdGeral.PrintInfo.PageHeader = "Conta Corrente - emissão: " & Format(Now, "DD/MM/YYYY hh:mm")
'''  'grdGeral.PrintInfo.PageHeader = grdGeral.PrintInfo.PageHeader & vbCrLf & ""
'''  grdGeral.PrintInfo.RepeatColumnHeaders = True
'''  '
'''  grdGeral.PrintInfo.SettingsMarginBottom = 400
'''  grdGeral.PrintInfo.SettingsMarginLeft = 1000
'''  grdGeral.PrintInfo.SettingsMarginRight = 1000
'''  grdGeral.PrintInfo.SettingsMarginTop = 600
'''  grdGeral.PrintInfo.PreviewMaximize = True
'''  grdGeral.PrintInfo.SettingsOrientation = 1
'''  grdGeral.PrintInfo.PrintPreview
'''  '
'''  AmpN
'''  Exit Sub
'''
'''TratErro:
'''  AmpN
'''  MsgBox "O seguinte Erro Ocorreu: " & Err.Description, vbOKOnly, TITULOSISTEMA
'''
'''End Sub
'''
'''
Private Sub cmdIncluir_Click()
  On Error GoTo trata
  lngCCID = 0
  Status = tpStatus_Incluir
  Form_Load
  SetarFoco cboTipoPagamento
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub cmdOk_Click()
  Dim objCC               As busSisMed.clsContaCorrente
  Dim objGR               As busSisMed.clsGR
  Dim objGeral            As busSisMed.clsGeral
  Dim objParcela          As busSisMed.clsParcela
  Dim objRs               As ADODB.Recordset
  Dim strSql              As String
  '
  Dim vrTotGR      As Currency
  Dim vrTotDescGR  As Currency

  Dim vrPago        As Currency

  Dim vrValor       As Currency
  Dim vrGorjeta     As Currency
  '
  Dim vrValorJaPago   As Currency
  '
  Dim strFuncionarioId      As String
  Dim strCartaoId           As String
  Dim strCartaoDebId        As String
  Dim strBancoId            As String

  Dim strIndDebCred   As String
  Dim strStatusCC     As String
  Dim strFuncionario  As String
  Dim vrCalcTotal     As Currency

  On Error GoTo trata
  vrCalcTroco = 0
  If Not ValidaCampos Then
    Exit Sub
  End If
  '--------------------------------
  '--------------------------------
  Set objCC = New busSisMed.clsContaCorrente
  Set objGeral = New busSisMed.clsGeral
  '
  'Calcula campos para serem gravados na base
  vrValor = CCur(IIf(Not IsNumeric(mskValor.Text), 0, mskValor.Text))
  vrGorjeta = CCur(IIf(Not IsNumeric(mskGorjeta.Text), 0, mskGorjeta.Text))
  'Calcula Valor Pago
  vrPago = vrValor - vrGorjeta
  vrTotGR = CCur(IIf(Not IsNumeric(txtTotalsDesc.Text), 0, txtTotalsDesc.Text))
  vrTotDescGR = CCur(IIf(Not IsNumeric(txtDesconto.Text), 0, txtDesconto.Text))

  'vrCalcTroco = vrCalcTroco0 'Sem troco Por enquanto
'  If gbTrabComTroco Then
'    'Caso trabalhe com troco, joga diferença para garçom
'    vrCalcTroco = vrPago - (vrTotGR - vrTotDescGR)
'    If vrCalcTroco < 0 Then vrCalcTroco = 0
''''  Else
''''    'Se não trabalha com troco joga diferença para Troco + O troco Digitado
''''    mskGorjeta.Text = Format((vrPago - (vrTotGR - vrTotDescGR)) + vrGorjeta, "###,###,##0.00")
'  End If
  'Capturar campos para gravar na base
  lngTurnoRecebeId = IIf(lngTurnoRecebeId = 0, RetornaCodTurnoCorrente, lngTurnoRecebeId)
  '
  strFuncionarioId = ""
  strSql = "SELECT PRONTUARIO.PKID FROM PRONTUARIO INNER JOIN FUNCIONARIO ON PRONTUARIO.PKID = FUNCIONARIO.PRONTUARIOID " & _
    " WHERE FUNCIONARIO.INDEXCLUIDO = " & Formata_Dados("N", tpDados_Texto) & _
    " AND PRONTUARIO.NOME = " & Formata_Dados(cboFuncionario.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    strFuncionarioId = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  strCartaoId = ""
  strSql = "SELECT CARTAO.PKID FROM CARTAO WHERE CARTAO.NOME = " & Formata_Dados(cboCartao.Text, tpDados_Texto, tpNulo_Aceita)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    strCartaoId = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  strCartaoDebId = ""
  strSql = "SELECT CARTAODEBITO.PKID FROM CARTAODEBITO WHERE CARTAODEBITO.NOME = " & Formata_Dados(cboCartaoDebito.Text, tpDados_Texto, tpNulo_Aceita)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    strCartaoDebId = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  strBancoId = ""
  strSql = "SELECT BANCO.PKID FROM BANCO WHERE BANCO.NOME = " & Formata_Dados(cboBanco.Text, tpDados_Texto, tpNulo_Aceita)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    strBancoId = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  If cboDebitoCredito = "Débito" Then
    strIndDebCred = "D"
  Else
    strIndDebCred = "C"
  End If
  If cboTipoPagamento = "Cartão de Crédito" Then
    strStatusCC = "CC"
  ElseIf cboTipoPagamento = "Cartão de Débito" Then
    strStatusCC = "CD"
  ElseIf cboTipoPagamento = "Cheque" Then
    strStatusCC = "CH"
  ElseIf cboTipoPagamento = "Espécie" Then
    strStatusCC = "ES"
  ElseIf cboTipoPagamento = "Penhor" Then
    strStatusCC = "PH"
  ElseIf cboTipoPagamento = "Fatura" Then
    strStatusCC = "FT"
  End If
  strFuncionario = cboFuncionario.Text
  If Status = tpStatus_Incluir Then
    'Inclusão
    lngCCID = objCC.InserirCC(lngGRID, _
                              lngTurnoRecebeId, _
                              mskData(1).Text, _
                              mskValor.Text, _
                              strIndDebCred, _
                              strStatusCC, _
                              strStatusLanc, _
                              strCartaoId, _
                              strBancoId, _
                              strFuncionarioId, _
                              txtResponsavel.Text, _
                              txtAgencia.Text, _
                              txtConta.Text, _
                              txtNroCheque.Text, _
                              mskCpf.Text, _
                              txtCliente.Text, _
                              txtObjeto.Text, _
                              txtDocumentoPenhor.Text, _
                              txtLote.Text, _
                              strCartaoDebId, _
                              mskGorjeta.Text, _
                              vrCalcTroco & "", _
                              IIf(mskNroParcelas.ClipText = "", "", mskNroParcelas.Text), _
                              IIf(mskData(0).ClipText = "", "", mskData(0).Text))
  ElseIf Status = tpStatus_Alterar Then
    'Alteração
    objCC.AlterarCC strStatusCC, _
                    lngCCID, _
                    mskValor.Text, _
                    strCartaoId, _
                    strBancoId, _
                    strFuncionarioId, _
                    txtAgencia.Text, _
                    txtConta.Text, _
                    txtNroCheque.Text, _
                    mskCpf.Text, _
                    txtCliente.Text, _
                    txtObjeto.Text, _
                    txtDocumentoPenhor.Text, _
                    txtLote.Text, _
                    strCartaoDebId, _
                    mskGorjeta.Text, _
                    vrCalcTroco & "", _
                    IIf(mskNroParcelas.ClipText = "", "", mskNroParcelas.Text), _
                    IIf(mskData(0).ClipText = "", "", mskData(0).Text)

'''    If Len(Trim(mskPgtoPenhor.Text)) <> 0 Then
'''      'Houve Pagamento Em Penhor
'''      IMP_COMPROV_PENHOR lngGRID, gsNomeEmpresa, 3
'''    End If
'''    '----- Imprimir Impressora Fiscal
'''    If gbTrabComImpFiscal Then _
'''      IMP_CUPOM_FISCAL lngGRID & "", gsNomeEmpresa
'''    INCLUI_LOG_UNIDADE 0, lngGRID, "ContaCorrente da Unidade", "Unidade " & strFuncionarioNome & IIf(IsNumeric(mskValor.Text), " Espécie " & Format(mskValor.Text, "###,##0.00"), "") & IIf(IsNumeric(mskPgtoCartao.Text), " Cartão " & Format(mskPgtoCartao.Text, "###,##0.00"), "") & IIf(IsNumeric(mskPgtoCheque.Text), " Cheque " & Format(mskPgtoCheque.Text, "###,##0.00") & " " & mskCPF.Text, "") & IIf(IsNumeric(mskPgtoPenhor.Text), " Penhor " & Format(mskPgtoPenhor.Text, "###,##0.00"), ""), "", "", "", ""    'gsNomeUsuLib
'''    '------------
  End If
  'Após inclusão ou alteração, redefinir parcelas
  Set objParcela = New busSisMed.clsParcela
  'Excluir parcelas da CC
  objParcela.ExcluirParcelasDaCC lngCCID
  If cboTipoPagamento.Text = "Fatura" Then
    'Cadastrar parcelas da CC
    objParcela.CadastrarParcelas lngCCID, _
                                 IIf(mskNroParcelas.ClipText = "", "", mskNroParcelas.Text), _
                                 IIf(mskData(0).ClipText = "", "", mskData(0).Text), _
                                 mskValor.Text
    frmUserParcelaLis.lngCCID = lngCCID
    frmUserParcelaLis.strFuncionarioNome = strFuncionarioNome
    frmUserParcelaLis.Show vbModal
    'Impressão de nota
    IMP_COMPROV_FATURA lngCCID, gsNomeEmpresa, 1
  End If
  '
  Set objParcela = Nothing
  '
  lngCCID = 0
  Status = tpStatus_Incluir
  Form_Load
  '
  Set objCC = Nothing
  blnRetorno = True
  blnFechar = True
  'Verifica se continuará nesta tela
  Select Case strStatusLanc
  Case "RC"
    'Capturar valor já pago
    vrValorJaPago = 0
    strSql = "SELECT SUM(VALOR * (CASE INDDEBITOCREDITO WHEN 'C' THEN 1 ELSE -1 END)) AS VALORJAPAGO, SUM(VRGORJETA) AS VRGORJETAJAPAGO, SUM(VRTROCO) AS VRTROCOJAPAGO " & _
      "FROM CONTACORRENTE " & _
      "WHERE GRID = " & Formata_Dados(lngGRID, tpDados_Longo, tpNulo_Aceita)

    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      If Not IsNull(objRs.Fields("VALORJAPAGO").Value) Then
        vrValorJaPago = objRs.Fields("VALORJAPAGO").Value
      End If
      If Not IsNull(objRs.Fields("VRGORJETAJAPAGO").Value) Then
        vrValorJaPago = vrValorJaPago - objRs.Fields("VRGORJETAJAPAGO").Value
      End If
      If Not IsNull(objRs.Fields("VRTROCOJAPAGO").Value) Then
        vrValorJaPago = vrValorJaPago - objRs.Fields("VRTROCOJAPAGO").Value
      End If
    End If
    objRs.Close
    Set objRs = Nothing
    'Depende do Tipo
    vrTotGR = CCur(IIf(Not IsNumeric(txtTotalaPagar.Text), 0, txtTotalaPagar.Text))
    If vrValorJaPago < vrTotGR Then
      'Valor do pagamento < que valor a pagar
      SetarFoco cboTipoPagamento
      'Voltar com garçom
      If strFuncionario <> "" Then
        cboFuncionario.Text = strFuncionario
      End If
    Else
      'Está ok, se for recebimento,
      'dar fechamento na GR e entra com os valores do financeiro
      If strStatusLanc = "RC" Then
        Set objGR = New busSisMed.clsGR
        objGR.AlterarStatusGR lngGRID, _
                              "F", _
                              ""
        Set objGR = Nothing
        '
      End If
      If strNivelAcesso = gsLaboratorio Then
        'Se for uma GR de laboratório, alterar turno para o turno do caixa
        Set objGR = New busSisMed.clsGR
        objGR.AlterarTurnoGR lngGRID, _
                             RetornaCodTurnoCorrente, _
                             strFuncionarioId
        Set objGR = Nothing
        'Fecou GR do laboratório, emitir comprovante de pagamento
        IMP_COMPROV_REC lngGRID, gsNomeEmpresa, 1
        
      End If
      Unload Me
    End If
  Case Else
    'Unload Me
    SetarFoco cboTipoPagamento
  End Select
  Set objGeral = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub
'Propósito: Validar o ContaCorrente
Public Function ValidaContaCorrente() As Boolean
  Dim strMsg        As String
  Dim strMsgAlerta  As String
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim vrPago        As Currency

  Dim vrTotGR      As Currency
  Dim vrTotDescGR  As Currency

  Dim vrValor       As Currency
  Dim vrValorJaPago As Currency

  Dim vrGorjeta     As Currency

  Dim DtAtualMenosNDias  As Date

  Dim strMsgAux               As String
  Dim objGeral                As busSisMed.clsGeral
  Dim blnSetarFocoControle    As Boolean
  Dim strCredito              As String
  Dim blnPossuiChqDevolvido   As Boolean
  Dim blnFecharRecebimento    As Boolean
  '
  On Error GoTo trata
  Set objGeral = New busSisMed.clsGeral
  blnSetarFocoControle = True
  If Not Valida_Moeda(mskValor, TpObrigatorio, blnSetarFocoControle) Then
    'Informar o Valor
    strMsg = strMsg & "Informar o Valor válido" & vbCrLf
  End If
  'FATURA
  If Len(strMsg) = 0 And cboTipoPagamento.Text = "Fatura" Then
    If Not Valida_Moeda(mskNroParcelas, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar o Número de parcelas válido." & vbCrLf
    End If
    If Not Valida_Data(mskData(0), TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar a data de pagamento da primeira parcela válida." & vbCrLf
    End If
  End If
  If Len(strMsg) = 0 And cboTipoPagamento.Text = "Fatura" Then
    If CLng(mskNroParcelas.Text) = 0 Then
      strMsg = strMsg & "Informar o número de parcelas maior que zero." & vbCrLf
    End If
  End If
  'CHEQUE
  If Len(strMsg) = 0 And cboTipoPagamento.Text = "Cheque" Then
    'Informou Pagamento em Cheque
    'Testa CPF
    If Len(Trim(mskCpf.Text)) = 0 Then
      strMsg = strMsg & "Informar o CPF/CNPJ" & vbCrLf
      Pintar_Controle mskCpf, tpCorContr_Erro
    ElseIf Len(Trim(mskCpf.Text)) > 11 Then
      'CNPJ
      If Not IsNumeric(mskCpf.Text) Then
        'Não informou o cnpj válido, verifica qual msg será emitida
        If Not gbTrabComChequesBons Then
          strMsg = strMsg & "Informar o CNPJ válido" & vbCrLf
          Pintar_Controle mskCpf, tpCorContr_Erro
        Else
          strMsg = strMsg & "A unidade não possui liberação de cheque" & vbCrLf
        End If
      End If
    ElseIf Not TestaCPF(mskCpf.Text) Then
      'Não informou o cpf, verifica qual msg será emitida
      If Not gbTrabComChequesBons Then
        strMsg = strMsg & "Informar o CPF válido" & vbCrLf
        Pintar_Controle mskCpf, tpCorContr_Erro
      Else
        strMsg = strMsg & "A unidade não possui liberação de cheque" & vbCrLf
      End If
    End If

'''    blnPossuiChqDevolvido = False
'''    blnFecharRecebimento = False
'''    If Len(strMsg) = 0 Then
'''      strSql = "Select CHEQUE.* from CLIENTE INNER JOIN CHEQUE ON CLIENTE.PKID = CHEQUE.CLIENTEID WHERE CLIENTE.CPF  = " & Formata_Dados(Left(mskCPF.Text, 9) & "/" & Right(mskCPF.Text, 2), tpDados_Texto) & " AND STATUS = " & Formata_Dados("D", tpDados_Texto)
'''      Set objRs = objGeral.ExecutarSQL(strSql)
'''      '
'''      If Not objRs.EOF Then
'''        blnPossuiChqDevolvido = True
'''        'MsgBox "Existem cheques devolvidos cadastrados para este CPF. Contacte seu Gerente.", vbOKOnly, TITULOSISTEMA
'''        If gbPedirSenhaSupLibChqReceb Then
'''          'PEDIR SENHA SUPERIOR
'''          '----------------------------
'''          '----------------------------
'''          'Pede Senha Superior (Diretor, Gerente ou Administrador
'''          If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
'''            'Só pede senha superior se quem estiver logado não for superior
'''            frmUserLoginSup.Show vbModal
'''
'''            If Len(Trim(gsNomeUsuLib)) = 0 Then
'''              strMsg = strMsg & "É preciso senha superior para liberar recebimento com cliente possuindo cheques devolvidos. Contacte seu Gerente." & vbCrLf & vbCrLf & "O sistema retornará a tela de entrada sem efetuar o recebimento." & vbCrLf
'''              Pintar_Controle mskCPF, tpCorContr_Erro
'''              blnFecharRecebimento = True
'''            End If
'''            '
'''            'Capturou Nome do Usuário, continua processo de Sangria
'''          Else
'''            gsNomeUsuLib = gsNomeUsu
'''            TratarErroPrevisto "Existem cheques devolvidos cadastrados para este CPF. O pagamento será liberado pois o usuário logado possui permissão."
'''          End If
'''          If Not blnFecharRecebimento Then
'''            'gravar log
'''            INCLUI_LOG_UNIDADE MODOALTERAR, lngGRID, "Liberação de pagamento com cheque, com cliente possuindo cheque devolvido", "Unidade " & strFuncionarioNome & " - CPF Nr. " & mskCPF.Text, "", "", "", gsNomeUsuLib
'''
'''          End If
'''          '--------------------------------
'''          '--------------------------------
'''
'''        Else
'''          strMsg = strMsg & "Existem cheques devolvidos cadastrados para este CPF. Contacte seu Gerente." & vbCrLf & vbCrLf & "O sistema retornará a tela de entrada sem efetuar o recebimento." & vbCrLf
'''          Pintar_Controle mskCPF, tpCorContr_Erro
'''          blnFecharRecebimento = True
'''        End If
'''      End If
'''      objRs.Close
'''      Set objRs = Nothing
'''    End If

  End If
  'CARTÃO DE CRÉDITO
  If Len(strMsg) = 0 And cboTipoPagamento.Text = "Cartão de Crédito" Then
    If Not Valida_String(cboCartao, TpObrigatorio, blnSetarFocoControle) Then
      'Informar o Cartão
      strMsg = strMsg & "Selecionar o Cartão" & vbCrLf
    End If
  End If
  'CARTÃO DE DÉBITO
  If Len(strMsg) = 0 And cboTipoPagamento.Text = "Cartão de Débito" Then
    If Not Valida_String(cboCartaoDebito, TpObrigatorio, blnSetarFocoControle) Then
      'Informar o Cartão
      strMsg = strMsg & "Selecionar o Cartão de Débito" & vbCrLf
    End If
  End If
  'PENHOR
  If Len(strMsg) = 0 And cboTipoPagamento.Text = "Penhor" Then
    If Not Valida_String(txtCliente, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar o Nome do Cliente" & vbCrLf
    End If
    If Not Valida_String(txtDocumentoPenhor, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar o número do Documento" & vbCrLf
    End If
    If Not Valida_String(txtObjeto, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar a Descrição do Objeto deixado como Penhor" & vbCrLf
    End If
  End If

  If Len(strMsg) = 0 Then
    'Capturar valor já pago
    vrValorJaPago = 0
    vrCalcTroco = 0
    strSql = "SELECT SUM(VALOR * (CASE INDDEBITOCREDITO WHEN 'C' THEN 1 ELSE -1 END)) AS VALORJAPAGO, SUM(VRGORJETA) AS VRGORJETAJAPAGO, SUM(VRTROCO) AS VRTROCOJAPAGO " & _
      "FROM CONTACORRENTE "
    Select Case strStatusLanc
    Case "RC"
      strSql = strSql & "WHERE GRID = " & Formata_Dados(lngGRID, tpDados_Longo, tpNulo_Aceita)
    End Select
    If strStatusLanc = "RC" Then
      strSql = strSql & " AND STATUSLANCAMENTO IN ('RC')"
    End If
    strSql = strSql & " AND CONTACORRENTE.PKID <> " & Formata_Dados(lngCCID, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      If Not IsNull(objRs.Fields("VALORJAPAGO").Value) Then
        vrValorJaPago = objRs.Fields("VALORJAPAGO").Value
      End If
      If Not IsNull(objRs.Fields("VRGORJETAJAPAGO").Value) Then
        If strStatusLanc <> "DE" Then
          vrValorJaPago = vrValorJaPago - objRs.Fields("VRGORJETAJAPAGO").Value
        End If
      End If
      If Not IsNull(objRs.Fields("VRTROCOJAPAGO").Value) Then
        If strStatusLanc <> "DE" Then
          vrValorJaPago = vrValorJaPago - objRs.Fields("VRTROCOJAPAGO").Value
        End If
      End If
    End If
    objRs.Close
    Set objRs = Nothing
    'Validar Valor
    vrValor = CCur(IIf(Not IsNumeric(mskValor.Text), 0, mskValor.Text))
    vrGorjeta = CCur(IIf(Not IsNumeric(mskGorjeta.Text), 0, mskGorjeta.Text))
    'Calcula Valor Pago
    vrPago = vrValor + vrValorJaPago - vrGorjeta
    'Depende do Tipo
    Select Case strStatusLanc
    Case "RC"
      vrTotGR = CCur(IIf(Not IsNumeric(txtTotalaPagar.Text), 0, txtTotalaPagar.Text))
      'vrTotDescGR = CCur(IIf(Not IsNumeric(txtDesconto.Text), 0, txtDesconto.Text))
      'Validar Valor
      If vrPago < vrTotGR Then
        'Valor do pagamento < que valor a pagar
        strMsgAux = "" & vbCrLf
        strMsgAux = "Valor pago menor que valor a pagar" & vbCrLf & vbCrLf & _
          "Caso confirme, terá que fazer um novo lançamento para complementar o recebimento. Deseja continuar ?"
        If MsgBox(strMsgAux, vbYesNo, TITULOSISTEMA) = vbNo Then
          strMsg = "Valor pago não pode ser menor que valor a pagar" & vbCrLf
        End If
      ElseIf vrPago > vrTotGR Then
        'Valor do pagamento > que valor a pagar
        If gbTrabComTroco Then
          strMsgAux = "Atenção: Valor pago maior que valor a pagar" & vbCrLf & vbCrLf & _
            "Confirma diferença para Troco (R$ " & Format(vrPago - (vrTotGR - vrTotDescGR), "###,###,##0.00") & ")?"
        Else
          strMsgAux = "Atenção: Valor pago maior que valor a pagar" & vbCrLf & vbCrLf & _
            "Confirma diferença para Garçom (R$ " & Format(vrPago - (vrTotGR - vrTotDescGR), "###,###,##0.00") & ")?"
        End If
        If MsgBox(strMsgAux, vbYesNo, TITULOSISTEMA) = vbNo Then
          strMsg = "Valor pago não pode ser maior que valor a pagar" & vbCrLf
        Else
          If gbTrabComTroco Then
            'Caso trabalhe com troco, joga diferença para garçom
            vrCalcTroco = vrPago - (vrTotGR - vrTotDescGR)
          Else
            'Se não trabalha com troco joga diferença para Troco + O troco Digitado
            mskGorjeta.Text = Format((vrPago - (vrTotGR - vrTotDescGR)) + vrGorjeta, "###,###,##0.00")
          End If

        End If
      End If
    End Select
  End If
  If Len(strMsg) = 0 Then
    If blnFatura = True And Status = tpStatus_Alterar Then
      strMsgAlerta = "ATENÇÃO: " & vbCrLf & vbCrLf
      If cboTipoPagamento.Text = "Fatura" Then
        strMsgAlerta = strMsgAlerta & "As parcelas referentes a esta fatura serão regeradas." & vbCrLf & vbCrLf
      Else
        strMsgAlerta = strMsgAlerta & "A mudança do tipo de pagamento de [FATURA] para [" & UCase(cboTipoPagamento.Text) & "] implicará na exclusão das parcelas." & vbCrLf & vbCrLf
      End If
      strMsgAlerta = strMsgAlerta & "Deseja continuar ?"
      If MsgBox(strMsgAlerta, vbYesNo, TITULOSISTEMA) = vbNo Then
        strMsg = "Operação cancelada."
      End If
    End If
  End If
  Set objGeral = Nothing
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "ValidaContaCorrente"
    ValidaContaCorrente = False
  Else
    ValidaContaCorrente = True
  End If
  Exit Function
trata:
  ValidaContaCorrente = False
  TratarErro Err.Number, Err.Description, Err.Source
End Function

Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  '
  ValidaCampos = True
  ValidaCampos = ValidaContaCorrente
  If blnFecharContaCorrente Then
    blnFecharContaCorrente = False
    cmdCancelar_Click
    ValidaCampos = False
  End If
  Exit Function
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Function


Private Sub cmdParcela_Click()
  On Error GoTo trata
  If Not IsNumeric(grdGeral.Columns("ID").Value & "") Then
    MsgBox "Selecione um lançamento !", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  ElseIf strStatusLanc <> grdGeral.Columns("StatusLancamento").Value & "" Then
    MsgBox "Selecione um lançamento da mesma origem que " & IIf(strStatusLanc = "CC", "Conta Corrente", IIf(strStatusLanc = "RC", "Recebimento", IIf(strStatusLanc = "RE", "Recebimento Empresa", IIf(strStatusLanc = "DP", "Depósito", IIf(strStatusLanc = "DE", "Despesa", IIf(strStatusLanc = "VD", "Venda", IIf(strStatusLanc = "EX", "Extra", ""))))))), vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  ElseIf grdGeral.Columns("Tipo de Pagamento").Value & "" <> "Fatura" Then
    MsgBox "Selecione um lançamento do tipo Fatura", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If

  frmUserParcelaLis.lngCCID = grdGeral.Columns("ID").Value
  frmUserParcelaLis.strFuncionarioNome = strFuncionarioNome
  frmUserParcelaLis.Show vbModal
  '
  SetarFoco grdGeral

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    blnPrimeiraVez = False
    SetarFoco cboFuncionario
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserContaCorrente.Form_Activate]"
End Sub


Public Sub MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMed.clsGeral
  '
  On Error GoTo trata

  Set clsGer = New busSisMed.clsGeral
  '
  strSql = "SELECT CONTACORRENTE.PKID, CONTACORRENTE.STATUSLANCAMENTO, case STATUSCC when 'CC' then 'Cartão de Crédito' when 'CD' then 'Cartão de Débito' when 'ES' then 'Espécie' when 'PH' then 'Penhor' when 'CH' then 'Cheque' else 'Fatura' end , " & _
            " case STATUSLANCAMENTO when 'RC' then 'Recebimento' else '' end, case INDDEBITOCREDITO when 'D' then 'Débito' else 'Crédito' end ,CONTACORRENTE.DTHORACC, CONTACORRENTE.VALOR  " & _
            "FROM CONTACORRENTE " & _
            " WHERE CONTACORRENTE."
  Select Case strStatusLanc & ""
  Case "RC"
    strSql = strSql & "GRID"
  Case Else
    strSql = strSql & "" 'Para forçar erro
  End Select
  strSql = strSql & " = " & Formata_Dados(lngGRID, tpDados_Longo, tpNulo_Aceita)
  If strStatusLanc = "RC" Then
    'Recebimento
    strSql = strSql & " AND STATUSLANCAMENTO IN ('RC')"
  End If
  strSql = strSql & " ORDER BY CONTACORRENTE.PKID DESC;"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim Matriz(0 To COLUNASMATRIZ - 1, 0 To LINHASMATRIZ - 1)
  Else
    ReDim Matriz(0 To COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To COLUNASMATRIZ - 1  'varre as colunas
          Matriz(intJ, intI) = objRs(intJ) & ""
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

Private Sub Form_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  'Tratamento de tecla para verificação de chamada de Outras telas
  Select Case KeyAscii
  Case 6
    cmdOk_Click
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserContaCorrente.Form_KeyPress]"
End Sub

Private Sub Form_Load()
On Error GoTo trata
  Dim objRs     As ADODB.Recordset
  Dim strSql    As String
  Dim objCC     As busSisMed.clsContaCorrente
  '
  blnFechar = False
  blnRetorno = False
  blnFatura = False
  '
  Set objCC = New busSisMed.clsContaCorrente
  '
  AmpS
  Me.Height = 7005
  Me.Width = 11355
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, cmdExcluir, , cmdIncluir, cmdAlterar, , cmdImprimir
  LerFigurasAvulsas cmdCalculadora, "Cortesia.ico", "CortesiaDown.ico", "Resgatar Cheque"
  LerFigurasAvulsas cmdParcela, "Parcela.ico", "ParcelaDown.ico", "Parcelamento"
  '
  picListaPgto.Left = 150
  picListaPgto.Top = 2400
  picListaPgto.Height = 2775
  grdGeral.Height = 2700
  '
  'Limpar Campos
  LimparCampos
  '
  'Tratar Combos
  TratarCombos
  'Tratar Campos
  TratarCampos
  'Tratar Totais
  TratarTotais

  'If Status = tpStatus_Incluir Or Status = tpStatus_Consultar Then
  If Status = tpStatus_Incluir Or Status = tpStatus_Consultar Then
    cboTipoPagamento.Text = "Espécie"
    cboDebitoCredito.Text = "Crédito"
    INCLUIR_VALOR_NO_MASK mskData(1), Now, TpMaskData
    txtResponsavel.Text = gsNomeUsu

    If strStatusLanc = "CC" Then
    End If
    'Tratar Botões
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = True
    cmdExcluir.Enabled = True
    cmdCalculadora.Enabled = True
    If Status = tpStatus_Consultar Then
      cmdAlterar.Enabled = False
      cmdExcluir.Enabled = False
      cmdParcela.Enabled = False
      cmdCalculadora.Enabled = False
    End If
  ElseIf Status = tpStatus_Alterar Then
    'Pega Dados do Banco de dados
    Set objRs = objCC.SelecionarContaCorrente(lngCCID)
    '
    If Not objRs.EOF Then
      If objRs.Fields("DESC_STATUSCC").Value & "" = "Fatura" Then
        blnFatura = True
      End If
      cboTipoPagamento.Text = objRs.Fields("DESC_STATUSCC").Value
      cboDebitoCredito.Text = objRs.Fields("DESC_INDDEBCRED").Value
      If objRs.Fields("NOME_FUNC").Value & "" <> "" Then
        cboFuncionario.Text = objRs.Fields("NOME_FUNC").Value
      End If
      If objRs.Fields("NOME_CARTAO").Value & "" <> "" Then
        cboCartao.Text = objRs.Fields("NOME_CARTAO").Value
      End If
      If objRs.Fields("NOME_BANCO").Value & "" <> "" Then
        cboBanco.Text = objRs.Fields("NOME_BANCO").Value
      End If
      If objRs.Fields("NOME_CARTAODEBITO").Value & "" <> "" Then
        cboCartaoDebito.Text = objRs.Fields("NOME_CARTAODEBITO").Value
      End If

      INCLUIR_VALOR_NO_MASK mskData(1), objRs.Fields("DTHORACC").Value, TpMaskData
      txtResponsavel.Text = objRs.Fields("RESPONSAVEL").Value
      'txtTotalsDesc.Text = Format(IIf(Not IsNumeric(objRs.Fields("VRCALCTOTAL").Value), 0, objRs.Fields("VRCALCTOTAL").Value), "###,##0.00")
      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VALOR").Value, TpMaskMoeda
      'txtDesconto.Text = Format(IIf(Not IsNumeric(objRs.Fields("VRCALCDESCONTO").Value), 0, objRs.Fields("VRCALCDESCONTO").Value), "###,##0.00")
      'txtTotalaPagar.Text = Format(IIf(Not IsNumeric(objRs.Fields("VRCALCTOTAL").Value), 0, objRs.Fields("VRCALCTOTAL").Value) - IIf(Not IsNumeric(objRs.Fields("VRCALCDESCONTO").Value), 0, objRs.Fields("VRCALCDESCONTO").Value), "###,##0.00")
      INCLUIR_VALOR_NO_MASK mskGorjeta, objRs.Fields("VRGORJETA").Value, TpMaskMoeda
      txtTroco.Text = Format(IIf(Not IsNumeric(objRs.Fields("VRTROCO").Value), 0, objRs.Fields("VRTROCO").Value), "###,##0.00")
      txtLote.Text = objRs.Fields("LOTE").Value & ""
      mskCpf.Text = objRs.Fields("CPF").Value & ""
      txtNroCheque.Text = objRs.Fields("NROCHEQUE").Value & ""
      txtAgencia.Text = objRs.Fields("AGENCIA").Value & ""
      txtConta.Text = objRs.Fields("CONTA").Value & ""
      txtCliente.Text = objRs.Fields("CLIENTE").Value & ""
      txtDocumentoPenhor.Text = objRs.Fields("DOCUMENTOPENHOR").Value & ""
      txtObjeto.Text = objRs.Fields("DESCOBJETO").Value & ""
      lngTurnoRecebeId = IIf(Not IsNumeric(objRs.Fields("TURNOCCID").Value), 0, objRs.Fields("TURNOCCID").Value)
      INCLUIR_VALOR_NO_MASK mskNroParcelas, objRs.Fields("NROPARCELAS").Value, TpMaskLongo
      INCLUIR_VALOR_NO_MASK mskData(0), objRs.Fields("DTPRIMEIRAPARCELA").Value, TpMaskData
    End If
    objRs.Close
    Set objRs = Nothing
    '
    'Tratar Botões
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = False
    cmdExcluir.Enabled = False
    cmdCalculadora.Enabled = False
  End If
  'Capturar totais restantes
  CapturaTotais
  Set objCC = Nothing
  If Status = tpStatus_Consultar Then
    pictrava(4).Enabled = False
    pictrava(5).Enabled = False
    cmdOk.Enabled = False
    cmdCalculadora.Enabled = False

  End If
  'Carregar Grid
  COLUNASMATRIZ = grdGeral.Columns.Count
  LINHASMATRIZ = 0
  MontaMatriz
  grdGeral.Bookmark = Null
  grdGeral.ReBind
  grdGeral.ApproxCount = LINHASMATRIZ

  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo trata
  If Not blnFechar Then Cancel = True
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub


Private Sub grdGeral_UnboundReadDataEx( _
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
               Offset + intI, LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, COLUNASMATRIZ, LINHASMATRIZ, Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserContaCorrente.grdGeral_UnboundReadDataEx]"
End Sub



Private Sub mskCpf_GotFocus()
  Seleciona_Conteudo_Controle mskCpf
End Sub

Private Sub mskCpf_LostFocus()
  Pintar_Controle mskCpf, tpCorContr_Normal
'''  If StatusEdicao = MODOINSERIR Or StatusEdicao = MODOALTERAR Then
'''    If Not TestaCPF(mskCPF.Text) Then
'''      If Screen.ActiveControl.Tag <> "A" And Screen.ActiveControl.Tag <> "B" Then
'''        Call MsgBox("O CPF digitado não é válido !", vbExclamation, TITULOSISTEMA)
'''        Exit Sub
'''      End If
'''    End If
'''  End If
End Sub

Private Sub mskData_GotFocus(Index As Integer)
  Seleciona_Conteudo_Controle mskData(Index)
End Sub

Private Sub mskData_LostFocus(Index As Integer)
  Pintar_Controle mskData(Index), tpCorContr_Normal
End Sub

Private Sub mskGorjeta_GotFocus()
  Seleciona_Conteudo_Controle mskGorjeta
End Sub

Private Sub mskGorjeta_LostFocus()
  Pintar_Controle mskGorjeta, tpCorContr_Normal
End Sub


Private Sub mskValor_GotFocus()
  Seleciona_Conteudo_Controle mskValor
End Sub

Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub
Private Sub mskNroParcelas_GotFocus()
  Seleciona_Conteudo_Controle mskNroParcelas
End Sub
Private Sub mskNroParcelas_LostFocus()
  Pintar_Controle mskNroParcelas, tpCorContr_Normal
End Sub

Private Sub txtAgencia_GotFocus()
  Seleciona_Conteudo_Controle txtAgencia
End Sub

Private Sub txtAgencia_LostFocus()
  Pintar_Controle txtAgencia, tpCorContr_Normal
End Sub

Private Sub txtCliente_GotFocus()
  Seleciona_Conteudo_Controle txtCliente
End Sub

Private Sub txtCliente_LostFocus()
  Pintar_Controle txtCliente, tpCorContr_Normal
End Sub

Private Sub txtConta_GotFocus()
  Seleciona_Conteudo_Controle txtConta
End Sub

Private Sub txtConta_LostFocus()
  Pintar_Controle txtConta, tpCorContr_Normal
End Sub
Private Sub txtDocumentoPenhor_GotFocus()
  Seleciona_Conteudo_Controle txtDocumentoPenhor
End Sub

Private Sub txtDocumentoPenhor_LostFocus()
  Pintar_Controle txtDocumentoPenhor, tpCorContr_Normal
End Sub

Private Sub txtNroCheque_GotFocus()
  Seleciona_Conteudo_Controle txtNroCheque
End Sub

Private Sub txtNroCheque_LostFocus()
  Pintar_Controle txtNroCheque, tpCorContr_Normal
End Sub

Private Sub txtObjeto_GotFocus()
  Seleciona_Conteudo_Controle txtObjeto
End Sub

Private Sub txtObjeto_LostFocus()
  Pintar_Controle txtObjeto, tpCorContr_Normal
End Sub

