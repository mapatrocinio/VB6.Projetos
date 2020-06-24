VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmInsumoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de "
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5610
      Left            =   8430
      ScaleHeight     =   5610
      ScaleWidth      =   1860
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2925
         Left            =   90
         ScaleHeight     =   2865
         ScaleWidth      =   1605
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2490
         Width           =   1665
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   150
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1020
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   1860
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5295
      Left            =   120
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userInsumoInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Complemento"
      TabPicture(1)   =   "userInsumoInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Imagem"
      TabPicture(2)   =   "userInsumoInc.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Preço por filial"
      TabPicture(3)   =   "userInsumoInc.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grdPrecoFilial"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame3 
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
         Height          =   4695
         Left            =   -74850
         TabIndex        =   83
         Top             =   360
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   4365
            Index           =   5
            Left            =   120
            ScaleHeight     =   4365
            ScaleWidth      =   7575
            TabIndex        =   84
            Top             =   210
            Width           =   7575
            Begin VB.Image Image1 
               Height          =   4305
               Left            =   30
               Stretch         =   -1  'True
               Top             =   30
               Width           =   7515
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
         Height          =   4695
         Left            =   -74850
         TabIndex        =   79
         Top             =   360
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2295
            Index           =   4
            Left            =   120
            ScaleHeight     =   2295
            ScaleWidth      =   7575
            TabIndex        =   80
            Top             =   210
            Width           =   7575
            Begin VB.ComboBox cboIPI 
               Height          =   315
               Left            =   5700
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   60
               Width           =   1695
            End
            Begin VB.ComboBox cboICMS 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   60
               Width           =   1695
            End
            Begin MSMask.MaskEdBox mskEstMinimo 
               Height          =   255
               Left            =   1320
               TabIndex        =   19
               Top             =   720
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.0000;($#,##0.0000)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskFinancVenda 
               Height          =   255
               Left            =   1320
               TabIndex        =   18
               Top             =   420
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskMargemEst 
               Height          =   255
               Left            =   5700
               TabIndex        =   20
               Top             =   720
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskSaldoEst 
               Height          =   255
               Left            =   1320
               TabIndex        =   21
               Top             =   1020
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskMargemAjuste 
               Height          =   255
               Left            =   1320
               TabIndex        =   23
               Top             =   1320
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskPrecoVenda 
               Height          =   255
               Left            =   5700
               TabIndex        =   24
               Top             =   1320
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskCustoProduto 
               Height          =   255
               Left            =   5700
               TabIndex        =   22
               Top             =   1020
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskTAM 
               Height          =   255
               Left            =   1320
               TabIndex        =   25
               Top             =   1620
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.0000;($#,##0.0000)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskPAD 
               Height          =   255
               Left            =   5700
               TabIndex        =   26
               Top             =   1620
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskSOB 
               Height          =   255
               Left            =   1320
               TabIndex        =   27
               Top             =   1920
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label23 
               Caption         =   "SOB"
               Height          =   255
               Left            =   90
               TabIndex        =   94
               Top             =   1920
               Width           =   1095
            End
            Begin VB.Label Label20 
               Caption         =   "PAD"
               Height          =   255
               Left            =   4470
               TabIndex        =   93
               Top             =   1620
               Width           =   1095
            End
            Begin VB.Label Label9 
               Caption         =   "TAM"
               Height          =   255
               Left            =   90
               TabIndex        =   92
               Top             =   1620
               Width           =   1095
            End
            Begin VB.Label Label8 
               Caption         =   "Custo Produto"
               Height          =   255
               Left            =   4470
               TabIndex        =   91
               Top             =   1020
               Width           =   1095
            End
            Begin VB.Label Label22 
               Caption         =   "Preço Venda"
               Height          =   255
               Left            =   4470
               TabIndex        =   90
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label Label21 
               Caption         =   "Margem Ajuste"
               Height          =   255
               Left            =   90
               TabIndex        =   89
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label Label19 
               Caption         =   "Saldo Est."
               Height          =   255
               Left            =   90
               TabIndex        =   88
               Top             =   1020
               Width           =   1095
            End
            Begin VB.Label Label18 
               Caption         =   "Margem Est."
               Height          =   255
               Left            =   4470
               TabIndex        =   87
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label17 
               Caption         =   "Financ. Venda"
               Height          =   255
               Left            =   90
               TabIndex        =   86
               Top             =   420
               Width           =   1095
            End
            Begin VB.Label Label16 
               Caption         =   "Perc. IPI"
               Height          =   255
               Left            =   4470
               TabIndex        =   85
               Top             =   120
               Width           =   1095
            End
            Begin VB.Label Label11 
               Caption         =   "Perc. ICMS"
               Height          =   255
               Left            =   90
               TabIndex        =   82
               Top             =   120
               Width           =   1095
            End
            Begin VB.Label Label10 
               Caption         =   "Est. Mínimo"
               Height          =   255
               Left            =   90
               TabIndex        =   81
               Top             =   720
               Width           =   1095
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
         Height          =   4755
         Left            =   150
         TabIndex        =   47
         Top             =   360
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   1755
            Index           =   1
            Left            =   120
            ScaleHeight     =   1755
            ScaleWidth      =   7575
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   1770
            Width           =   7575
            Begin VB.TextBox txtCodigo 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   28
               Text            =   "txtCodigo"
               Top             =   30
               Width           =   5865
            End
            Begin VB.TextBox txtLinhaFim 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   3690
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   30
               TabStop         =   0   'False
               Text            =   "txtLinhaFim"
               Top             =   390
               Width           =   3495
            End
            Begin VB.TextBox txtCodigoFim 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   29
               TabStop         =   0   'False
               Text            =   "txtCodigoFim"
               Top             =   390
               Width           =   2355
            End
            Begin VB.ComboBox cboCor 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   720
               Width           =   3855
            End
            Begin MSMask.MaskEdBox mskPesoMinimo 
               Height          =   255
               Left            =   5700
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   1080
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   "#,##0.0000;($#,##0.0000)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskPesoEstoque 
               Height          =   255
               Left            =   5700
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   1380
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   "#,##0.0000;($#,##0.0000)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskQtdMinPerfil 
               Height          =   255
               Left            =   1320
               TabIndex        =   32
               Top             =   1050
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskQtdEstPerfil 
               Height          =   255
               Left            =   1320
               TabIndex        =   33
               Top             =   1410
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Qtd Estoque"
               Height          =   195
               Index           =   8
               Left            =   90
               TabIndex        =   64
               Top             =   1380
               Width           =   1155
            End
            Begin VB.Label Label5 
               Caption         =   "Qtd Mínima"
               Height          =   195
               Index           =   7
               Left            =   90
               TabIndex        =   63
               Top             =   1080
               Width           =   945
            End
            Begin VB.Label Label1 
               Caption         =   "Nome da Linha/Código Perfil"
               Height          =   615
               Index           =   0
               Left            =   120
               TabIndex        =   62
               Top             =   60
               Width           =   1095
            End
            Begin VB.Label Label5 
               Caption         =   "Peso Mínimo"
               ForeColor       =   &H80000011&
               Height          =   195
               Index           =   6
               Left            =   4470
               TabIndex        =   61
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Peso Estoque"
               ForeColor       =   &H80000011&
               Height          =   195
               Index           =   3
               Left            =   4470
               TabIndex        =   60
               Top             =   1380
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cor"
               Height          =   195
               Index           =   2
               Left            =   90
               TabIndex        =   57
               Top             =   750
               Width           =   1215
            End
         End
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   1965
            Index           =   2
            Left            =   240
            ScaleHeight     =   1965
            ScaleWidth      =   7575
            TabIndex        =   53
            Top             =   4560
            Width           =   7575
            Begin VB.ComboBox cboEmbalagem 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   420
               Width           =   3855
            End
            Begin VB.ComboBox cboGrupo 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   38
               Top             =   750
               Width           =   3855
            End
            Begin VB.TextBox txtNome 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   36
               Text            =   "txtNome"
               Top             =   90
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskValor 
               Height          =   255
               Left            =   1320
               TabIndex        =   39
               Top             =   1080
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskQtdMinima 
               Height          =   255
               Left            =   1320
               TabIndex        =   40
               Top             =   1380
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskQtdEstoque 
               Height          =   255
               Left            =   1320
               TabIndex        =   41
               Top             =   1680
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Embalagem"
               Height          =   195
               Index           =   10
               Left            =   90
               TabIndex        =   66
               Top             =   450
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Grupo"
               Height          =   195
               Index           =   9
               Left            =   90
               TabIndex        =   65
               Top             =   780
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Qtd. Mínima"
               Height          =   195
               Index           =   33
               Left            =   60
               TabIndex        =   59
               Top             =   1380
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Qtd. Estoque"
               Height          =   195
               Index           =   4
               Left            =   60
               TabIndex        =   58
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Nome"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   56
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label Label4 
               Caption         =   "Valor"
               Height          =   255
               Left            =   90
               TabIndex        =   54
               Top             =   1080
               Width           =   1095
            End
         End
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   3585
            Index           =   3
            Left            =   120
            ScaleHeight     =   3585
            ScaleWidth      =   7575
            TabIndex        =   67
            Top             =   960
            Width           =   7575
            Begin VB.CheckBox chkComissaoVendedor 
               Caption         =   "Comissão Vendedor?"
               Height          =   375
               Left            =   5610
               TabIndex        =   15
               Top             =   3030
               Width           =   1875
            End
            Begin VB.ComboBox cboFamilia 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   3060
               Width           =   3855
            End
            Begin VB.TextBox txtTabela 
               Height          =   285
               Left            =   1320
               MaxLength       =   20
               TabIndex        =   13
               Text            =   "txtTabela"
               Top             =   2730
               Width           =   6075
            End
            Begin VB.TextBox txtModRef 
               Height          =   285
               Left            =   1320
               MaxLength       =   20
               TabIndex        =   12
               Text            =   "txtModRef"
               Top             =   2400
               Width           =   6075
            End
            Begin VB.ComboBox cboFornecedor 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   1440
               Width           =   3855
            End
            Begin VB.TextBox txtNomeAbrevProduto 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   4
               Text            =   "txtNomeAbrevProduto"
               Top             =   420
               Width           =   6075
            End
            Begin VB.TextBox txtNomeProduto 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   3
               Text            =   "txtNomeProduto"
               Top             =   90
               Width           =   6075
            End
            Begin VB.ComboBox cboGrupoProduto 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   1080
               Width           =   3855
            End
            Begin VB.ComboBox cboEmbalgemProduto 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   750
               Width           =   3855
            End
            Begin MSMask.MaskEdBox mskValorProduto 
               Height          =   255
               Left            =   1320
               TabIndex        =   8
               Top             =   1800
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskAltura 
               Height          =   255
               Left            =   1320
               TabIndex        =   10
               Top             =   2100
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.0000;($#,##0.0000)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskLargura 
               Height          =   255
               Left            =   5700
               TabIndex        =   11
               Top             =   2100
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.0000;($#,##0.0000)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskPeso 
               Height          =   255
               Left            =   5700
               TabIndex        =   9
               Top             =   1800
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.0000;($#,##0.0000)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Família"
               Height          =   195
               Index           =   18
               Left            =   90
               TabIndex        =   95
               Top             =   3090
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Tabela"
               Height          =   195
               Index           =   14
               Left            =   90
               TabIndex        =   78
               Top             =   2775
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Mod. Ref."
               Height          =   195
               Index           =   13
               Left            =   90
               TabIndex        =   77
               Top             =   2445
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Peso"
               Height          =   255
               Left            =   4470
               TabIndex        =   76
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label Label6 
               Caption         =   "Largura"
               Height          =   255
               Left            =   4470
               TabIndex        =   75
               Top             =   2100
               Width           =   1095
            End
            Begin VB.Label Label3 
               Caption         =   "Altura"
               Height          =   255
               Left            =   90
               TabIndex        =   74
               Top             =   2100
               Width           =   1095
            End
            Begin VB.Label Label5 
               Caption         =   "Fornecedor"
               Height          =   195
               Index           =   17
               Left            =   90
               TabIndex        =   73
               Top             =   1470
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Nome Abrev."
               Height          =   195
               Index           =   16
               Left            =   90
               TabIndex        =   72
               Top             =   465
               Width           =   1215
            End
            Begin VB.Label Label2 
               Caption         =   "Valor"
               Height          =   255
               Left            =   90
               TabIndex        =   71
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label Label5 
               Caption         =   "Nome"
               Height          =   195
               Index           =   15
               Left            =   90
               TabIndex        =   70
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Grupo"
               Height          =   195
               Index           =   12
               Left            =   90
               TabIndex        =   69
               Top             =   1110
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Unid. Medida"
               Height          =   195
               Index           =   11
               Left            =   90
               TabIndex        =   68
               Top             =   780
               Width           =   1215
            End
         End
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   735
            Index           =   0
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   7575
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   210
            Width           =   7575
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1290
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   450
               Width           =   2235
               Begin VB.OptionButton optStatus 
                  Caption         =   "Inativo"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   2
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton optStatus 
                  Caption         =   "Ativo"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   1
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   825
               End
            End
            Begin VB.TextBox txtCodigoInsumo 
               Height          =   285
               Left            =   1320
               MaxLength       =   30
               TabIndex        =   0
               Text            =   "txtCodigoInsumo"
               Top             =   75
               Width           =   2925
            End
            Begin VB.Label Label5 
               Caption         =   "Status"
               Height          =   315
               Index           =   5
               Left            =   60
               TabIndex        =   50
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Código"
               ForeColor       =   &H80000011&
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   49
               Top             =   120
               Width           =   1215
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdPrecoFilial 
         Height          =   4725
         Left            =   -74910
         OleObjectBlob   =   "userInsumoInc.frx":0070
         TabIndex        =   96
         Top             =   390
         Width           =   7905
      End
   End
End
Attribute VB_Name = "frmInsumoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intTipoInsumo            As tpInsumo

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long

Private blnPrimeiraVez          As Boolean

Dim PRFIL_COLUNASMATRIZ        As Long
Dim PRFIL_LINHASMATRIZ         As Long
Private PRFIL_Matriz()         As String

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Insumo
  LimparCampoTexto txtCodigoInsumo
  LimparCampoOption optStatus
  'Perfil
  LimparCampoTexto txtCodigo
  LimparCampoTexto txtCodigoFim
  LimparCampoTexto txtLinhaFim
  LimparCampoCombo cboCor
  LimparCampoCombo cboGrupo
  LimparCampoCombo cboEmbalagem
  LimparCampoMask mskPesoMinimo
  LimparCampoMask mskPesoEstoque
  'Acessorio
  LimparCampoTexto txtNome
  LimparCampoMask mskValor
  LimparCampoMask mskQtdMinima
  LimparCampoMask mskQtdEstoque
  '
  'Produto
  LimparCampoTexto txtNomeProduto
  LimparCampoTexto txtNomeAbrevProduto
  LimparCampoCombo cboEmbalgemProduto
  LimparCampoCombo cboGrupoProduto
  LimparCampoCombo cboFornecedor
  LimparCampoMask mskValorProduto
  LimparCampoMask mskPeso
  LimparCampoMask mskAltura
  LimparCampoMask mskLargura
  LimparCampoTexto txtModRef
  LimparCampoTexto txtTabela
  LimparCampoCombo cboFamilia
  LimparCampoCombo cboICMS
  LimparCampoCombo cboIPI
  LimparCampoMask mskFinancVenda
  LimparCampoMask mskEstMinimo
  LimparCampoMask mskMargemEst
  LimparCampoMask mskSaldoEst
  LimparCampoMask mskCustoProduto
  LimparCampoMask mskMargemAjuste
  LimparCampoMask mskPrecoVenda
  LimparCampoMask mskTAM
  LimparCampoMask mskPAD
  LimparCampoMask mskSOB
  LimparCampoCheck chkComissaoVendedor
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserInsumoInc.LimparCampos]", _
            Err.Description
End Sub


Private Sub cboCor_LostFocus()
  Pintar_Controle cboCor, tpCorContr_Normal
End Sub


Private Sub cboEmbalagem_LostFocus()
  Pintar_Controle cboEmbalagem, tpCorContr_Normal
End Sub

Private Sub cboEmbalgemProduto_LostFocus()
  Pintar_Controle cboEmbalgemProduto, tpCorContr_Normal
End Sub

Private Sub cboFamilia_LostFocus()
  Pintar_Controle cboFamilia, tpCorContr_Normal
End Sub

Private Sub cboFornecedor_LostFocus()
  Pintar_Controle cboFornecedor, tpCorContr_Normal
End Sub

Private Sub cboGrupo_LostFocus()
  Pintar_Controle cboGrupo, tpCorContr_Normal
End Sub

Private Sub cboGrupoProduto_LostFocus()
  Pintar_Controle cboGrupoProduto, tpCorContr_Normal
End Sub

Private Sub cboICMS_LostFocus()
  Pintar_Controle cboICMS, tpCorContr_Normal
End Sub

Private Sub cboIPI_LostFocus()
  Pintar_Controle cboIPI, tpCorContr_Normal
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

Private Sub cmdOk_Click()
  Dim objInsumo                 As busSisMetal.clsInsumo
  Dim objItemPedido             As busSisMetal.clsItemPedido
  Dim objGeral                  As busSisMetal.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim strStatus                 As String
  Dim lngLINHAID                As Long
  Dim lngCORID                  As Long
  Dim lngGRUPOID                As Long
  Dim lngEMBALAGEMID            As Long
  '
  Dim lngEMBALAGEMPRODUTOID     As Long
  Dim lngGRUPOPRODUTOID         As Long
  Dim lngFORNECEDORID           As Long
  Dim lngFAMILIAPRODUTOID       As Long
  Dim lngIPIID                  As Long
  Dim lngICMSID                 As Long
  '
  Dim strCODIGOINSUMO           As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busSisMetal.clsGeral
  Set objInsumo = New busSisMetal.clsInsumo
  '
  'Status
  If optStatus(0).Value Then
    strStatus = "A"
  Else
    strStatus = "I"
  End If
  'OBETR CÓDIGO COMBOS E CÓDIGOS
  strCODIGOINSUMO = txtCodigoFim.Text
  'LINHA
  lngLINHAID = 0
  strSql = "SELECT LINHA.PKID FROM LINHA " & _
      " INNER JOIN TIPO_LINHA ON TIPO_LINHA.PKID = LINHA.TIPO_LINHAID  " & _
      " WHERE LINHA.CODIGO = " & Formata_Dados(txtCodigoFim.Text, tpDados_Texto) & _
      " AND TIPO_LINHA.NOME = " & Formata_Dados(txtLinhaFim.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngLINHAID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'COR
  lngCORID = 0
  strSql = "SELECT PKID, SIGLA FROM COR WHERE NOME = " & Formata_Dados(cboCor.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngCORID = objRs.Fields("PKID").Value
    strCODIGOINSUMO = strCODIGOINSUMO & objRs.Fields("SIGLA").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'GRUPO
  lngGRUPOID = 0
  strSql = "SELECT PKID, NOME FROM GRUPO WHERE NOME = " & Formata_Dados(cboGrupo.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngGRUPOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'EMBALAGEM
  lngEMBALAGEMID = 0
  strSql = "SELECT PKID, NOME FROM EMBALAGEM WHERE NOME = " & Formata_Dados(cboEmbalagem.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngEMBALAGEMID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'EMBALAGEM PRODUTO
  lngEMBALAGEMPRODUTOID = 0
  strSql = "SELECT PKID, NOME FROM EMBALAGEM WHERE NOME = " & Formata_Dados(cboEmbalgemProduto.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngEMBALAGEMPRODUTOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'GRUPO PRODUTO
  lngGRUPOPRODUTOID = 0
  strSql = "SELECT PKID, NOME FROM GRUPO_PRODUTO WHERE NOME = " & Formata_Dados(cboGrupoProduto.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngGRUPOPRODUTOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'FORNECEDOR
  lngFORNECEDORID = 0
  strSql = "SELECT PKID FROM LOJA WHERE NOME = " & Formata_Dados(cboFornecedor.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngFORNECEDORID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'FAMILIA PRODUTO
  lngFAMILIAPRODUTOID = 0
  strSql = "SELECT PKID FROM FAMILIAPRODUTOS WHERE DESCRICAO = " & Formata_Dados(cboFamilia.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngFAMILIAPRODUTOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'IPI
  lngIPIID = 0
  strSql = "SELECT PKID FROM IPI WHERE IPI = " & Formata_Dados(cboIPI.Text, tpDados_Moeda)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngIPIID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'ICMS
  lngICMSID = 0
  strSql = "SELECT PKID FROM ICMS WHERE ICMS = " & Formata_Dados(cboICMS.Text, tpDados_Moeda)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngICMSID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'VALIDAÇÕES DIFERENTES PARA ACESSORIO E PERFIL
  Select Case intTipoInsumo
  Case tpInsumo_Perfil
    '
    'Validar se PERFIL JÁ CADASTRADO
    strSql = "SELECT * FROM INSUMO " & _
      " INNER JOIN PERFIL ON INSUMO.PKID = PERFIL.INSUMOID " & _
      " WHERE PERFIL.LINHAID = " & Formata_Dados(lngLINHAID, tpDados_Longo) & _
      " AND PERFIL.CORID = " & Formata_Dados(lngCORID, tpDados_Longo) & _
      " AND INSUMO.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      TratarErroPrevisto "Linha-perfil/cor já cadastrado"
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      Set objInsumo = Nothing
      cmdOk.Enabled = True
      SetarFoco txtCodigo
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
  Case tpInsumo_Acessorio
    'Validar se ACESSÓRIO JÁ CADASTRADO
    strSql = "SELECT * FROM INSUMO " & _
      " INNER JOIN ACESSORIO ON INSUMO.PKID = ACESSORIO.INSUMOID " & _
      " WHERE (ACESSORIO.NOME = " & Formata_Dados(txtNome.Text, tpDados_Texto) & _
      " OR INSUMO.CODIGO = " & Formata_Dados(txtCodigoInsumo.Text, tpDados_Texto) & ") " & _
      " AND INSUMO.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      Pintar_Controle txtNome, tpCorContr_Erro
      TratarErroPrevisto "Nome ou código já cadastrado para acessório"
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      Set objInsumo = Nothing
      cmdOk.Enabled = True
      SetarFoco txtNome
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    '
    strCODIGOINSUMO = txtCodigoInsumo.Text
  Case tpInsumo_Produto
    'Validar se PRODUTO JÁ CADASTRADO
    strSql = "SELECT * FROM INSUMO " & _
      " INNER JOIN PRODUTO ON INSUMO.PKID = PRODUTO.INSUMOID " & _
      " WHERE (PRODUTO.NOME = " & Formata_Dados(txtNome.Text, tpDados_Texto) & _
      " OR INSUMO.CODIGO = " & Formata_Dados(txtCodigoInsumo.Text, tpDados_Texto) & ") " & _
      " AND INSUMO.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      Pintar_Controle txtNome, tpCorContr_Erro
      TratarErroPrevisto "Nome ou código já cadastrado para o produto"
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      Set objInsumo = Nothing
      cmdOk.Enabled = True
      SetarFoco txtNome
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    '
    strCODIGOINSUMO = txtCodigoInsumo.Text
  End Select
  
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar Insumo
    objInsumo.AlterarInsumo lngPKID, _
                            strCODIGOINSUMO, _
                            strStatus
    '
    Set objItemPedido = New busSisMetal.clsItemPedido
    '
    Select Case intTipoInsumo
    Case tpInsumo_Perfil: objInsumo.AlterarPerfil lngPKID, _
                                                  lngLINHAID, _
                                                  lngCORID, _
                                                  objItemPedido.CalculoPesoPedido(lngLINHAID, Format(IIf(Len(mskQtdMinPerfil.ClipText) = 0, "", mskQtdMinPerfil.Text), "###,##0")), _
                                                  objItemPedido.CalculoPesoPedido(lngLINHAID, Format(IIf(Len(mskQtdEstPerfil.ClipText) = 0, "", mskQtdEstPerfil.Text), "###,##0"))
    Case tpInsumo_Acessorio: objInsumo.AlterarAcessorio lngPKID, _
                                                        lngGRUPOID, _
                                                        lngEMBALAGEMID, _
                                                        txtNome.Text, _
                                                        IIf(Len(mskValor.ClipText) = 0, "", mskValor.Text), _
                                                        IIf(Len(mskQtdMinima.ClipText) = 0, "", mskQtdMinima.Text), _
                                                        IIf(Len(mskQtdEstoque.ClipText) = 0, "", mskQtdEstoque.Text)
    Case tpInsumo_Produto: objInsumo.AlterarProduto lngPKID, _
                                                    lngGRUPOPRODUTOID, _
                                                    lngEMBALAGEMPRODUTOID, _
                                                    lngFORNECEDORID, _
                                                    lngFAMILIAPRODUTOID, _
                                                    lngIPIID, _
                                                    lngICMSID, _
                                                    txtNomeProduto.Text, _
                                                    txtNomeAbrevProduto.Text, _
                                                    IIf(Len(mskValorProduto.ClipText) = 0, "", mskValorProduto.Text), _
                                                    IIf(Len(mskPeso.ClipText) = 0, "", mskPeso.Text), _
                                                    IIf(Len(mskAltura.ClipText) = 0, "", mskAltura.Text), _
                                                    IIf(Len(mskLargura.ClipText) = 0, "", mskLargura.Text), _
                                                    txtModRef.Text, _
                                                    txtTabela.Text, _
                                                    IIf(Len(mskFinancVenda.ClipText) = 0, "", mskFinancVenda.Text), _
                                                    IIf(Len(mskEstMinimo.ClipText) = 0, "", mskEstMinimo.Text), _
                                                    IIf(Len(mskMargemEst.ClipText) = 0, "", mskMargemEst.Text), _
                                                    IIf(Len(mskSaldoEst.ClipText) = 0, "", mskSaldoEst.Text), _
                                                    IIf(Len(mskCustoProduto.ClipText) = 0, "", mskCustoProduto.Text), _
                                                    IIf(Len(mskMargemAjuste.ClipText) = 0, "", mskMargemAjuste.Text), _
                                                    IIf(Len(mskPrecoVenda.ClipText) = 0, "", mskPrecoVenda.Text), _
                                                    IIf(Len(mskTAM.ClipText) = 0, "", mskTAM.Text), _
                                                    IIf(Len(mskPAD.ClipText) = 0, "", mskPAD.Text), _
                                                    IIf(Len(mskSOB.ClipText) = 0, "", mskSOB.Text), chkComissaoVendedor.Value

    End Select
    '
    Set objItemPedido = Nothing
    '
    blnRetorno = True
    blnFechar = True
    Unload Me
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir Insumo
    objInsumo.InserirInsumo lngPKID, _
                            strCODIGOINSUMO, _
                            strStatus
    '
    Set objItemPedido = New busSisMetal.clsItemPedido
    '
    Select Case intTipoInsumo
    Case tpInsumo_Perfil: objInsumo.InserirPerfil lngPKID, _
                                                  lngLINHAID, _
                                                  lngCORID, _
                                                  objItemPedido.CalculoPesoPedido(lngLINHAID, Format(IIf(Len(mskQtdMinPerfil.ClipText) = 0, "", mskQtdMinPerfil.Text), "###,##0")), _
                                                  objItemPedido.CalculoPesoPedido(lngLINHAID, Format(IIf(Len(mskQtdEstPerfil.ClipText) = 0, "", mskQtdEstPerfil.Text), "###,##0"))
    Case tpInsumo_Acessorio: objInsumo.InserirAcessorio lngPKID, _
                                                        lngGRUPOID, _
                                                        lngEMBALAGEMID, _
                                                        txtNome.Text, _
                                                        IIf(Len(mskValor.ClipText) = 0, "", mskValor.Text), _
                                                        IIf(Len(mskQtdMinima.ClipText) = 0, "", mskQtdMinima.Text), _
                                                        IIf(Len(mskQtdEstoque.ClipText) = 0, "", mskQtdEstoque.Text)
    Case tpInsumo_Produto: objInsumo.InserirProduto lngPKID, _
                                                    lngGRUPOPRODUTOID, _
                                                    lngEMBALAGEMPRODUTOID, _
                                                    lngFORNECEDORID, _
                                                    lngFAMILIAPRODUTOID, _
                                                    lngIPIID, _
                                                    lngICMSID, _
                                                    txtNomeProduto.Text, _
                                                    txtNomeAbrevProduto.Text, _
                                                    IIf(Len(mskValorProduto.ClipText) = 0, "", mskValorProduto.Text), _
                                                    IIf(Len(mskPeso.ClipText) = 0, "", mskPeso.Text), _
                                                    IIf(Len(mskAltura.ClipText) = 0, "", mskAltura.Text), _
                                                    IIf(Len(mskLargura.ClipText) = 0, "", mskLargura.Text), _
                                                    txtModRef.Text, _
                                                    txtTabela.Text, _
                                                    IIf(Len(mskFinancVenda.ClipText) = 0, "", mskFinancVenda.Text), _
                                                    IIf(Len(mskEstMinimo.ClipText) = 0, "", mskEstMinimo.Text), _
                                                    IIf(Len(mskMargemEst.ClipText) = 0, "", mskMargemEst.Text), _
                                                    IIf(Len(mskSaldoEst.ClipText) = 0, "", mskSaldoEst.Text), _
                                                    IIf(Len(mskCustoProduto.ClipText) = 0, "", mskCustoProduto.Text), _
                                                    IIf(Len(mskMargemAjuste.ClipText) = 0, "", mskMargemAjuste.Text), _
                                                    IIf(Len(mskPrecoVenda.ClipText) = 0, "", mskPrecoVenda.Text), _
                                                    IIf(Len(mskTAM.ClipText) = 0, "", mskTAM.Text), _
                                                    IIf(Len(mskPAD.ClipText) = 0, "", mskPAD.Text), _
                                                    IIf(Len(mskSOB.ClipText) = 0, "", mskSOB.Text), chkComissaoVendedor.Value

    End Select
    '
    Set objItemPedido = Nothing
    '
    If Status = tpStatus_Alterar Then
      blnRetorno = True
      blnFechar = True
      Unload Me
    ElseIf Status = tpStatus_Incluir Then
      'Selecionar prontuario pelo nome
      Status = tpStatus_Alterar
      'Reload na tela
      Form_Load
      'Acerta tabs
      Select Case intTipoInsumo
      Case tpInsumo_Produto
       tabDetalhes.Tab = 3
      Case Else
        tabDetalhes.Tab = 0
      End Select
      blnRetorno = True
    End If
    
  End If
  Set objInsumo = Nothing
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
  If Not Valida_Option(optStatus, blnSetarFocoControle) Then
    strMsg = strMsg & "Slecionar o status" & vbCrLf
  End If
  '
  Select Case intTipoInsumo
  Case tpInsumo_Perfil
    If txtCodigoFim.Text = "" Or txtLinhaFim.Text = "" Then
      strMsg = strMsg & "Selecionar a linha" & vbCrLf
      Pintar_Controle txtCodigo, tpCorContr_Erro
      SetarFoco txtCodigo
      blnSetarFocoControle = False
    End If
    If Not Valida_String(cboCor, TpObrigatorio) Then
      strMsg = strMsg & "Selecionar a cor" & vbCrLf
    End If
    If Not Valida_Moeda(mskQtdMinPerfil, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar a quantidade mínima em estoque válida" & vbCrLf
    End If
    If Not Valida_Moeda(mskQtdEstPerfil, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar a quantidade em estoque válida" & vbCrLf
    End If
  Case tpInsumo_Acessorio
    If Not Valida_String(txtCodigoInsumo, TpObrigatorio) Then
      strMsg = strMsg & "Informar o código válido" & vbCrLf
    End If
    If Not Valida_String(txtNome, TpObrigatorio) Then
      strMsg = strMsg & "Informar o nome válido" & vbCrLf
    End If
    If Not Valida_String(cboGrupo, TpObrigatorio) Then
      strMsg = strMsg & "Selecionar o grupo" & vbCrLf
    End If
    If Not Valida_String(cboEmbalagem, TpObrigatorio) Then
      strMsg = strMsg & "Selecionar a embalagem" & vbCrLf
    End If
    If Not Valida_Moeda(mskValor, TpObrigatorio) Then
      strMsg = strMsg & "Informar o valor válido" & vbCrLf
      Pintar_Controle mskValor, tpCorContr_Erro
    End If
    If Not Valida_Moeda(mskQtdMinima, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar a quantidade mínima em estoque válida" & vbCrLf
    End If
    If Not Valida_Moeda(mskQtdEstoque, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar a quantidade em estoque válida" & vbCrLf
    End If
  Case tpInsumo_Produto
    If Not Valida_String(txtCodigoInsumo, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar o código válido" & vbCrLf
    End If
    If Not Valida_String(txtNomeProduto, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar o nome válido" & vbCrLf
    End If
    If Not Valida_Moeda(mskValorProduto, TpNaoObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar o valor do produto válido" & vbCrLf
      Pintar_Controle mskValorProduto, tpCorContr_Erro
    End If
    If Not Valida_Moeda(mskPeso, TpNaoObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar o peso do produto válido" & vbCrLf
      Pintar_Controle mskPeso, tpCorContr_Erro
    End If
    If Not Valida_Moeda(mskAltura, TpNaoObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar a altura do produto válida" & vbCrLf
      Pintar_Controle mskAltura, tpCorContr_Erro
    End If
    If Not Valida_Moeda(mskLargura, TpNaoObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar a largura do produto válida" & vbCrLf
      Pintar_Controle mskLargura, tpCorContr_Erro
    End If
    If Not Valida_Moeda(mskFinancVenda, TpNaoObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar o financ. venda do produto válido" & vbCrLf
      Pintar_Controle mskFinancVenda, tpCorContr_Erro
    End If
    If Not Valida_Moeda(mskEstMinimo, TpNaoObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar o estoque mínimo do produto válido" & vbCrLf
      Pintar_Controle mskEstMinimo, tpCorContr_Erro
    End If
    If Not Valida_Moeda(mskMargemEst, TpNaoObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar a margem do estoque do produto válida" & vbCrLf
      Pintar_Controle mskMargemEst, tpCorContr_Erro
    End If
    If Not Valida_Moeda(mskSaldoEst, TpNaoObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar o saldo em estoque do produto válido" & vbCrLf
      Pintar_Controle mskSaldoEst, tpCorContr_Erro
    End If
    If Not Valida_Moeda(mskCustoProduto, TpNaoObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar o custo do produto válido" & vbCrLf
      Pintar_Controle mskCustoProduto, tpCorContr_Erro
    End If
    If Not Valida_Moeda(mskMargemAjuste, TpNaoObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar a margem de ajuste do produto válida" & vbCrLf
      Pintar_Controle mskMargemAjuste, tpCorContr_Erro
    End If
    If Not Valida_Moeda(mskPrecoVenda, TpNaoObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar o preço de venda do produto válido" & vbCrLf
      Pintar_Controle mskPrecoVenda, tpCorContr_Erro
    End If
    If Not Valida_Moeda(mskTAM, TpNaoObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar o tam. do produto válido" & vbCrLf
      Pintar_Controle mskTAM, tpCorContr_Erro
    End If
    If Not Valida_Moeda(mskPAD, TpNaoObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar o pad. do produto válido" & vbCrLf
      Pintar_Controle mskPAD, tpCorContr_Erro
    End If
    If Not Valida_Moeda(mskSOB, TpNaoObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Informar o sob. do produto válido" & vbCrLf
      Pintar_Controle mskSOB, tpCorContr_Erro
    End If
  End Select
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserInsumoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserInsumoInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    '
    'Select Case intTipoInsumo
    'Case tpInsumo_Perfil: SetarFoco txtCodigo
    'Case tpInsumo_Acessorio: SetarFoco txtCodigoInsumo
    'Case tpInsumo_Produto: SetarFoco txtCodigoInsumo
    'End Select
    '
    tabDetalhes_Click (0)
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserInsumoInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim objRsProd               As ADODB.Recordset
  Dim strSql                  As String
  Dim objInsumo               As busSisMetal.clsInsumo
  Dim objItemPedido           As busSisMetal.clsItemPedido
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 6090
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  Select Case intTipoInsumo
  Case tpInsumo_Perfil
    Me.Caption = "Cadastro de Perfil"
    pictrava(1).Visible = True
    pictrava(2).Visible = False
    pictrava(3).Visible = False
    pictrava(4).Visible = False
    pictrava(5).Visible = False
    tabDetalhes.TabVisible(1) = False
    tabDetalhes.TabVisible(2) = False
    tabDetalhes.TabVisible(3) = False
    pictrava(1).Top = 960
    pictrava(2).Top = 4650
    pictrava(3).Top = 4650
    txtCodigoInsumo.Locked = True
    Label5(0).ForeColor = &H80000011  'disabilitado cinza
  Case tpInsumo_Acessorio
    Me.Caption = "Cadastro de Acessório"
    pictrava(1).Visible = False
    pictrava(2).Visible = True
    pictrava(3).Visible = False
    pictrava(4).Visible = False
    pictrava(5).Visible = False
    tabDetalhes.TabVisible(1) = False
    tabDetalhes.TabVisible(2) = False
    tabDetalhes.TabVisible(3) = False
    pictrava(1).Top = 4650
    pictrava(2).Top = 960
    pictrava(3).Top = 4650
    txtCodigoInsumo.Locked = False
    Label5(0).ForeColor = &H80000012 'preto
  Case tpInsumo_Produto
    Me.Caption = "Cadastro de Produto"
    pictrava(1).Visible = False
    pictrava(2).Visible = False
    pictrava(3).Visible = True
    pictrava(4).Visible = True
    pictrava(5).Visible = True
    tabDetalhes.TabVisible(1) = True
    tabDetalhes.TabVisible(2) = True
    tabDetalhes.TabVisible(3) = True
    pictrava(1).Top = 4650
    pictrava(2).Top = 4650
    pictrava(3).Top = 960
    txtCodigoInsumo.Locked = False
    Label5(0).ForeColor = &H80000012 'preto
  End Select
  'Limpar Campos
  LimparCampos
  'Cor
  strSql = "Select NOME from COR ORDER BY NOME"
  PreencheCombo cboCor, strSql, False, True
  'Grupo
  strSql = "Select NOME from GRUPO ORDER BY NOME"
  PreencheCombo cboGrupo, strSql, False, True
  'Embalagem
  strSql = "Select NOME from EMBALAGEM ORDER BY NOME"
  PreencheCombo cboEmbalagem, strSql, False, True
  'Embalagem Produto
  PreencheCombo cboEmbalgemProduto, strSql, False, True
  'Grupo Produto
  strSql = "Select NOME from GRUPO_PRODUTO ORDER BY NOME"
  PreencheCombo cboGrupoProduto, strSql, False, True
  'Fornecedor
  strSql = "Select LOJA.NOME from FORNECEDOR " & _
      " INNER JOIN LOJA ON LOJA.PKID = FORNECEDOR.LOJAID ORDER BY LOJA.NOME"
  PreencheCombo cboFornecedor, strSql, False, True
  'Família Produto
  strSql = "Select DESCRICAO from FAMILIAPRODUTOS ORDER BY DESCRICAO"
  PreencheCombo cboFamilia, strSql, False, True
  '
  'IPI
  strSql = "Select IPI from IPI ORDER BY IPI"
  PreencheCombo cboIPI, strSql, False, True
  'ICMS
  strSql = "Select ICMS from ICMS ORDER BY ICMS"
  PreencheCombo cboICMS, strSql, False, True
  '
  '
  If Status = tpStatus_Incluir Then
    '
    optStatus(0).Value = True
    'Visible
    optStatus(0).Visible = False
    optStatus(1).Visible = False
    Label5(5).Visible = False
    '
    Select Case intTipoInsumo
    Case tpInsumo_Produto
      tabDetalhes.TabEnabled(3) = False
    End Select
    
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objInsumo = New busSisMetal.clsInsumo
    Set objRs = objInsumo.SelecionarInsumoPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      'INSUMO
      txtCodigoInsumo.Text = objRs.Fields("CODIGO").Value & ""
      If objRs.Fields("STATUS").Value & "" = "A" Then
        optStatus(0).Value = True
        optStatus(1).Value = False
      ElseIf objRs.Fields("STATUS").Value & "" = "I" Then
        optStatus(0).Value = False
        optStatus(1).Value = True
      Else
        optStatus(0).Value = False
        optStatus(1).Value = False
      End If
      'PERFIL
      txtCodigoFim.Text = objRs.Fields("CODIGO_LINHA").Value & ""
      txtLinhaFim.Text = objRs.Fields("NOME_LINHA").Value & ""
      '
      If objRs.Fields("NOME_COR").Value & "" <> "" Then
        cboCor.Text = objRs.Fields("NOME_COR").Value & ""
      End If
      INCLUIR_VALOR_NO_MASK mskPesoMinimo, objRs.Fields("PESO_MINIMO").Value, TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskPesoEstoque, objRs.Fields("PESO_ESTOQUE").Value, TpMaskMoeda
      
      'ACESSORIO
      txtNome.Text = objRs.Fields("NOME").Value & ""
      If objRs.Fields("NOME_GRUPO").Value & "" <> "" Then
        cboGrupo.Text = objRs.Fields("NOME_GRUPO").Value & ""
      End If
      If objRs.Fields("NOME_EMBALAGEM").Value & "" <> "" Then
        cboEmbalagem.Text = objRs.Fields("NOME_EMBALAGEM").Value & ""
      End If
      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VALOR").Value, TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskQtdMinima, objRs.Fields("QTD_MINIMA").Value, TpMaskLongo
      INCLUIR_VALOR_NO_MASK mskQtdEstoque, objRs.Fields("QTD_ESTOQUE").Value, TpMaskLongo
      '
      If intTipoInsumo = tpInsumo_Perfil Then
        Set objItemPedido = New busSisMetal.clsItemPedido
        
        INCLUIR_VALOR_NO_MASK mskQtdMinPerfil, _
          objItemPedido.CalculoQuantidadePedido(objRs.Fields("LINHAID").Value, _
                                                Format(objRs.Fields("PESO_MINIMO").Value, "###,##0.0000")), _
          TpMaskLongo
        INCLUIR_VALOR_NO_MASK mskQtdEstPerfil, _
          objItemPedido.CalculoQuantidadePedido(objRs.Fields("LINHAID").Value, _
                                                Format(objRs.Fields("PESO_ESTOQUE").Value, "###,##0.0000")), _
          TpMaskLongo
        '
        Set objItemPedido = Nothing
      ElseIf intTipoInsumo = tpInsumo_Produto Then
        Set objRsProd = objInsumo.SelecionarProdutoPeloPkid(lngPKID)
        If Not objRsProd.EOF Then
          txtNomeProduto.Text = objRsProd.Fields("NOME").Value & ""
          txtNomeAbrevProduto.Text = objRsProd.Fields("NOMEABREVIADO").Value & ""
          If objRsProd.Fields("NOME_EMBALAGEM").Value & "" <> "" Then
            cboEmbalgemProduto.Text = objRsProd.Fields("NOME_EMBALAGEM").Value & ""
          End If
          If objRsProd.Fields("NOME_GRUPO").Value & "" <> "" Then
            cboGrupoProduto.Text = objRsProd.Fields("NOME_GRUPO").Value & ""
          End If
          If objRsProd.Fields("NOME_FORNECEDOR").Value & "" <> "" Then
            cboFornecedor.Text = objRsProd.Fields("NOME_FORNECEDOR").Value & ""
          End If
          INCLUIR_VALOR_NO_MASK mskValorProduto, objRsProd.Fields("PRECO").Value, TpMaskMoeda
          INCLUIR_VALOR_NO_MASK mskPeso, objRsProd.Fields("PESO").Value, TpMaskMoeda
          INCLUIR_VALOR_NO_MASK mskAltura, objRsProd.Fields("ALTESQUADRIA").Value, TpMaskMoeda
          INCLUIR_VALOR_NO_MASK mskLargura, objRsProd.Fields("LARGESQUADRIA").Value, TpMaskMoeda
          txtModRef.Text = objRsProd.Fields("MODELOREFERENCIA").Value & ""
          txtTabela.Text = objRsProd.Fields("TABELA").Value & ""
          If objRsProd.Fields("NOME_FAMILIA").Value & "" <> "" Then
            cboFamilia.Text = objRsProd.Fields("NOME_FAMILIA").Value & ""
          End If
          If objRsProd.Fields("NOME_ICMS").Value & "" <> "" Then
            cboICMS.Text = objRsProd.Fields("NOME_ICMS").Value & ""
          End If
          If objRsProd.Fields("NOME_IPI").Value & "" <> "" Then
            cboIPI.Text = objRsProd.Fields("NOME_IPI").Value & ""
          End If
          INCLUIR_VALOR_NO_MASK mskFinancVenda, objRsProd.Fields("FINANCVENDA").Value, TpMaskMoeda
          INCLUIR_VALOR_NO_MASK mskEstMinimo, objRsProd.Fields("ESTMINIMO").Value, TpMaskMoeda
          INCLUIR_VALOR_NO_MASK mskMargemEst, objRsProd.Fields("MARGEMESTOQUE").Value, TpMaskMoeda
          INCLUIR_VALOR_NO_MASK mskSaldoEst, objRsProd.Fields("SALDOESTOQUE").Value, TpMaskMoeda
          INCLUIR_VALOR_NO_MASK mskCustoProduto, objRsProd.Fields("CUSTOPRODUTO").Value, TpMaskMoeda
          INCLUIR_VALOR_NO_MASK mskMargemAjuste, objRsProd.Fields("MARGEMAJUSTE").Value, TpMaskMoeda
          INCLUIR_VALOR_NO_MASK mskPrecoVenda, objRsProd.Fields("PRECOVENDA").Value, TpMaskMoeda
          INCLUIR_VALOR_NO_MASK mskTAM, objRsProd.Fields("TAM").Value, TpMaskMoeda
          INCLUIR_VALOR_NO_MASK mskPAD, objRsProd.Fields("PAD").Value, TpMaskLongo
          INCLUIR_VALOR_NO_MASK mskSOB, objRsProd.Fields("SOB").Value, TpMaskLongo
          If objRsProd.Fields("COMISSAO_VENDEDOR").Value Then
            chkComissaoVendedor.Value = 1
          Else
            chkComissaoVendedor.Value = 0
          End If
          '
        End If
        objRsProd.Close
        Set objRsProd = Nothing
        '
        tabDetalhes.TabEnabled(3) = True
      End If
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objInsumo = Nothing
    'Visible
    optStatus(0).Visible = True
    optStatus(1).Visible = True
    Label5(5).Visible = True
    '
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


Private Sub grdPrecoFilial_UnboundReadDataEx( _
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
               Offset + intI, PRFIL_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, PRFIL_COLUNASMATRIZ, PRFIL_LINHASMATRIZ, PRFIL_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, PRFIL_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmInsumoInc.grdGeral_UnboundReadDataEx]"
End Sub


Private Sub mskAltura_GotFocus()
  Seleciona_Conteudo_Controle mskAltura
End Sub
Private Sub mskAltura_LostFocus()
  Pintar_Controle mskAltura, tpCorContr_Normal
End Sub

Private Sub mskCustoProduto_GotFocus()
  Seleciona_Conteudo_Controle mskCustoProduto
End Sub
Private Sub mskCustoProduto_LostFocus()
  Pintar_Controle mskCustoProduto, tpCorContr_Normal
End Sub

Private Sub mskEstMinimo_GotFocus()
  Seleciona_Conteudo_Controle mskEstMinimo
End Sub
Private Sub mskEstMinimo_LostFocus()
  Pintar_Controle mskEstMinimo, tpCorContr_Normal
End Sub

Private Sub mskFinancVenda_GotFocus()
  Seleciona_Conteudo_Controle mskFinancVenda
End Sub
Private Sub mskFinancVenda_LostFocus()
  Pintar_Controle mskFinancVenda, tpCorContr_Normal
End Sub

Private Sub mskLargura_GotFocus()
  Seleciona_Conteudo_Controle mskLargura
End Sub
Private Sub mskLargura_LostFocus()
  Pintar_Controle mskLargura, tpCorContr_Normal
End Sub


Private Sub mskMargemAjuste_GotFocus()
  Seleciona_Conteudo_Controle mskMargemAjuste
End Sub
Private Sub mskMargemAjuste_LostFocus()
  Pintar_Controle mskMargemAjuste, tpCorContr_Normal
End Sub

Private Sub mskMargemEst_GotFocus()
  Seleciona_Conteudo_Controle mskMargemEst
End Sub
Private Sub mskMargemEst_LostFocus()
  Pintar_Controle mskMargemEst, tpCorContr_Normal
End Sub

Private Sub mskPAD_GotFocus()
  Seleciona_Conteudo_Controle mskPAD
End Sub
Private Sub mskPAD_LostFocus()
  Pintar_Controle mskPAD, tpCorContr_Normal
End Sub

Private Sub mskPeso_GotFocus()
  Seleciona_Conteudo_Controle mskPeso
End Sub
Private Sub mskPeso_LostFocus()
  Pintar_Controle mskPeso, tpCorContr_Normal
End Sub

Private Sub mskPesoEstoque_GotFocus()
  Seleciona_Conteudo_Controle mskPesoEstoque
End Sub
Private Sub mskPesoEstoque_LostFocus()
  Pintar_Controle mskPesoEstoque, tpCorContr_Normal
End Sub

Private Sub mskPesoMinimo_GotFocus()
  Seleciona_Conteudo_Controle mskPesoMinimo
End Sub
Private Sub mskPesoMinimo_LostFocus()
  Pintar_Controle mskPesoMinimo, tpCorContr_Normal
End Sub

Private Sub mskPrecoVenda_GotFocus()
  Seleciona_Conteudo_Controle mskPrecoVenda
End Sub
Private Sub mskPrecoVenda_LostFocus()
  Pintar_Controle mskPrecoVenda, tpCorContr_Normal
End Sub

Private Sub mskQtdEstPerfil_GotFocus()
  Seleciona_Conteudo_Controle mskQtdEstPerfil
End Sub
Private Sub mskQtdEstPerfil_LostFocus()
  On Error GoTo trata
  Dim objLinhaCons    As Form
  Dim objLinhaPerfil  As busSisMetal.clsLinhaPerfil
  Dim objItemPedido   As busSisMetal.clsItemPedido
  Dim objRs           As ADODB.Recordset
  If Me.ActiveControl.Name = "cmdCancelar" Then Exit Sub
  Pintar_Controle mskQtdEstPerfil, tpCorContr_Normal
  LimparCampoMask mskPesoEstoque
  If Not Valida_Moeda(mskQtdEstPerfil, TpObrigatorio, False) Then Exit Sub
  If Len(txtCodigoFim.Text) <> 0 And Len(txtLinhaFim.Text) <> 0 Then
    'Somente calcula a quantidade se ouver lançado a linha
    Set objLinhaPerfil = New busSisMetal.clsLinhaPerfil
    '
    Set objRs = objLinhaPerfil.CapturaItemLinha(txtCodigoFim.Text)
    If Not objRs.EOF Then
      If objRs.RecordCount = 1 Then
        Set objItemPedido = New busSisMetal.clsItemPedido
        
        INCLUIR_VALOR_NO_MASK mskPesoEstoque, _
          objItemPedido.CalculoPesoPedido(objRs.Fields("PKID").Value, _
                                          Format(IIf(Len(mskQtdEstPerfil.ClipText) = 0, "", mskQtdEstPerfil.Text), "###,##0")), _
          TpMaskMoeda
        Set objItemPedido = Nothing
      End If
      '
      objRs.Close
      Set objRs = Nothing
      Set objLinhaPerfil = Nothing
    End If
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub mskQtdMinima_GotFocus()
  Seleciona_Conteudo_Controle mskQtdMinima
End Sub
Private Sub mskQtdMinima_LostFocus()
  Pintar_Controle mskQtdMinima, tpCorContr_Normal
End Sub


Private Sub mskQtdMinPerfil_GotFocus()
  Seleciona_Conteudo_Controle mskQtdMinPerfil
End Sub
Private Sub mskQtdMinPerfil_LostFocus()
  On Error GoTo trata
  Dim objLinhaCons    As Form
  Dim objLinhaPerfil  As busSisMetal.clsLinhaPerfil
  Dim objItemPedido   As busSisMetal.clsItemPedido
  Dim objRs           As ADODB.Recordset
  If Me.ActiveControl.Name = "cmdCancelar" Then Exit Sub
  Pintar_Controle mskQtdMinPerfil, tpCorContr_Normal
  LimparCampoMask mskPesoMinimo
  If Not Valida_Moeda(mskQtdMinPerfil, TpObrigatorio, False) Then Exit Sub
  If Len(txtCodigoFim.Text) <> 0 And Len(txtLinhaFim.Text) <> 0 Then
    'Somente calcula a quantidade se ouver lançado a linha
    Set objLinhaPerfil = New busSisMetal.clsLinhaPerfil
    '
    Set objRs = objLinhaPerfil.CapturaItemLinha(txtCodigoFim.Text)
    If Not objRs.EOF Then
      If objRs.RecordCount = 1 Then
        Set objItemPedido = New busSisMetal.clsItemPedido
        
        INCLUIR_VALOR_NO_MASK mskPesoMinimo, _
          objItemPedido.CalculoPesoPedido(objRs.Fields("PKID").Value, _
                                          Format(IIf(Len(mskQtdMinPerfil.ClipText) = 0, "", mskQtdMinPerfil.Text), "###,##0")), _
          TpMaskMoeda
        Set objItemPedido = Nothing
      End If
      '
      objRs.Close
      Set objRs = Nothing
      Set objLinhaPerfil = Nothing
    End If
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub mskSaldoEst_GotFocus()
  Seleciona_Conteudo_Controle mskSaldoEst
End Sub
Private Sub mskSaldoEst_LostFocus()
  Pintar_Controle mskSaldoEst, tpCorContr_Normal
End Sub

Private Sub mskSOB_GotFocus()
  Seleciona_Conteudo_Controle mskSOB
End Sub
Private Sub mskSOB_LostFocus()
  Pintar_Controle mskSOB, tpCorContr_Normal
End Sub

Private Sub mskTAM_GotFocus()
  Seleciona_Conteudo_Controle mskTAM
End Sub
Private Sub mskTAM_LostFocus()
  Pintar_Controle mskTAM, tpCorContr_Normal
End Sub

Private Sub mskValor_GotFocus()
  Seleciona_Conteudo_Controle mskValor
End Sub
Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub

Private Sub mskQtdEstoque_GotFocus()
  Seleciona_Conteudo_Controle mskQtdEstoque
End Sub
Private Sub mskQtdEstoque_LostFocus()
  Pintar_Controle mskQtdEstoque, tpCorContr_Normal
End Sub


Private Sub mskValorProduto_GotFocus()
  Seleciona_Conteudo_Controle mskValorProduto
End Sub
Private Sub mskValorProduto_LostFocus()
  Pintar_Controle mskValorProduto, tpCorContr_Normal
End Sub

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    'Dados cadastrais
    Select Case intTipoInsumo
    Case tpInsumo_Perfil
      pictrava(0).Enabled = True
      pictrava(1).Enabled = True
      pictrava(2).Enabled = False
      pictrava(3).Enabled = False
      pictrava(4).Enabled = False
      pictrava(5).Enabled = False
      'tabDetalhes.TabVisible(1) = False
      'tabDetalhes.TabVisible(2) = False
    Case tpInsumo_Acessorio
      pictrava(0).Enabled = True
      pictrava(1).Enabled = False
      pictrava(2).Enabled = True
      pictrava(3).Enabled = False
      pictrava(4).Enabled = False
      pictrava(5).Enabled = False
      'tabDetalhes.TabVisible(1) = False
      'tabDetalhes.TabVisible(2) = False
    Case tpInsumo_Produto
      grdPrecoFilial.Enabled = False
      pictrava(0).Enabled = True
      pictrava(1).Enabled = False
      pictrava(2).Enabled = False
      pictrava(3).Enabled = True
      pictrava(4).Enabled = False
      pictrava(5).Enabled = False
      'tabDetalhes.TabVisible(1) = True
      'tabDetalhes.TabVisible(2) = True
    End Select
    'grdEspecialidade.Enabled = False
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    'cmdExcluir.Enabled = False
    'cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    Select Case intTipoInsumo
    Case tpInsumo_Perfil: SetarFoco txtCodigo
    Case tpInsumo_Acessorio: SetarFoco txtCodigoInsumo
    Case tpInsumo_Produto: SetarFoco txtCodigoInsumo
    End Select
  Case 1
    'Complemento
    Select Case intTipoInsumo
    Case tpInsumo_Perfil
      pictrava(0).Enabled = False
      pictrava(1).Enabled = False
      pictrava(2).Enabled = False
      pictrava(3).Enabled = False
      pictrava(4).Enabled = False
      pictrava(5).Enabled = False
      'tabDetalhes.TabVisible(1) = False
      'tabDetalhes.TabVisible(2) = False
    Case tpInsumo_Acessorio
      pictrava(0).Enabled = False
      pictrava(1).Enabled = False
      pictrava(2).Enabled = False
      pictrava(3).Enabled = False
      pictrava(4).Enabled = False
      pictrava(5).Enabled = False
      'tabDetalhes.TabVisible(1) = False
      'tabDetalhes.TabVisible(2) = False
    Case tpInsumo_Produto
      grdPrecoFilial.Enabled = False
      pictrava(0).Enabled = False
      pictrava(1).Enabled = False
      pictrava(2).Enabled = False
      pictrava(3).Enabled = False
      pictrava(4).Enabled = True
      pictrava(5).Enabled = False
      'tabDetalhes.TabVisible(1) = True
      'tabDetalhes.TabVisible(2) = True
    End Select
    'grdEspecialidade.Enabled = False
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    'cmdExcluir.Enabled = False
    'cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    SetarFoco cboICMS
  Case 2
    'Imagem
    Select Case intTipoInsumo
    Case tpInsumo_Perfil
      pictrava(0).Enabled = False
      pictrava(1).Enabled = False
      pictrava(2).Enabled = False
      pictrava(3).Enabled = False
      pictrava(4).Enabled = False
      pictrava(5).Enabled = False
      'tabDetalhes.TabVisible(1) = False
      'tabDetalhes.TabVisible(2) = False
    Case tpInsumo_Acessorio
      pictrava(0).Enabled = False
      pictrava(1).Enabled = False
      pictrava(2).Enabled = False
      pictrava(3).Enabled = False
      pictrava(4).Enabled = False
      pictrava(5).Enabled = False
      'tabDetalhes.TabVisible(1) = False
      'tabDetalhes.TabVisible(2) = False
    Case tpInsumo_Produto
      grdPrecoFilial.Enabled = False
      pictrava(0).Enabled = False
      pictrava(1).Enabled = False
      pictrava(2).Enabled = False
      pictrava(3).Enabled = False
      pictrava(4).Enabled = False
      pictrava(5).Enabled = True
      'tabDetalhes.TabVisible(1) = True
      'tabDetalhes.TabVisible(2) = True
    End Select
    'grdEspecialidade.Enabled = False
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    'cmdExcluir.Enabled = False
    'cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    'TRATAR IMAGEM DO SCANNER
    Image1.Picture = LoadPicture(gsCaminhoImagemCompra & txtCodigoInsumo & ".jpg")
    Image1.Refresh
    'SetarFoco cboICMS
  Case 3
    'Preçofilial
    grdPrecoFilial.Enabled = True
    pictrava(0).Enabled = False
    pictrava(1).Enabled = False
    pictrava(2).Enabled = False
    pictrava(3).Enabled = False
    pictrava(4).Enabled = False
    pictrava(5).Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
'    cmdExcluir.Enabled = True
'    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = True
    '
    'Montar RecordSet
    PRFIL_COLUNASMATRIZ = grdPrecoFilial.Columns.Count
    PRFIL_LINHASMATRIZ = 0
    PRFIL_MontaMatriz
    grdPrecoFilial.Bookmark = Null
    grdPrecoFilial.ReBind
    grdPrecoFilial.ApproxCount = PRFIL_LINHASMATRIZ
    '
    SetarFoco grdPrecoFilial
  End Select
  Exit Sub
trata:
  'Tratamento de erro de leitura de imagem (File not found)
  If Err.Number = 53 Then
    Image1.Picture = LoadPicture(gsIconsPath & "Excluir.ico")
    Image1.Refresh
    Resume Next
  End If
  TratarErro Err.Number, Err.Description, "frmInsumoInc.tabDetalhes"
  AmpN
End Sub

Public Sub PRFIL_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGer    As busSisMetal.clsGeral
  '
  On Error GoTo trata
  
  Set objGer = New busSisMetal.clsGeral
  '
  strSql = "SELECT TAB_PRECOFILIAL.PKID, LOJA.NOME, TAB_PRECOFILIAL.VALOR " & _
          "FROM TAB_PRECOFILIAL INNER JOIN LOJA ON LOJA.PKID = TAB_PRECOFILIAL.FILIALID " & _
          "WHERE TAB_PRECOFILIAL.PRODUTOID = " & Formata_Dados(lngPKID, tpDados_Longo) & _
          " ORDER BY LOJA.NOME"

  '
  Set objRs = objGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    PRFIL_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim PRFIL_Matriz(0 To PRFIL_COLUNASMATRIZ - 1, 0 To PRFIL_LINHASMATRIZ - 1)
  Else
    ReDim PRFIL_Matriz(0 To PRFIL_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To PRFIL_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To PRFIL_COLUNASMATRIZ - 1  'varre as colunas
          PRFIL_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set objGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtCodigo_GotFocus()
  Seleciona_Conteudo_Controle txtCodigo
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  KeyAscii = TRANSFORMA_MAIUSCULA(KeyAscii)
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtCodigo_LostFocus()
  On Error GoTo trata
  Dim objLinhaCons    As Form
  Dim objLinhaPerfil  As busSisMetal.clsLinhaPerfil
  Dim objRs           As ADODB.Recordset
  If Me.ActiveControl.Name = "cmdCancelar" Then Exit Sub

  Pintar_Controle txtCodigo, tpCorContr_Normal
  If Len(txtCodigo.Text) = 0 Then
    If Len(txtCodigoFim.Text) <> 0 And Len(txtLinhaFim.Text) <> 0 Then
      Exit Sub
    Else
      TratarErroPrevisto "Entre com o código ou descrição da linha."
      Pintar_Controle txtCodigo, tpCorContr_Erro
      SetarFoco txtCodigo
      Exit Sub
    End If
  End If
  Set objLinhaPerfil = New busSisMetal.clsLinhaPerfil
  '
  Set objRs = objLinhaPerfil.CapturaItemLinha(txtCodigo.Text)
  If objRs.EOF Then
    LimparCampoTexto txtCodigoFim
    LimparCampoTexto txtLinhaFim
    TratarErroPrevisto "Descrição/Código da linha não cadastrado"
    Pintar_Controle txtCodigo, tpCorContr_Erro
    SetarFoco txtCodigo
    Exit Sub
  Else
    If objRs.RecordCount = 1 Then
      txtCodigoFim.Text = objRs.Fields("CODIGO").Value & ""
      txtLinhaFim.Text = objRs.Fields("NOME").Value & ""
    Else
      'Novo : apresentar tela para seleção da linha
      Set objLinhaCons = New frmLinhaCons
      objLinhaCons.strCodigoDescricao = txtCodigo.Text
      objLinhaCons.intIcOrigemLn = 1
      objLinhaCons.Show vbModal
    End If
    '
    objRs.Close
    Set objRs = Nothing
    Set objLinhaPerfil = Nothing
'''    cmdOk.Default = True
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtCodigoInsumo_LostFocus()
  Pintar_Controle txtCodigoInsumo, tpCorContr_Normal
End Sub

Private Sub txtModRef_GotFocus()
  Seleciona_Conteudo_Controle txtModRef
End Sub
Private Sub txtModRef_LostFocus()
  Pintar_Controle txtModRef, tpCorContr_Normal
End Sub

Private Sub txtNome_GotFocus()
  Seleciona_Conteudo_Controle txtNome
End Sub
Private Sub txtNome_LostFocus()
  Pintar_Controle txtNome, tpCorContr_Normal
End Sub

Private Sub txtNomeAbrevProduto_GotFocus()
  Seleciona_Conteudo_Controle txtNomeAbrevProduto
End Sub
Private Sub txtNomeAbrevProduto_LostFocus()
  Pintar_Controle txtNomeAbrevProduto, tpCorContr_Normal
End Sub

Private Sub txtNomeProduto_GotFocus()
  Seleciona_Conteudo_Controle txtNomeProduto
End Sub
Private Sub txtNomeProduto_LostFocus()
  Pintar_Controle txtNomeProduto, tpCorContr_Normal
End Sub

Private Sub txtTabela_GotFocus()
  Seleciona_Conteudo_Controle txtTabela
End Sub
Private Sub txtTabela_LostFocus()
  Pintar_Controle txtTabela, tpCorContr_Normal
End Sub

