VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserPessoaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de"
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
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4605
         Left            =   90
         ScaleHeight     =   4545
         ScaleWidth      =   1605
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1665
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&X"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   3570
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   2730
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5985
      Left            =   120
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   10557
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Cadastro"
      TabPicture(0)   =   "userPessoaInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&End. res."
      TabPicture(1)   =   "userPessoaInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Especialidade"
      TabPicture(2)   =   "userPessoaInc.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdEspecialidade"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Procedimento"
      TabPicture(3)   =   "userPessoaInc.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grdProcedimento"
      Tab(3).ControlCount=   1
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
         Left            =   -74910
         TabIndex        =   52
         Top             =   360
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   4515
            Index           =   3
            Left            =   120
            ScaleHeight     =   4515
            ScaleWidth      =   7575
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtEstadoRes 
               Height          =   285
               Left            =   1320
               MaxLength       =   2
               TabIndex        =   28
               Text            =   "txtEstadoRes"
               Top             =   1080
               Width           =   435
            End
            Begin VB.TextBox txtComplementoRes 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   27
               Text            =   "txtComplementoRes"
               Top             =   750
               Width           =   6075
            End
            Begin VB.TextBox txtNumeroRes 
               Height          =   285
               Left            =   1320
               MaxLength       =   15
               TabIndex        =   26
               Text            =   "txtNumeroRes"
               Top             =   420
               Width           =   2175
            End
            Begin VB.TextBox txtRuaRes 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   25
               Text            =   "txtRuaRes"
               Top             =   90
               Width           =   6075
            End
            Begin VB.TextBox txtBairroRes 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   30
               Text            =   "txtBairroRes"
               Top             =   1410
               Width           =   6075
            End
            Begin VB.TextBox txtCidadeRes 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   31
               Text            =   "txtCidadeRes"
               Top             =   1740
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskCepRes 
               Height          =   285
               Left            =   5220
               TabIndex        =   29
               Top             =   1080
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##.###-###"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Estado"
               Height          =   285
               Index           =   9
               Left            =   60
               TabIndex        =   60
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Complemento"
               Height          =   195
               Index           =   8
               Left            =   60
               TabIndex        =   59
               Top             =   750
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Número"
               Height          =   195
               Index           =   7
               Left            =   60
               TabIndex        =   58
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Rua"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   57
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Bairro"
               Height          =   285
               Index           =   2
               Left            =   60
               TabIndex        =   56
               Top             =   1455
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cidade"
               Height          =   285
               Index           =   16
               Left            =   60
               TabIndex        =   55
               Top             =   1785
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cep"
               Height          =   285
               Index           =   3
               Left            =   3960
               TabIndex        =   54
               Top             =   1080
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
         TabIndex        =   42
         Top             =   360
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   1545
            Index           =   2
            Left            =   120
            ScaleHeight     =   1545
            ScaleWidth      =   7575
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   3960
            Width           =   7575
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   4800
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   1110
               Width           =   2235
               Begin VB.OptionButton optFuncExcluido 
                  Caption         =   "Sim"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   23
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1065
               End
               Begin VB.OptionButton optFuncExcluido 
                  Caption         =   "Não"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   24
                  Top             =   0
                  Width           =   1095
               End
            End
            Begin VB.TextBox txtNovaSenha 
               Height          =   288
               IMEMode         =   3  'DISABLE
               Left            =   1320
               MaxLength       =   6
               PasswordChar    =   "#"
               TabIndex        =   21
               Text            =   "txtNov"
               Top             =   750
               Width           =   1095
            End
            Begin VB.TextBox txtConfSenha 
               Height          =   288
               IMEMode         =   3  'DISABLE
               Left            =   1320
               MaxLength       =   6
               PasswordChar    =   "#"
               TabIndex        =   22
               Text            =   "txtCon"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtUsuario 
               Height          =   285
               Left            =   1320
               MaxLength       =   30
               TabIndex        =   20
               Top             =   420
               Width           =   2745
            End
            Begin VB.ComboBox cboNivel 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   60
               Width           =   2775
            End
            Begin VB.Label Label5 
               Caption         =   "Excluido"
               Height          =   285
               Index           =   14
               Left            =   3870
               TabIndex        =   77
               Top             =   1140
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "Senha"
               Height          =   255
               Left            =   90
               TabIndex        =   73
               Top             =   750
               Width           =   975
            End
            Begin VB.Label Label3 
               Caption         =   "Confirmar Senha"
               Height          =   255
               Left            =   90
               TabIndex        =   72
               Top             =   1080
               Width           =   1335
            End
            Begin VB.Label Label6 
               Caption         =   "Nível"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   71
               Top             =   60
               Width           =   1455
            End
            Begin VB.Label Label6 
               Caption         =   "Usuário"
               Height          =   255
               Index           =   0
               Left            =   90
               TabIndex        =   70
               Top             =   420
               Width           =   1455
            End
         End
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   1005
            Index           =   1
            Left            =   120
            ScaleHeight     =   1005
            ScaleWidth      =   7575
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   2940
            Width           =   7575
            Begin VB.TextBox txtConta 
               Height          =   285
               Left            =   5070
               MaxLength       =   20
               TabIndex        =   16
               Text            =   "txtConta"
               Top             =   390
               Width           =   2325
            End
            Begin VB.ComboBox cboBanco 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   60
               Width           =   6105
            End
            Begin VB.PictureBox Picture4 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1290
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   690
               Width           =   2235
               Begin VB.OptionButton optMotExcluido 
                  Caption         =   "Não"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   18
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton optMotExcluido 
                  Caption         =   "Sim"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   17
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1065
               End
            End
            Begin VB.TextBox txtAgencia 
               Height          =   285
               Left            =   1320
               MaxLength       =   5
               TabIndex        =   15
               Text            =   "txtAgencia"
               Top             =   390
               Width           =   2325
            End
            Begin VB.Label Label5 
               Caption         =   "Conta"
               Height          =   195
               Index           =   12
               Left            =   3840
               TabIndex        =   79
               Top             =   405
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Banco"
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   78
               Top             =   90
               Width           =   1455
            End
            Begin VB.Label Label5 
               Caption         =   "Excluido"
               Height          =   285
               Index           =   13
               Left            =   90
               TabIndex        =   75
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Agência"
               Height          =   195
               Index           =   11
               Left            =   90
               TabIndex        =   69
               Top             =   405
               Width           =   1215
            End
         End
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2775
            Index           =   0
            Left            =   120
            ScaleHeight     =   2775
            ScaleWidth      =   7575
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1320
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   360
               Width           =   2235
               Begin VB.OptionButton optTipoPessoa 
                  Caption         =   "&Física"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   1
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1065
               End
               Begin VB.OptionButton optTipoPessoa 
                  Caption         =   "&Jurídica"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   2
                  Top             =   0
                  Width           =   1095
               End
            End
            Begin VB.TextBox txtOrgaoEmissor 
               Height          =   285
               Left            =   1320
               MaxLength       =   30
               TabIndex        =   9
               Text            =   "txtOrgaoEmissor"
               Top             =   1530
               Width           =   2325
            End
            Begin VB.TextBox txtIdentidade 
               Height          =   285
               Left            =   4830
               MaxLength       =   30
               TabIndex        =   8
               Text            =   "txtIdentidade"
               Top             =   1200
               Width           =   2565
            End
            Begin VB.TextBox txtObservacao 
               Height          =   615
               Left            =   1320
               MaxLength       =   255
               MultiLine       =   -1  'True
               TabIndex        =   13
               Text            =   "userPessoaInc.frx":0070
               Top             =   2130
               Width           =   6075
            End
            Begin VB.TextBox txtCelular 
               Height          =   285
               Left            =   5070
               MaxLength       =   20
               TabIndex        =   12
               Text            =   "txtCelular"
               Top             =   1830
               Width           =   2175
            End
            Begin VB.TextBox txtTelefoneRes 
               Height          =   285
               Left            =   1320
               MaxLength       =   20
               TabIndex        =   11
               Text            =   "txtTelefoneRes"
               Top             =   1830
               Width           =   2175
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1320
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   930
               Width           =   2235
               Begin VB.OptionButton optSexo 
                  Caption         =   "&Feminino"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   6
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton optSexo 
                  Caption         =   "&Masculino"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   5
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
               Top             =   60
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskDtNascimento 
               Height          =   255
               Left            =   1320
               TabIndex        =   7
               Top             =   1260
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
               TabIndex        =   3
               Top             =   660
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   14
               Mask            =   "###.###.###-##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDtExpedicao 
               Height          =   255
               Left            =   5070
               TabIndex        =   10
               Top             =   1530
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskCnpj 
               Height          =   255
               Left            =   4440
               TabIndex        =   4
               Top             =   660
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   18
               Mask            =   "##.###.###/####-##"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Tipo de Pessoa"
               Height          =   315
               Index           =   10
               Left            =   90
               TabIndex        =   67
               Top             =   390
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cnpj"
               Height          =   195
               Index           =   6
               Left            =   3180
               TabIndex        =   65
               Top             =   660
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Dt. Expedição"
               Height          =   255
               Index           =   0
               Left            =   3810
               TabIndex        =   64
               Top             =   1530
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Órg. Emissor"
               Height          =   195
               Index           =   42
               Left            =   90
               TabIndex        =   63
               Top             =   1545
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Identidade"
               Height          =   195
               Index           =   39
               Left            =   3840
               TabIndex        =   62
               Top             =   1245
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Observação"
               Height          =   195
               Index           =   32
               Left            =   90
               TabIndex        =   51
               Top             =   2175
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Celular"
               Height          =   195
               Index           =   28
               Left            =   3810
               TabIndex        =   50
               Top             =   1830
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Telefone res."
               Height          =   195
               Index           =   27
               Left            =   90
               TabIndex        =   49
               Top             =   1830
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "CPF"
               Height          =   195
               Index           =   4
               Left            =   90
               TabIndex        =   48
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Dt. Nascimento"
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   47
               Top             =   1260
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Sexo"
               Height          =   315
               Index           =   5
               Left            =   90
               TabIndex        =   45
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Nome"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   44
               Top             =   120
               Width           =   1215
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdEspecialidade 
         Height          =   4725
         Left            =   -74910
         OleObjectBlob   =   "userPessoaInc.frx":007E
         TabIndex        =   32
         Top             =   390
         Width           =   7905
      End
      Begin TrueDBGrid60.TDBGrid grdProcedimento 
         Height          =   4725
         Left            =   -74910
         OleObjectBlob   =   "userPessoaInc.frx":45E0
         TabIndex        =   33
         Top             =   390
         Width           =   7905
      End
   End
End
Attribute VB_Name = "frmUserPessoaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strNomeInicial           As String

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean
Public blnPrimeiraVez           As Boolean
Public IcPessoa                 As tpIcPessoa


Public lngPKID                  As Long
Public lngPESSOAID              As Long

Dim ESPEC_COLUNASMATRIZ         As Long
Dim ESPEC_LINHASMATRIZ          As Long
Private ESPEC_Matriz()          As String

Dim PROCED_COLUNASMATRIZ        As Long
Dim PROCED_LINHASMATRIZ         As Long
Private PROCED_Matriz()         As String

Public intQuemChamou            As Integer
'Assume
'0  Chamada do cadastro
'1  Chamada da GR

Private Sub TratarCampos()
  On Error GoTo trata
  Dim intTopAux As Integer
  intTopAux = 2940
  If IcPessoa = tpIcPessoa.tpIcPessoa_Func Then
    'Funcionário
    picTrava(2).Top = intTopAux
    '
    picTrava(1).Visible = False
    picTrava(2).Visible = True
    tabDetalhes.TabVisible(2) = False
    tabDetalhes.TabVisible(3) = False
    '
    Me.Caption = "Cadastro de Funcionário"
    If Status = tpStatus_Incluir Then
      'Trtar exclusão
      optFuncExcluido(1).Value = True
      'Visible
      optFuncExcluido(0).Visible = False
      optFuncExcluido(1).Visible = False
      Label5(14).Visible = False
    ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
      'Visible
      optFuncExcluido(0).Visible = True
      optFuncExcluido(1).Visible = True
      Label5(14).Visible = True
    End If
    'Tratar tipo de pessoa
    Label5(10).Enabled = False
    optTipoPessoa(0).Enabled = False
    optTipoPessoa(1).Enabled = False
  ElseIf IcPessoa = tpIcPessoa.tpIcPessoa_Mot Then
    'Motorista
    picTrava(1).Top = intTopAux
    '
    picTrava(1).Visible = True
    picTrava(2).Visible = False
    tabDetalhes.TabVisible(2) = True
    tabDetalhes.TabVisible(3) = True
    '
    Me.Caption = "Cadastro de Motorista"
    '
    If Status = tpStatus_Incluir Then
      'Trtar exclusão
      optMotExcluido(1).Value = True
      'Visible
      optMotExcluido(0).Visible = False
      optMotExcluido(1).Visible = False
      Label5(13).Visible = False
    ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
      'Visible
      optMotExcluido(0).Visible = True
      optMotExcluido(1).Visible = True
      Label5(13).Visible = True
    End If
    'Tratar tipo de pessoa
    Label5(10).Enabled = True
    optTipoPessoa(0).Enabled = True
    optTipoPessoa(1).Enabled = True
  End If
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserPessoaInc.TratarCampos]", _
            Err.Description
End Sub

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Pessoa
  'Dados cadastrais
  LimparCampoTexto txtNome
  optTipoPessoa(0).Value = False
  optTipoPessoa(1).Value = False
  LimparCampoMask mskCpf
  LimparCampoMask mskCnpj
  optSexo(0).Value = False
  optSexo(1).Value = False
  LimparCampoMask mskDtNascimento
  LimparCampoTexto txtIdentidade
  LimparCampoTexto txtOrgaoEmissor
  LimparCampoMask mskDtExpedicao
  LimparCampoTexto txtTelefoneRes
  LimparCampoTexto txtCelular
  LimparCampoTexto txtObservacao
  'FUNCIONARIO
  LimparCampoTexto txtUsuario
  LimparCampoCombo cboNivel
  LimparCampoTexto txtNovaSenha
  LimparCampoTexto txtConfSenha
  optFuncExcluido(0).Value = False
  optFuncExcluido(1).Value = False
  'MOTORISTA
  LimparCampoCombo cboBanco
  LimparCampoTexto txtAgencia
  LimparCampoTexto txtConta
  optMotExcluido(0).Value = False
  optMotExcluido(1).Value = False
  'Endereço res
  LimparCampoTexto txtRuaRes
  LimparCampoTexto txtNumeroRes
  LimparCampoTexto txtComplementoRes
  LimparCampoTexto txtEstadoRes
  LimparCampoMask mskCepRes
  LimparCampoTexto txtBairroRes
  LimparCampoTexto txtCidadeRes
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserPessoaInc.LimparCampos]", _
            Err.Description
End Sub



Private Sub cboBanco_LostFocus()
  Pintar_Controle cboBanco, tpCorContr_Normal
End Sub


Private Sub cboNivel_Click()
  On Error GoTo trata
  Select Case Left(cboNivel.Text, 3)
  Case "SEM"
    Label6(0).Enabled = False
    txtUsuario.Enabled = False
    Label1.Enabled = False
    txtNovaSenha.Enabled = False
    Label3.Enabled = False
    txtConfSenha.Enabled = False
    '
    txtUsuario.Text = ""
    txtNovaSenha.Text = ""
    txtConfSenha.Text = ""
  Case Else
    Label6(0).Enabled = True
    txtUsuario.Enabled = True
    Label1.Enabled = True
    txtNovaSenha.Enabled = True
    Label3.Enabled = True
    txtConfSenha.Enabled = True
    '
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cboNivel_LostFocus()
  Pintar_Controle cboNivel, tpCorContr_Normal
End Sub

Private Sub cmdAlterar_Click()
  On Error GoTo trata
  'Dim objFormProcPrestador As Elite.frmUserProcPrestadorInc
  Select Case tabDetalhes.Tab
  Case 3
'''    'Proc Prestador
'''    If Not IsNumeric(grdProcedimento.Columns("PKID").Value & "") Then
'''      MsgBox "Selecione um procedimento !", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdProcedimento
'''      Exit Sub
'''    End If
'''
'''    Set objFormProcPrestador = New Elite.frmUserProcPrestadorInc
'''    objFormProcPrestador.Status = tpStatus_Alterar
'''    objFormProcPrestador.lngPKID = grdProcedimento.Columns("PKID").Value
'''    objFormProcPrestador.lngPRESTADORID = lngPKID
'''    objFormProcPrestador.strNomePrestador = txtNome.Text
'''    objFormProcPrestador.Show vbModal
'''    If objFormProcPrestador.blnRetorno Then
'''      PROCED_MontaMatriz
'''      grdProcedimento.Bookmark = Null
'''      grdProcedimento.ReBind
'''      grdProcedimento.ApproxCount = PROCED_LINHASMATRIZ
'''    End If
'''    Set objFormProcPrestador = Nothing
'''    SetarFoco grdProcedimento
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
'''  Dim objPrestProcedimento      As busElite.clsPrestProcedimento
'''  '
'''  Select Case tabDetalhes.Tab
'''  Case 3 'Exclusão de associado
'''    If Not IsNumeric(grdProcedimento.Columns("PKID").Value & "") Then
'''      MsgBox "Selecione um procedimento para exclusão.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdProcedimento
'''      Exit Sub
'''    End If
'''    '
'''    If MsgBox("Confirma exclusão do Procedimento " & grdProcedimento.Columns("Procedimento").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco grdProcedimento
'''      Exit Sub
'''    End If
'''    Set objPrestProcedimento = New busElite.clsPrestProcedimento
'''    objPrestProcedimento.ExcluirPrestProcedimento CLng(grdProcedimento.Columns("PKID").Value)
'''    Set objPrestProcedimento = Nothing
'''    '
'''    PROCED_MontaMatriz
'''    grdProcedimento.Bookmark = Null
'''    grdProcedimento.ReBind
'''    SetarFoco grdProcedimento
'''
'''  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdIncluir_Click()
  On Error GoTo trata
'''  Dim objFormProcPrestador As Elite.frmUserProcPrestadorInc
'''  '
'''  Select Case tabDetalhes.Tab
'''  Case 2
'''    'Linha
'''    frmUserPrestadorEspInc.lngPRESTADORID = lngPKID
'''    frmUserPrestadorEspInc.strPrestador = txtNome.Text
'''    frmUserPrestadorEspInc.Show vbModal
'''
'''    If frmUserPrestadorEspInc.bRetorno Then
'''      ESPEC_MontaMatriz
'''      grdEspecialidade.Bookmark = Null
'''      grdEspecialidade.ReBind
'''      grdEspecialidade.ApproxCount = ESPEC_LINHASMATRIZ
'''    End If
'''    SetarFoco grdEspecialidade
'''  Case 3
'''    'PRESTADOR PROCEDIMENTO
'''    Set objFormProcPrestador = New Elite.frmUserProcPrestadorInc
'''    objFormProcPrestador.Status = tpStatus_Incluir
'''    objFormProcPrestador.lngPKID = 0
'''    objFormProcPrestador.lngPRESTADORID = lngPKID
'''    objFormProcPrestador.strNomePrestador = txtNome.Text
'''    objFormProcPrestador.Show vbModal
'''    If objFormProcPrestador.blnRetorno Then
'''      PROCED_MontaMatriz
'''      grdProcedimento.Bookmark = Null
'''      grdProcedimento.ReBind
'''      grdProcedimento.ApproxCount = PROCED_LINHASMATRIZ
'''    End If
'''    Set objFormProcPrestador = Nothing
'''    SetarFoco grdProcedimento
'''  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Resume Next
End Sub

Private Sub cmdOk_Click()
  Dim objPessoa                  As busElite.clsPessoa
  Dim objFuncionario             As busElite.clsFuncionario
'''  Dim objPaciente                As busElite.clsPaciente
  Dim objMotorista              As busElite.clsMotorista
  Dim objGeral                  As busElite.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim strTipoPessoa             As String
  Dim strSexo                   As String
  Dim strAceitaCheque           As String
  Dim strMotExcluido            As String
  Dim strFuncExcluido           As String
  Dim lngBANCOID                As Long
  Dim strDtDesativacao          As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busElite.clsGeral
  Set objPessoa = New busElite.clsPessoa
  If IcPessoa = tpIcPessoa_Func Then
'''    'QUALIFICACAO
'''    lngQualificacao = 0
'''    strSql = "SELECT QUALIFICACAO.PKID FROM QUALIFICACAO WHERE QUALIFICACAO.DESCRICAO = " & Formata_Dados(cboQualificacao.Text, tpDados_Texto)
'''    Set objRs = objGeral.ExecutarSQL(strSql)
'''    If Not objRs.EOF Then
'''      lngQualificacao = objRs.Fields("PKID").Value
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
  ElseIf IcPessoa = tpIcPessoa_Mot Then
'''  ElseIf IcPessoa = tpIcPessoa_Prest Then
    'BANCO
    lngBANCOID = 0
    strSql = "SELECT BANCO.PKID FROM BANCO WHERE BANCO.NOME = " & Formata_Dados(cboBanco.Text, tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngBANCOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  'tipo de pessoa
  If optTipoPessoa(0).Value Then
    strTipoPessoa = "F"
  ElseIf optTipoPessoa(1).Value Then
    strTipoPessoa = "J"
  Else
    strTipoPessoa = ""
  End If
  'Sexo
  If optSexo(0).Value Then
    strSexo = "M"
  ElseIf optSexo(1).Value Then
    strSexo = "F"
  Else
    strSexo = ""
  End If
  If IcPessoa = tpIcPessoa_Func Then
    'Exlcuido Func
    If optFuncExcluido(0).Value Then
      strFuncExcluido = "S"
    ElseIf optFuncExcluido(1).Value Then
      strFuncExcluido = "N"
    Else
      strFuncExcluido = ""
    End If
  
  ElseIf IcPessoa = tpIcPessoa_Mot Then
    'Exlcuido Motorista
    If optMotExcluido(0).Value Then
      strMotExcluido = "S"
    ElseIf optMotExcluido(1).Value Then
      strMotExcluido = "N"
    Else
      strMotExcluido = ""
    End If
  End If
  
  'Validar se prestador já cadastrado
  'Por nome
  strSql = "SELECT * FROM PESSOA " & _
    " WHERE PESSOA.NOME = " & Formata_Dados(txtNome.Text, tpDados_Texto) & _
    " AND PESSOA.DTNASCIMENTO = " & Formata_Dados(mskDtNascimento.Text, tpDados_DataHora) & _
    " AND PESSOA.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle txtNome, tpCorContr_Erro
    TratarErroPrevisto "prontuário já cadastrado"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objPessoa = Nothing
    cmdOk.Enabled = True
    SetarFoco txtNome
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  'Por CPF/CNPJ
  strSql = "SELECT * FROM PESSOA " & _
    " WHERE " & IIf(optTipoPessoa(0).Value, "PESSOA.CPF = " & Formata_Dados(mskCpf.ClipText, tpDados_Texto), "PESSOA.CNPJ = " & Formata_Dados(mskCnpj.ClipText, tpDados_Texto)) & _
    " AND PESSOA.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle txtNome, tpCorContr_Erro
    TratarErroPrevisto "CPF/CNPJ já cadastrado"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objPessoa = Nothing
    cmdOk.Enabled = True
    SetarFoco mskCpf
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  If IcPessoa = tpIcPessoa_Func Then
    'Por login
    strSql = "SELECT * FROM FUNCIONARIO " & _
      " WHERE FUNCIONARIO.USUARIO = " & Formata_Dados(txtUsuario.Text, tpDados_Texto) & _
      " AND FUNCIONARIO.PESSOAID <> " & Formata_Dados(lngPKID, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      Pintar_Controle txtNome, tpCorContr_Erro
      TratarErroPrevisto "Usuário já cadastrado"
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      Set objPessoa = Nothing
      cmdOk.Enabled = True
      SetarFoco txtUsuario
      tabDetalhes.Tab = 0
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
  
  ElseIf IcPessoa = tpIcPessoa_Mot Then
  
  End If
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar Pessoa
    objPessoa.AlterarPessoa lngPKID, _
                                  IIf(mskDtNascimento.ClipText = "", "", mskDtNascimento.Text), _
                                  txtNome.Text, _
                                  txtIdentidade.Text, _
                                  txtOrgaoEmissor.Text, _
                                  IIf(mskDtExpedicao.ClipText = "", "", mskDtExpedicao.Text), _
                                  strTipoPessoa, _
                                  mskCnpj.ClipText, _
                                  mskCpf.ClipText, _
                                  strSexo, _
                                  txtTelefoneRes.Text, _
                                  txtCelular.Text, _
                                  txtRuaRes.Text, _
                                  txtNumeroRes.Text, _
                                  txtComplementoRes.Text, _
                                  txtEstadoRes.Text, _
                                  IIf(mskCepRes.ClipText = "", "", mskCepRes.ClipText), _
                                  txtBairroRes.Text, _
                                  txtCidadeRes.Text, _
                                  txtObservacao.Text
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir Pessoa
    objPessoa.InserirPessoa lngPKID, _
                                  IIf(mskDtNascimento.ClipText = "", "", mskDtNascimento.Text), _
                                  txtNome.Text, _
                                  txtIdentidade.Text, _
                                  txtOrgaoEmissor.Text, _
                                  IIf(mskDtExpedicao.ClipText = "", "", mskDtExpedicao.Text), _
                                  strTipoPessoa, _
                                  mskCnpj.ClipText, _
                                  mskCpf.ClipText, _
                                  strSexo, _
                                  txtTelefoneRes.Text, _
                                  txtCelular.Text, _
                                  txtRuaRes.Text, _
                                  txtNumeroRes.Text, _
                                  txtComplementoRes.Text, _
                                  txtEstadoRes.Text, _
                                  IIf(mskCepRes.ClipText = "", "", mskCepRes.ClipText), _
                                  txtBairroRes.Text, _
                                  txtCidadeRes.Text, _
                                  txtObservacao.Text

    'PESSOA
    'Set objRs = objPessoa.SelecionarPessoaPeloNome(txtNome.Text)
    'If Not objRs.EOF Then
    '  lngPKID = objRs.Fields("PKID").Value
    'End If
    'objRs.Close
    'Set objRs = Nothing
    '
  End If
  'Verificação
'''  If IcPessoa = tpIcPessoa_Pac Then
'''    'Paciente
'''    Set objPaciente = New busElite.clsPaciente
'''    'Verifica se paciente já cadastrado
'''    Set objRs = objPaciente.SelecionarPacientePeloPkid(lngPKID)
'''    If Not objRs.EOF Then
'''      'Paciente já cadastrado
'''      objPaciente.AlterarPaciente lngPKID
'''    Else
'''      'Paciente não cadastrado
'''      objPaciente.InserirPaciente lngPKID
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
'''    '
'''    Set objPaciente = Nothing
'''    '-----------------------------
'''    'Retorno para GR
'''    '-----------------------------
'''    If intQuemChamou = 1 Then
'''      'Chamada da GR
'''      'Retorna ao cadastro da GR
'''      frmUserGRCons.objUserGRInc.txtPessoaFim.Text = txtNome.Text
'''      INCLUIR_VALOR_NO_MASK frmUserGRCons.objUserGRInc.mskDataNascFim, mskDtNascimento.Text, TpMaskData
'''      blnRetorno = True
'''      blnFechar = True
'''      Unload Me
'''      Exit Sub
'''    End If
  If IcPessoa = tpIcPessoa_Func Then
    'Funcionario
    Set objFuncionario = New busElite.clsFuncionario
    'Verifica se Funcionario já cadastrado
    Set objRs = objFuncionario.SelecionarFuncionarioPeloPkid(lngPKID)
    If Not objRs.EOF Then
      'Funcionario já cadastrado
      objFuncionario.AlterarFuncionario lngPKID, _
                                      txtUsuario.Text, _
                                      Left(cboNivel.Text, 3), _
                                      Encripta(UCase$(txtNovaSenha.Text)), _
                                      strFuncExcluido
    Else
      'Funcionario não cadastrado
      objFuncionario.InserirFuncionario lngPKID, _
                                      txtUsuario.Text, _
                                      Left(cboNivel.Text, 3), _
                                      Encripta(UCase$(txtNovaSenha.Text)), _
                                      strFuncExcluido
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objFuncionario = Nothing
  ElseIf IcPessoa = tpIcPessoa_Mot Then
    'Motorista
    Set objMotorista = New busElite.clsMotorista
    'Verifica se Motorista já cadastrado
    Set objRs = objMotorista.SelecionarMotoristaPeloPkid(lngPKID)
    If Not objRs.EOF Then
      'Motorista já cadastrado
      objMotorista.AlterarMotorista lngPKID, _
                                    lngBANCOID, _
                                    txtAgencia.Text, _
                                    txtConta.Text, _
                                    strMotExcluido
    Else
      'Motorista não cadastrado
      objMotorista.InserirMotorista lngPKID, _
                                    lngBANCOID, _
                                    txtAgencia.Text, _
                                    txtConta.Text, _
                                    strMotExcluido
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objFuncionario = Nothing
  End If
  
  
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
'''    If IcPessoa = tpIcPessoa_Prest Then
'''      tabDetalhes.Tab = 2
'''    Else
      tabDetalhes.Tab = 0
'''    End If
    blnRetorno = True
  End If
  
  Set objPessoa = Nothing
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
  If Not Valida_Option(optTipoPessoa, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o tippo de pessoa" & vbCrLf
    tabDetalhes.Tab = 0
  End If
'''  If Len(Trim(mskCpf.ClipText)) = 0 Then
'''    strMsg = strMsg & "Informar o CPF" & vbCrLf
'''    Pintar_Controle mskCpf, tpCorContr_Erro
'''    SetarFoco mskCpf
'''    tabDetalhes.Tab = 0
'''    blnSetarFocoControle = False
'''  End If
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
  If Not Valida_Data(mskDtNascimento, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de nascimento válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Data(mskDtExpedicao, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de expedição válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If IcPessoa = tpIcPessoa.tpIcPessoa_Func Then
    If Not Valida_String(cboNivel, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Selecionar o nível do usuário" & vbCrLf
      tabDetalhes.Tab = 0
    End If
    If Left(cboNivel.Text, 3) <> "SEM" Then
      If Not Valida_String(txtUsuario, TpObrigatorio, blnSetarFocoControle) Then
        strMsg = strMsg & "Informar o nome do Usuario" & vbCrLf
        tabDetalhes.Tab = 0
      End If
      If Not Valida_String(txtNovaSenha, TpObrigatorio, blnSetarFocoControle) Then
        strMsg = strMsg & "Informar a senha" & vbCrLf
        tabDetalhes.Tab = 0
      End If
      If Not Valida_String(txtConfSenha, TpObrigatorio, blnSetarFocoControle) Then
        strMsg = strMsg & "Informar a confirmação da senha" & vbCrLf
        tabDetalhes.Tab = 0
      End If
      If Len(txtNovaSenha.Text) < 4 Then
        strMsg = strMsg & "Informar a nova Senha com mínimo de 4 dígitos" & vbCrLf
        tabDetalhes.Tab = 0
      End If
      If Len(txtNovaSenha.Text) < 4 Then
        strMsg = strMsg & "Informar a nova Senha com mínimo de 4 dígitos" & vbCrLf
        tabDetalhes.Tab = 0
      End If
      '
      If txtNovaSenha.Text <> txtConfSenha Then
        strMsg = strMsg & "Senhas digitadas não conferem" & vbCrLf
        tabDetalhes.Tab = 0
      End If
      '
    End If
    If Not Valida_Option(optFuncExcluido, blnSetarFocoControle) Then
      strMsg = strMsg & "Selecionar se o funcionário está excluido" & vbCrLf
      tabDetalhes.Tab = 0
    End If
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserPessoaInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserPessoaInc.ValidaCampos]", _
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
  TratarErro Err.Number, Err.Description, "[frmUserPessoaInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                       As ADODB.Recordset
  Dim strSql                      As String
  Dim objPessoa                   As busElite.clsPessoa
  Dim objFuncionario              As busElite.clsFuncionario
  Dim objMotorista                As busElite.clsMotorista
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
  '
  If IcPessoa = tpIcPessoa.tpIcPessoa_Func Then
    'Nivel
    cboNivel.Clear
    cboNivel.AddItem ""
    If gsNivel = gsAdmin Then _
      cboNivel.AddItem "ADMINISTRADOR"
  '  cboNivel.AddItem "ARQUIVISTA"
  '  cboNivel.AddItem "CAIXA"
  '  cboNivel.AddItem "DIRETOR"
  '  cboNivel.AddItem "FINANCEIRO"
    cboNivel.AddItem "GERENTE"
  '  cboNivel.AddItem "LABORATÓRIO"
  '  cboNivel.AddItem "SEM ACESSO AO SISTEMA"
  ElseIf IcPessoa = tpIcPessoa.tpIcPessoa_Mot Then
    'Banco
    strSql = "Select BANCO.NOME " & _
          " FROM BANCO " & _
          " ORDER BY BANCO.NOME"
    PreencheCombo cboBanco, strSql, False, True
  End If
  '
  
  tabDetalhes_Click 1
  If Status = tpStatus_Incluir Then
    optTipoPessoa(0).Value = True
    optSexo(0).Value = True
    If intQuemChamou = 1 Then
      txtNome.Text = strNomeInicial
    End If
    '
    tabDetalhes.TabVisible(2) = False
    tabDetalhes.TabVisible(3) = False
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    '-----------------------------
    'PESSOA
    '------------------------------
    Set objPessoa = New busElite.clsPessoa
    Set objRs = objPessoa.SelecionarPessoaPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      'Pessoa
      'Dados cadastrais
      txtNome.Text = objRs.Fields("NOME").Value & ""
      If objRs.Fields("TIPO_PESSOA").Value & "" = "F" Then
        optTipoPessoa(0).Value = True
        optTipoPessoa(1).Value = False
      ElseIf objRs.Fields("TIPO_PESSOA").Value & "" = "J" Then
        optTipoPessoa(0).Value = False
        optTipoPessoa(1).Value = True
      Else
        optTipoPessoa(0).Value = False
        optTipoPessoa(1).Value = False
      End If
      INCLUIR_VALOR_NO_MASK mskCpf, objRs.Fields("CPF").Value, TpMaskSemMascara
      INCLUIR_VALOR_NO_MASK mskCnpj, objRs.Fields("CNPJ").Value, TpMaskSemMascara
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
      INCLUIR_VALOR_NO_MASK mskDtNascimento, objRs.Fields("DTNASCIMENTO").Value, TpMaskData
      txtIdentidade.Text = objRs.Fields("RGNUMERO").Value & ""
      txtOrgaoEmissor.Text = objRs.Fields("RGORGAO").Value & ""
      txtOrgaoEmissor.Text = objRs.Fields("RGORGAO").Value & ""
      INCLUIR_VALOR_NO_MASK mskDtExpedicao, objRs.Fields("RGDTEXPEDICAO").Value, TpMaskData
      txtTelefoneRes.Text = objRs.Fields("TELEFONE").Value & ""
      txtCelular.Text = objRs.Fields("CELULAR").Value & ""
      txtObservacao.Text = objRs.Fields("OBSERVACAO").Value & ""
      'Endereço residencial
      txtRuaRes.Text = objRs.Fields("ENDRUA").Value & ""
      txtNumeroRes.Text = objRs.Fields("ENDNUMERO").Value & ""
      txtComplementoRes.Text = objRs.Fields("ENDCOMPLEMENTO").Value & ""
      txtEstadoRes.Text = objRs.Fields("ENDESTADO").Value & ""
      INCLUIR_VALOR_NO_MASK mskCepRes, objRs.Fields("ENDCEP").Value, TpMaskSemMascara
      txtBairroRes.Text = objRs.Fields("ENDBAIRRO").Value & ""
      txtCidadeRes.Text = objRs.Fields("ENDCIDADE").Value & ""
    End If
    objRs.Close
    Set objRs = Nothing
    Set objPessoa = Nothing
    If IcPessoa = tpIcPessoa.tpIcPessoa_Mot Then
      '-----------------------------
      'MOTORISTA
      '------------------------------
      Set objMotorista = New busElite.clsMotorista
      Set objRs = objMotorista.SelecionarMotoristaPeloPkid(lngPKID)
      '
      If Not objRs.EOF Then
        'MOTORISTA
        If objRs.Fields("NOME_BANCO").Value & "" <> "" Then
          cboBanco.Text = objRs.Fields("NOME_BANCO").Value & ""
        End If
        txtAgencia.Text = objRs.Fields("AGENCIA").Value & ""
        txtConta.Text = objRs.Fields("CONTA").Value & ""
        If objRs.Fields("INDEXCLUIDO").Value & "" = "S" Then
          optMotExcluido(0).Value = True
          optMotExcluido(1).Value = False
        ElseIf objRs.Fields("INDEXCLUIDO").Value & "" = "N" Then
          optMotExcluido(0).Value = False
          optMotExcluido(1).Value = True
        Else
          optMotExcluido(0).Value = False
          optMotExcluido(1).Value = False
        End If
        '
      End If
      objRs.Close
      Set objRs = Nothing
      Set objMotorista = Nothing
      '
    ElseIf IcPessoa = tpIcPessoa.tpIcPessoa_Func Then

      '-----------------------------
      'FUNCIONARIO
      '------------------------------
      Set objFuncionario = New busElite.clsFuncionario
      Set objRs = objFuncionario.SelecionarFuncionarioPeloPkid(lngPKID)
      '
      If Not objRs.EOF Then
        'Funcionario
        txtUsuario.Text = objRs.Fields("USUARIO").Value & ""
        txtNovaSenha.Text = Encripta(UCase$(objRs.Fields("SENHA").Value & "")) & ""
        txtConfSenha.Text = Encripta(UCase$(objRs.Fields("SENHA").Value & "")) & ""
        If objRs.Fields("DESCNIVEL").Value & "" <> "" Then
          cboNivel.Text = objRs.Fields("DESCNIVEL").Value & ""
        End If
        '
        If objRs.Fields("INDEXCLUIDO").Value & "" = "S" Then
          optFuncExcluido(0).Value = True
          optFuncExcluido(1).Value = False
        ElseIf objRs.Fields("INDEXCLUIDO").Value & "" = "N" Then
          optFuncExcluido(0).Value = False
          optFuncExcluido(1).Value = True
        Else
          optFuncExcluido(0).Value = False
          optFuncExcluido(1).Value = False
        End If
        '
      End If
      objRs.Close
      Set objRs = Nothing
      Set objFuncionario = Nothing
    End If
    tabDetalhes.TabVisible(2) = False
    tabDetalhes.TabVisible(3) = False
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




Private Sub grdProcedimento_UnboundReadDataEx( _
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
               Offset + intI, PROCED_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, PROCED_COLUNASMATRIZ, PROCED_LINHASMATRIZ, PROCED_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, PROCED_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserPessoaInc.grdGeral_UnboundReadDataEx]"
End Sub


Private Sub grdEspecialidade_UnboundReadDataEx( _
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
               Offset + intI, ESPEC_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, ESPEC_COLUNASMATRIZ, ESPEC_LINHASMATRIZ, ESPEC_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ESPEC_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserPessoaInc.grdGeral_UnboundReadDataEx]"
End Sub




Private Sub mskCepRes_GotFocus()
  Seleciona_Conteudo_Controle mskCepRes
End Sub
Private Sub mskCepRes_LostFocus()
  Pintar_Controle mskCepRes, tpCorContr_Normal
End Sub

Private Sub mskCnpj_GotFocus()
  Seleciona_Conteudo_Controle mskCnpj
End Sub
Private Sub mskCnpj_LostFocus()
  Pintar_Controle mskCnpj, tpCorContr_Normal
End Sub
Private Sub mskCPF_GotFocus()
  Seleciona_Conteudo_Controle mskCpf
End Sub
Private Sub mskCPF_LostFocus()
  Pintar_Controle mskCpf, tpCorContr_Normal
End Sub


Private Sub mskDtExpedicao_GotFocus()
  Seleciona_Conteudo_Controle mskDtExpedicao
End Sub
Private Sub mskDtExpedicao_LostFocus()
  Pintar_Controle mskDtExpedicao, tpCorContr_Normal
End Sub

Private Sub mskDtNascimento_GotFocus()
  Seleciona_Conteudo_Controle mskDtNascimento
End Sub
Private Sub mskDtNascimento_LostFocus()
  Pintar_Controle mskDtNascimento, tpCorContr_Normal
End Sub

Private Sub optTipoPessoa_Click(Index As Integer)
  On Error GoTo trata
  '
  Select Case Index
  Case 0
    'Pessoa Física
    Label5(4).Enabled = True
    mskCpf.Enabled = True
    '
    LimparCampoMask mskCnpj
    Label5(6).Enabled = False
    mskCnpj.Enabled = False
  Case 1
    'Pessoa Jurídica
    LimparCampoMask mskCpf
    Label5(4).Enabled = False
    mskCpf.Enabled = False
    '
    Label5(6).Enabled = True
    mskCnpj.Enabled = True
  End Select
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "frmUserPessoaInc.tabDetalhes"
  AmpN
End Sub



Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  Dim intTopAux As Integer
  On Error GoTo trata
  intTopAux = 2940
  Select Case tabDetalhes.Tab
  Case 0
    'Dados cadastrais
    grdEspecialidade.Enabled = False
    grdProcedimento.Enabled = False
    picTrava(0).Enabled = True
    If IcPessoa = tpIcPessoa.tpIcPessoa_Func Then
      'Funcionário
      picTrava(2).Top = intTopAux
      '
      picTrava(1).Visible = False
      picTrava(2).Visible = True
    ElseIf IcPessoa = tpIcPessoa.tpIcPessoa_Mot Then
      'Funcionário
      picTrava(1).Top = intTopAux
      '
      picTrava(1).Visible = True
      picTrava(2).Visible = False
    End If
    '
    picTrava(3).Enabled = False
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
    grdEspecialidade.Enabled = False
    grdProcedimento.Enabled = False
    picTrava(0).Enabled = False
    picTrava(1).Visible = False
    picTrava(2).Visible = False
    picTrava(3).Enabled = True
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    SetarFoco txtRuaRes
  Case 2
    'Especialidade
    grdEspecialidade.Enabled = True
    grdProcedimento.Enabled = False
    picTrava(0).Enabled = False
    picTrava(1).Visible = False
    picTrava(2).Visible = False
    picTrava(3).Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = False
    '
    '
    'Montar RecordSet
    ESPEC_COLUNASMATRIZ = grdEspecialidade.Columns.Count
    ESPEC_LINHASMATRIZ = 0
    ESPEC_MontaMatriz
    grdEspecialidade.Bookmark = Null
    grdEspecialidade.ReBind
    grdEspecialidade.ApproxCount = ESPEC_LINHASMATRIZ
    '
    SetarFoco grdEspecialidade
  Case 3
    'Procedimento
    grdEspecialidade.Enabled = False
    grdProcedimento.Enabled = True
    picTrava(0).Enabled = False
    picTrava(1).Visible = False
    picTrava(2).Visible = False
    picTrava(3).Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = True
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = True
    '
    '
    'Montar RecordSet
    PROCED_COLUNASMATRIZ = grdProcedimento.Columns.Count
    PROCED_LINHASMATRIZ = 0
    PROCED_MontaMatriz
    grdProcedimento.Bookmark = Null
    grdProcedimento.ReBind
    grdProcedimento.ApproxCount = PROCED_LINHASMATRIZ
    '
    SetarFoco grdProcedimento
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "frmUserPessoaInc.tabDetalhes"
  AmpN
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

Private Sub txtCidadeRes_GotFocus()
  Seleciona_Conteudo_Controle txtCidadeRes
End Sub
Private Sub txtCidadeRes_LostFocus()
  Pintar_Controle txtCidadeRes, tpCorContr_Normal
End Sub

Private Sub txtComplementoRes_GotFocus()
  Seleciona_Conteudo_Controle txtComplementoRes
End Sub
Private Sub txtComplementoRes_LostFocus()
  Pintar_Controle txtComplementoRes, tpCorContr_Normal
End Sub

Private Sub txtConfSenha_Gotfocus()
  Seleciona_Conteudo_Controle txtConfSenha
End Sub
Private Sub txtConfSenha_LostFocus()
  Pintar_Controle txtConfSenha, tpCorContr_Normal
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

Private Sub txtNome_GotFocus()
  Seleciona_Conteudo_Controle txtNome
End Sub
Private Sub txtNome_LostFocus()
  Pintar_Controle txtNome, tpCorContr_Normal
End Sub

Public Sub ESPEC_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busElite.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busElite.clsGeral
  '
  strSql = "SELECT PRESTADORESPECIALIDADE.PKID, ESPECIALIDADE.ESPECIALIDADE " & _
          "FROM PRESTADORESPECIALIDADE INNER JOIN ESPECIALIDADE ON ESPECIALIDADE.PKID = PRESTADORESPECIALIDADE.ESPECIALIDADEID " & _
          "WHERE PRESTADORESPECIALIDADE.PESSOAID = " & Formata_Dados(lngPKID, tpDados_Longo) & _
          " ORDER BY ESPECIALIDADE.ESPECIALIDADE"

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    ESPEC_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim ESPEC_Matriz(0 To ESPEC_COLUNASMATRIZ - 1, 0 To ESPEC_LINHASMATRIZ - 1)
  Else
    ReDim ESPEC_Matriz(0 To ESPEC_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To ESPEC_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To ESPEC_COLUNASMATRIZ - 1  'varre as colunas
          ESPEC_Matriz(intJ, intI) = objRs(intJ) & ""
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

Public Sub PROCED_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busElite.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busElite.clsGeral
  '
  strSql = "SELECT PRESTADORPROCEDIMENTO.PKID, PROCEDIMENTO.PROCEDIMENTO, PRESTADORPROCEDIMENTO.PERCCASA, PRESTADORPROCEDIMENTO.PERCPRESTADOR, PRESTADORPROCEDIMENTO.PERCRX, PRESTADORPROCEDIMENTO.PERCTECRX, PRESTADORPROCEDIMENTO.PERCULTRA " & _
          "FROM PRESTADORPROCEDIMENTO INNER JOIN PROCEDIMENTO ON PROCEDIMENTO.PKID = PRESTADORPROCEDIMENTO.PROCEDIMENTOID " & _
          "WHERE PRESTADORPROCEDIMENTO.PESSOAID = " & Formata_Dados(lngPKID, tpDados_Longo) & _
          " ORDER BY PROCEDIMENTO.PROCEDIMENTO"

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    PROCED_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim PROCED_Matriz(0 To PROCED_COLUNASMATRIZ - 1, 0 To PROCED_LINHASMATRIZ - 1)
  Else
    ReDim PROCED_Matriz(0 To PROCED_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To PROCED_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To PROCED_COLUNASMATRIZ - 1  'varre as colunas
          PROCED_Matriz(intJ, intI) = objRs(intJ) & ""
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

Private Sub txtNovaSenha_Gotfocus()
  Seleciona_Conteudo_Controle txtNovaSenha
End Sub
Private Sub txtNovaSenha_LostFocus()
  Pintar_Controle txtNovaSenha, tpCorContr_Normal
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

Private Sub txtRuaRes_GotFocus()
  Seleciona_Conteudo_Controle txtRuaRes
End Sub
Private Sub txtRuaRes_LostFocus()
  Pintar_Controle txtRuaRes, tpCorContr_Normal
End Sub

Private Sub txtTelefoneRes_GotFocus()
  Seleciona_Conteudo_Controle txtTelefoneRes
End Sub
Private Sub txtTelefoneRes_LostFocus()
  Pintar_Controle txtTelefoneRes, tpCorContr_Normal
End Sub


Private Sub txtUsuario_GotFocus()
  Seleciona_Conteudo_Controle txtUsuario
End Sub
Private Sub txtUsuario_LostFocus()
  Pintar_Controle txtUsuario, tpCorContr_Normal
End Sub

