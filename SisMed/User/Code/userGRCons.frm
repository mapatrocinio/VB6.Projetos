VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserGRCons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de GR�s"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraUnidade 
      Caption         =   "GR�s"
      Height          =   6015
      Left            =   60
      TabIndex        =   21
      Top             =   330
      Width           =   11835
      Begin TrueDBGrid60.TDBGrid grdGeral 
         Height          =   5730
         Left            =   90
         OleObjectBlob   =   "userGRCons.frx":0000
         TabIndex        =   0
         Top             =   180
         Width           =   11580
      End
   End
   Begin VB.CommandButton cmdInfFinanc 
      Caption         =   "&Z"
      Height          =   855
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6630
      Width           =   900
   End
   Begin VB.CommandButton cmdSairSelecao 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   855
      Left            =   11070
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6630
      Width           =   900
   End
   Begin VB.Frame fraImpressao 
      Caption         =   "Impress�o"
      Height          =   2085
      Left            =   7770
      TabIndex        =   19
      Top             =   6510
      Width           =   2355
      Begin VB.Label Label72 
         Caption         =   "CTRL + H - Consultar GR"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   7
         Left            =   90
         TabIndex        =   37
         Top             =   1680
         Width           =   2145
      End
      Begin VB.Label Label72 
         Caption         =   "CTRL + G - Consultar Proced."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   6
         Left            =   90
         TabIndex        =   36
         Top             =   1470
         Width           =   2145
      End
      Begin VB.Label Label72 
         Caption         =   "CTRL + F - Zerar senha"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   5
         Left            =   90
         TabIndex        =   35
         Top             =   1260
         Width           =   2145
      End
      Begin VB.Label Label72 
         Caption         =   "CTRL + E - Pesquisar prontu�rio"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   4
         Left            =   90
         TabIndex        =   34
         Top             =   1050
         Width           =   2145
      End
      Begin VB.Label Label72 
         Caption         =   "CTRL + D - Atualizar "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   3
         Left            =   90
         TabIndex        =   33
         Top             =   840
         Width           =   2145
      End
      Begin VB.Label Label72 
         Caption         =   "CTRL + B - Turno - Fechamento"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   2
         Left            =   90
         TabIndex        =   32
         Top             =   420
         Width           =   2175
      End
      Begin VB.Label Label72 
         Caption         =   "CTRL + C - Paciente"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   90
         TabIndex        =   26
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label72 
         Caption         =   "CTRL + A - Turno - Abre/Reimprime"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   210
         Width           =   2205
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Selecione a op��o"
      Height          =   1335
      Left            =   0
      TabIndex        =   18
      Top             =   6420
      Width           =   7665
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&I - Canc. outros      "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   3960
         TabIndex        =   9
         ToolTipText     =   "Cancelar GR outros prestadores"
         Top             =   630
         Width           =   1275
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&H - Comprov Rec  "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   2670
         TabIndex        =   8
         ToolTipText     =   "Comprovante Recebimento"
         Top             =   630
         Width           =   1275
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&F - Consultar GR    "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   60
         TabIndex        =   6
         ToolTipText     =   "Consultar GR"
         Top             =   630
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&G - Impress�o GR  "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   1380
         TabIndex        =   7
         ToolTipText     =   "impress�o da GR"
         Top             =   630
         Width           =   1275
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&A - Incluir GR            "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   1
         ToolTipText     =   "Incluir GR"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&B - Alterar GR          "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1350
         TabIndex        =   2
         ToolTipText     =   "Alterar GR"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&C - Itens da GR        "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   3
         ToolTipText     =   "�tens da GR"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&D - Receb. da GR    "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3930
         TabIndex        =   4
         ToolTipText     =   "Recebimento da GR"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&E - Cancelar GR      "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   5250
         TabIndex        =   5
         ToolTipText     =   "Cancelar GR"
         Top             =   240
         Width           =   1305
      End
      Begin ComctlLib.StatusBar StatusBar1 
         Height          =   255
         Left            =   60
         TabIndex        =   25
         Top             =   1020
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   450
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   5
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   6
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   1940
               MinWidth        =   1940
               TextSave        =   "9/9/2012"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   5
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   1235
               MinWidth        =   1235
               TextSave        =   "21:48"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   1
               Alignment       =   1
               Bevel           =   2
               Enabled         =   0   'False
               Object.Width           =   1235
               MinWidth        =   1235
               TextSave        =   "CAPS"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   2
               Alignment       =   1
               Bevel           =   2
               Enabled         =   0   'False
               Object.Width           =   1235
               MinWidth        =   1235
               TextSave        =   "NUM"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   3
               Alignment       =   1
               AutoSize        =   2
               Bevel           =   2
               Enabled         =   0   'False
               Object.Width           =   1244
               MinWidth        =   1235
               TextSave        =   "INS"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtTurno 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6660
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "txtTurno"
      Top             =   30
      Width           =   4785
   End
   Begin VB.TextBox txtUsuario 
      BackColor       =   &H00E0E0E0&
      Height          =   288
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "txtUsuario"
      Top             =   30
      Width           =   1695
   End
   Begin MSMask.MaskEdBox mskDataPrinc 
      Height          =   255
      Left            =   3990
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   60
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
      MaxLength       =   10
      Format          =   "DD/MMM/YYYY"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Atendida a posteriore"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   13
      Left            =   5940
      TabIndex        =   42
      Top             =   8100
      Width           =   1785
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "N�o Atendida"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   12
      Left            =   60
      TabIndex        =   41
      Top             =   8100
      Width           =   1185
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Liberada para atendimento"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   11
      Left            =   1290
      TabIndex        =   40
      Top             =   8100
      Width           =   2295
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Confirma��o de Expira��o"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   9
      Left            =   3630
      TabIndex        =   39
      Top             =   8100
      Width           =   2265
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Atendida"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   6270
      TabIndex        =   38
      Top             =   7830
      Width           =   675
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "Movimento ap�s o fechamento"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   7
      Left            =   3600
      TabIndex        =   31
      Top             =   7830
      Width           =   2625
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "N�o"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   1800
      TabIndex        =   30
      Top             =   8400
      Width           =   525
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Sim"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   1230
      TabIndex        =   29
      Top             =   8400
      Width           =   525
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      Caption         =   "Impress�o:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   60
      TabIndex        =   28
      Top             =   8400
      Width           =   915
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      Caption         =   "Status :"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   27
      Top             =   7830
      Width           =   765
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Cancelada"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   2520
      TabIndex        =   24
      Top             =   7830
      Width           =   1035
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Inicial"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   1230
      TabIndex        =   23
      Top             =   7830
      Width           =   525
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Fechada"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   1800
      TabIndex        =   22
      Top             =   7830
      Width           =   675
   End
   Begin VB.Label Label16 
      Caption         =   "Data"
      Height          =   255
      Left            =   3150
      TabIndex        =   17
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label21 
      Caption         =   "Turno Corrente"
      Height          =   255
      Left            =   5280
      TabIndex        =   16
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label Label22 
      Caption         =   "Usu�rio Logado"
      Height          =   255
      Left            =   30
      TabIndex        =   15
      Top             =   30
      Width           =   1215
   End
End
Attribute VB_Name = "frmUserGRCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''
Option Explicit

Public nGrupo                   As Integer
Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public objUserGRInc             As SisMed.frmUserGRInc
Public objUserContaCorrente     As SisMed.frmUserContaCorrente

Public blnPrimeiraVez           As Boolean 'Prop�sito: Preencher lista no combo

Private COLUNASMATRIZ           As Long
Private LINHASMATRIZ            As Long
Private Matriz()                As String


Public Sub Clique_botao(intIndice As Integer)
  If cmdSelecao(intIndice).Enabled = True And cmdSelecao(intIndice).Visible = True Then
    cmdSelecao_Click intIndice
  End If
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  'Tratamento de tecla para verifica��o de chamada de Outras telas
  'verifica se tem permiss�o
  'Tudo ok, faz chamada
  Select Case KeyAscii
  Case 1
    'TURNO - ABERTURA/REIMPRESS�O
    frmUserTurnoInc.Show vbModal 'Turno
    Form_Load
  Case 2
    'TURNO - FECHAMENTO
    FechamentoTurno
    Form_Load
  Case 3
    'PACIENTE
    frmUserProntuarioLis.IcProntuario = tpIcProntuario_Pac
    frmUserProntuarioLis.Show vbModal
    Form_Load
  Case 4
    'ATUALIZAR
    Form_Load
  Case 5
    'CONSULTAR PRONTU�RIO
    frmUserProntuarioGRCons.Show vbModal
    Form_Load
  Case 6
    'ZERAR SENHA
    frmUserZerarSenhaLis.Show vbModal
    Form_Load
  Case 7
    'CONSULTAR PROCEDIMENTO
    frmUserProcedimentoCons.indOrigem = 1
    frmUserProcedimentoCons.lngPRESTADORID = 0
    frmUserProcedimentoCons.Show vbModal
    Form_Load
  Case 8
    'CONSULTAR GR
    frmUserGRFinancCons.Show vbModal
    Form_Load
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserGRCons.Form_KeyPress]"
End Sub

Private Sub cmdInfFinanc_Click()
  On Error GoTo trata
  'Chamar o form de Consulta/Visualiza��o das Informa��es Financeiras.
  frmUserInfFinancLis.Show vbModal
  SetarFoco grdGeral
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

Private Sub cmdSairSelecao_Click()
  On Error GoTo trata
  blnFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

Public Sub cmdSelecao_Click(Index As Integer)
  On Error GoTo trata
  nGrupo = Index
  'strNumeroAptoPrinc = optUnidade
  'If Not ValiCamposPrinc Then Exit Sub
  VerificaQuemChamou
  'Atualiza Valores
  '
  COLUNASMATRIZ = grdGeral.Columns.Count
  LINHASMATRIZ = 0

  MontaMatriz
  grdGeral.Bookmark = Null
  grdGeral.ReBind
  grdGeral.ApproxCount = LINHASMATRIZ
  blnPrimeiraVez = False
  SetarFoco grdGeral
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[cmdSelecao_Click]"
  frmMDI.tmrUnidade.Enabled = True
End Sub


Public Sub VerificaQuemChamou()
  Dim objGR As busSisMed.clsGR
  Dim objGRTotalPrestCons As SisMed.frmUserGRTotalPrestCons
  Dim strMsg As String
  Dim objGeral      As busSisMed.clsGeral
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  On Error GoTo trata
  '
  Select Case nGrupo
  
  Case 0
    'Inclus�o da GR
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    
    Set objUserGRInc = New SisMed.frmUserGRInc
    objUserGRInc.Status = tpStatus_Incluir
    objUserGRInc.IcEstadoGR = tpIcEstadoGR_Inic
    objUserGRInc.lngGRID = 0
    objUserGRInc.Show vbModal
    Set objUserGRInc = Nothing
  Case 1
    'Altera��o da GR
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione uma GR para alter�-la.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(RetornaDescAtende(grdGeral.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
      MsgBox "Apenas o atendente que lan�ou a GR pode alter�-la.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "C" Then
      MsgBox "N�o pode haver altera��o em uma GR cancelada.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    
    If Trim(grdGeral.Columns("Status").Value & "") = "A" Then
      MsgBox "N�o pode haver altera��o em uma GR Atendida.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "N" Then
      MsgBox "N�o pode haver altera��o em uma GR n�o Atendida.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "L" Then
      MsgBox "N�o pode haver altera��o em uma GR Liberada para atendimento.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "E" Then
      MsgBox "N�o pode haver altera��o em uma GR com confirma��o de expira��o.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "P" Then
      MsgBox "N�o pode haver altera��o em uma GR atendida a posteriore.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    
    If Trim(grdGeral.Columns("Status").Value & "") = "F" Then
      'Pedir senha superior para alterar uma GR j� fechada
      '----------------------------
      '----------------------------
      'Pede Senha Superior (Diretor, Gerente ou Administrador
      If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
        'S� pede senha superior se quem estiver logado n�o for superior
        gsNomeUsuLib = ""
        gsNivelUsuLib = ""
        frmUserLoginSup.Show vbModal
        
        If Len(Trim(gsNomeUsuLib)) = 0 Then
          strMsg = "� necess�rio a confirma��o com senha superior para alterar uma GR."
          TratarErroPrevisto strMsg, "cmdConfirmar_Click"
          SetarFoco grdGeral
          Exit Sub
        Else
          'Capturou Nome do Usu�rio, continua com processo
        End If
      End If
      '--------------------------------
      '--------------------------------
    End If
    
    Set objUserGRInc = New SisMed.frmUserGRInc
    objUserGRInc.Status = tpStatus_Alterar
    objUserGRInc.IcEstadoGR = tpIcEstadoGR_Inic
    objUserGRInc.lngGRID = grdGeral.Columns("ID").Value
    objUserGRInc.Show vbModal
    Set objUserGRInc = Nothing
  Case 2
    'Itens da GR
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione uma GR para alterar seus �tens.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(RetornaDescAtende(grdGeral.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
      MsgBox "Apenas o atendente que lan�ou a GR pode alter�-la.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "C" Then
      MsgBox "N�o pode haver altera��o de �tens de uma GR cancelada.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    
    If Trim(grdGeral.Columns("Status").Value & "") = "A" Then
      MsgBox "N�o pode haver altera��o em uma GR Atendida.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "N" Then
      MsgBox "N�o pode haver altera��o em uma GR n�o Atendida.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "L" Then
      MsgBox "N�o pode haver altera��o em uma GR Liberada para atendimento.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "E" Then
      MsgBox "N�o pode haver altera��o em uma GR com confirma��o de expira��o.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "P" Then
      MsgBox "N�o pode haver altera��o em uma GR atendida a posteriore.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "F" Then
      'Pedir senha superior para alterar uma GR j� fechada
      '----------------------------
      '----------------------------
      'Pede Senha Superior (Diretor, Gerente ou Administrador
      If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
        'S� pede senha superior se quem estiver logado n�o for superior
        gsNomeUsuLib = ""
        gsNivelUsuLib = ""
        frmUserLoginSup.Show vbModal
        
        If Len(Trim(gsNomeUsuLib)) = 0 Then
          strMsg = "� necess�rio a confirma��o com senha superior para alterar uma GR."
          TratarErroPrevisto strMsg, "cmdConfirmar_Click"
          SetarFoco grdGeral
          Exit Sub
        Else
          'Capturou Nome do Usu�rio, continua com processo
        End If
      End If
      '--------------------------------
      '--------------------------------
    End If
    Set objUserGRInc = New SisMed.frmUserGRInc
    objUserGRInc.Status = tpStatus_Alterar
    objUserGRInc.IcEstadoGR = tpIcEstadoGR_Proc
    objUserGRInc.lngGRID = grdGeral.Columns("ID").Value
    objUserGRInc.Show vbModal
    Set objUserGRInc = Nothing
  Case 3
    'Altera��o de pagamento da GR
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione uma GR para alterar seus dados de pagamento.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(RetornaDescAtende(grdGeral.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
      If (gsNivel <> gsLaboratorio) And (Trim(RetornaNivelAtende(grdGeral.Columns("Atendente").Value & "")) <> gsLaboratorio) Then
        MsgBox "Apenas o atendente que lan�ou a GR pode efetuar o seu pagamento ou uma GR lan�ada pelo Laborat�rio.", vbExclamation, TITULOSISTEMA
        SetarFoco grdGeral
        Exit Sub
      End If
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "C" Then
      MsgBox "N�o pode haver pagamento de uma GR cancelada.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    
    If Trim(grdGeral.Columns("Status").Value & "") = "A" Then
      MsgBox "N�o pode haver pagamento em uma GR Atendida.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "N" Then
      MsgBox "N�o pode haver pagamento em uma GR n�o Atendida.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "L" Then
      MsgBox "N�o pode haver pagamento em uma GR Liberada para atendimento.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "E" Then
      MsgBox "N�o pode haver pagamento em uma GR com confirma��o de expira��o.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "P" Then
      MsgBox "N�o pode haver pagamento em uma GR atendida a posteriore.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    
    If Trim(grdGeral.Columns("Status").Value & "") = "F" Then
      'Pedir senha superior para alterar uma GR j� fechada
      '----------------------------
      '----------------------------
      'Pede Senha Superior (Diretor, Gerente ou Administrador
      If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
        'S� pede senha superior se quem estiver logado n�o for superior
        gsNomeUsuLib = ""
        gsNivelUsuLib = ""
        frmUserLoginSup.Show vbModal
        
        If Len(Trim(gsNomeUsuLib)) = 0 Then
          strMsg = "� necess�rio a confirma��o com senha superior para alterar pagamento de uma GR fechada."
          TratarErroPrevisto strMsg, "cmdConfirmar_Click"
          SetarFoco grdGeral
          Exit Sub
        Else
          'Capturou Nome do Usu�rio, continua com processo
        End If
      End If
      '--------------------------------
      '--------------------------------
    End If
    Set objUserContaCorrente = New frmUserContaCorrente
    objUserContaCorrente.lngCCID = 0
    objUserContaCorrente.lngGRID = grdGeral.Columns("ID").Value
    objUserContaCorrente.intGrupo = 0
    objUserContaCorrente.strFuncionarioNome = gsNomeUsuCompleto
    objUserContaCorrente.Status = tpStatus_Incluir
    objUserContaCorrente.strStatusLanc = "RC"
    objUserContaCorrente.strNivelAcesso = Trim(RetornaNivelAtende(grdGeral.Columns("Atendente").Value & ""))
    objUserContaCorrente.Show vbModal
    Set objUserContaCorrente = Nothing
  Case 4
    'Cancelamento da GR
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione uma GR para exclu�-la.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(RetornaDescAtende(grdGeral.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
      If Mid(grdGeral.Columns("Atendente").Value & "", 2, 3) <> gsLaboratorio Then
        MsgBox "Apenas o atendente que lan�ou a GR pode exclu�-la.", vbExclamation, TITULOSISTEMA
        SetarFoco grdGeral
        Exit Sub
      End If
    End If
    
    If Trim(grdGeral.Columns("Status").Value & "") = "A" Then
      MsgBox "N�o pode haver cancelamento em uma GR Atendida.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "N" Then
      MsgBox "N�o pode haver cancelamento em uma GR n�o Atendida.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "L" Then
      MsgBox "N�o pode haver cancelamento em uma GR Liberada para atendimento.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "E" Then
      MsgBox "N�o pode haver cancelamento em uma GR com confirma��o de expira��o.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "P" Then
      MsgBox "N�o pode haver cancelamento em uma GR atendida a posteriore.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    
    'Verifica GR j� cancelada
    Set objGeral = New busSisMed.clsGeral
    strSql = "SELECT GR.PKID, GR.STATUS FROM GR " & _
        " WHERE GR.PKID = " & Formata_Dados(grdGeral.Columns("ID").Value, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      If objRs.Fields("STATUS").Value & "" = "C" Then
        'GR CANCELADA
        objRs.Close
        Set objRs = Nothing
        Set objGeral = Nothing
        MsgBox "GR j� cancelada.", vbExclamation, TITULOSISTEMA
        SetarFoco grdGeral
        Exit Sub
      End If
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    
    
    
    
    
    'If Trim(grdGeral.Columns("Status").Value & "") <> "F" Then
    '  MsgBox "Apenas pode de excluida uma GR fechada.", vbExclamation, TITULOSISTEMA
    '  SetarFoco grdGeral
    '  Exit Sub
    'End If
    'If Trim(grdGeral.Columns("Status").Value & "") = "F" Then
      'Pedir senha superior para alterar uma GR j� fechada
      '----------------------------
      '----------------------------
      'Pede Senha Superior (Diretor, Gerente ou Administrador
      gsNomeUsuLib = ""
      gsNivelUsuLib = ""
      If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
        'S� pede senha superior se quem estiver logado n�o for superior
        frmUserLoginSup.Show vbModal
        
        If Len(Trim(gsNomeUsuLib)) = 0 Then
          strMsg = "� necess�rio a confirma��o com senha superior para cancelar uma GR."
          TratarErroPrevisto strMsg, "cmdConfirmar_Click"
          SetarFoco grdGeral
          Exit Sub
        Else
          'Capturou Nome do Usu�rio, continua com processo
        End If
      Else
        gsNomeUsuLib = gsNomeUsu
        gsNivelUsuLib = gsNivel
      End If
      '--------------------------------
      '--------------------------------
    'End If
    'Confirma��o
    If MsgBox("Confirma cancelamento da GR " & grdGeral.Columns("Seq.").Value & " de " & grdGeral.Columns("Prontu�rio").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
      SetarFoco grdGeral
      Exit Sub
    End If
    
    Set objGR = New busSisMed.clsGR
    objGR.AlterarStatusGR grdGeral.Columns("ID").Value, _
                          "C", _
                          "", _
                          RetornaCodTurnoCorrente, _
                          gsNomeUsuLib
    Set objGR = Nothing
    IMP_COMP_CANC_GR grdGeral.Columns("ID").Value, gsNomeEmpresa, 1
  Case 5
    'consultar GR
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione uma GR para alter�-la.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(RetornaDescAtende(grdGeral.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
      MsgBox "Apenas o atendente que lan�ou a GR pode consult�-la.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    Set objUserGRInc = New SisMed.frmUserGRInc
    objUserGRInc.Status = tpStatus_Consultar
    objUserGRInc.IcEstadoGR = tpIcEstadoGR_Con
    objUserGRInc.lngGRID = grdGeral.Columns("ID").Value
    objUserGRInc.Show vbModal
    Set objUserGRInc = Nothing
  Case 6
    'Imprimir GR
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione uma GR para imprim�-la.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(RetornaNivelAtende(grdGeral.Columns("Atendente").Value & "")) = gsLaboratorio Then
      MsgBox "N�o pode haver impress�o de uma GR lan�ada pelo Laborat�rio.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") <> "F" Then
      If Trim(grdGeral.Columns("Status").Value & "") <> "A" Then
        MsgBox "N�o pode haver impress�o de uma GR que n�o esteja fechada ou atendida.", vbExclamation, TITULOSISTEMA
        SetarFoco grdGeral
        Exit Sub
      End If
    End If
    'Confirma��o
    If MsgBox("Confirma impress�o da GR " & grdGeral.Columns("Seq.").Value & " de " & grdGeral.Columns("Prontu�rio").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Imp").Value & "") = "S" Then
      'Pedir senha superior para imprimir uma GR j� impressa
      '----------------------------
      '----------------------------
      'Pede Senha Superior (Diretor, Gerente ou Administrador
      If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
        'S� pede senha superior se quem estiver logado n�o for superior
        gsNomeUsuLib = ""
        gsNivelUsuLib = ""
        frmUserLoginSup.Show vbModal
        
        If Len(Trim(gsNomeUsuLib)) = 0 Then
          strMsg = "� necess�rio a confirma��o com senha superior para imprimir uma GR j� impressa."
          TratarErroPrevisto strMsg, "cmdConfirmar_Click"
          SetarFoco grdGeral
          Exit Sub
        Else
          'Capturou Nome do Usu�rio, continua com processo
        End If
      End If
      '--------------------------------
      '--------------------------------
    End If
    
    IMP_COMP_GR grdGeral.Columns("ID").Value, gsNomeEmpresa, 1, IIf(Trim(grdGeral.Columns("Imp").Value & "") = "S", True, False)
    'Ap�s impress�o altera status para impressa
    Set objGR = New busSisMed.clsGR
    objGR.AlterarStatusGR grdGeral.Columns("ID").Value, _
                          "", _
                          "S"
    Set objGR = Nothing
  
  Case 7
    'Imprimir Comprovante de Recebimento
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione uma GR para imprim�-la.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") <> "F" Then
      If Trim(grdGeral.Columns("Status").Value & "") <> "A" Then
        MsgBox "N�o pode haver impress�o de uma GR que n�o esteja fechada ou atendida.", vbExclamation, TITULOSISTEMA
        SetarFoco grdGeral
        Exit Sub
      End If
    End If
    'Confirma��o
    If MsgBox("Confirma impress�o da GR " & grdGeral.Columns("Seq.").Value & " de " & grdGeral.Columns("Prontu�rio").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
      SetarFoco grdGeral
      Exit Sub
    End If
    
    'Fecou GR do laborat�rio, emitir comprovante de pagamento
    IMP_COMPROV_REC grdGeral.Columns("ID").Value, gsNomeEmpresa, 1
  Case 8
    'Canlelar GR outros prestadores
    Set objGRTotalPrestCons = New SisMed.frmUserGRTotalPrestCons
    objGRTotalPrestCons.Show vbModal
    Set objGRTotalPrestCons = Nothing
  End Select
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  End
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim strSql            As String
  Dim datDataTurno      As Date
  Dim datDataIniAtual   As Date
  Dim datDataFimAtual   As Date
  '
  If RetornaCodTurnoCorrente(datDataTurno) = 0 Then
    TratarErroPrevisto "N�o h� turnos em aberto, favor abrir um turno antes de iniciair as GR�s", "Form_Load"
  Else
    'OK Para turno
    datDataIniAtual = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now) & " 00:00:00")
    datDataFimAtual = DateAdd("d", 1, datDataIniAtual)
    If datDataTurno < datDataIniAtual Or datDataTurno >= datDataFimAtual Then
      TratarErroPrevisto "ATEN��O" & vbCrLf & vbCrLf & "A data do turno atual aberto n�o corresponde a data de hoje:" & vbCrLf & vbCrLf & "Data do turno --> " & Format(datDataTurno, "DD/MM/YYYY") & vbCrLf & "Data Atual --> " & Format(datDataIniAtual, "DD/MM/YYYY") & vbCrLf & vbCrLf & "Por favor, feche o turno e abra-o novamente. Caso voc� n�o realize esta opera��o, as GR�S lan�adas n�o ser�o exibidas na consulta.", "Form_Load"
    End If
  End If
  
  If gsNivel = gsLaboratorio Then
    cmdSelecao(3).Enabled = False
    cmdSelecao(7).Enabled = False
  Else
    cmdSelecao(3).Enabled = True
    cmdSelecao(7).Enabled = True
  End If
  blnFechar = False
  blnRetorno = False
  blnPrimeiraVez = True
  AmpS
  If Me.ActiveControl Is Nothing Then
    Me.Top = 580
    Me.Left = 1
    Me.WindowState = 2 'Maximizado
  End If
  'Me.Height = 9195
  'Me.Width = 12090
  'CenterForm Me
  LerFigurasAvulsas cmdSairSelecao, "Sair.ico", "SairDown.ico", "Sair"
  LerFigurasAvulsas cmdInfFinanc, "InfFinanc.ico", "InfFinancDown.ico", "Informa��es financeiras do turno"
  '
  txtTurno.Text = RetornaDescTurnoCorrente
  txtUsuario.Text = gsNomeUsu
  mskDataPrinc.Text = Format(Date, "DD/MM/YYYY")

  'NOVO BOT�ES NOVOS
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
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
  TratarErro Err.Number, Err.Description, "[frmUserGRCons.grdGeral_UnboundReadDataEx]"
End Sub



Private Sub Form_Activate()
  If blnPrimeiraVez Then
    DoEvents
    '
    COLUNASMATRIZ = grdGeral.Columns.Count
    LINHASMATRIZ = 0

    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = LINHASMATRIZ
    blnPrimeiraVez = False
    SetarFoco grdGeral
  End If
End Sub

Public Sub MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGR     As busSisMed.clsGR
  '
  AmpS
  On Error GoTo trata
  '
  Set objGR = New busSisMed.clsGR
  '
  Set objRs = objGR.CapturaGRTurnoCorrente(RetornaCodTurnosPelaData(Now, gsNivel, giFuncionarioId), _
                                                                    giFuncionarioId)
  If Not objRs.EOF Then
    LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim Matriz(0 To COLUNASMATRIZ - 1, 0 To LINHASMATRIZ - 1)
  Else
    ReDim Matriz(0 To COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se j� houver algum item
    For intI = 0 To LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda n�o se atingiu fim do recordset
        For intJ = 0 To COLUNASMATRIZ - 1  'varre as colunas
          Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'pr�xima linha matriz
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGR = Nothing
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub
