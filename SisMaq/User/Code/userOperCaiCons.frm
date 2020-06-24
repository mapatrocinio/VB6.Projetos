VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserOperCaiCons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operacional Caixa"
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
      Caption         =   "Operacional Caixa"
      Height          =   6015
      Left            =   60
      TabIndex        =   40
      Top             =   360
      Width           =   11835
      Begin VB.TextBox txtLP 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9060
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   210
         Width           =   900
      End
      Begin VB.TextBox txtSaldoAtend 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   4470
         TabIndex        =   5
         Top             =   210
         Width           =   900
      End
      Begin VB.TextBox txtEmpTroca 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   9060
         TabIndex        =   13
         Top             =   510
         Width           =   900
      End
      Begin VB.TextBox txtDespesaTotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   4470
         TabIndex        =   10
         Top             =   510
         Width           =   900
      End
      Begin VB.TextBox txtEmCaixa 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   10800
         TabIndex        =   14
         Top             =   510
         Width           =   900
      End
      Begin VB.TextBox txtSangria 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   7500
         TabIndex        =   12
         Top             =   510
         Width           =   900
      End
      Begin VB.TextBox txtSaiSaldo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   5940
         TabIndex        =   11
         Top             =   510
         Width           =   900
      End
      Begin VB.TextBox txtDespesa 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2730
         TabIndex        =   9
         Top             =   510
         Width           =   900
      End
      Begin VB.TextBox txtSaiAtendente 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   8
         Top             =   510
         Width           =   900
      End
      Begin VB.TextBox txtEntradaTotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   7500
         TabIndex        =   6
         Top             =   210
         Width           =   900
      End
      Begin VB.TextBox txtEntradaArrec 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2730
         TabIndex        =   4
         Top             =   210
         Width           =   900
      End
      Begin VB.TextBox txtEntradaCaixa 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   3
         Top             =   210
         Width           =   900
      End
      Begin TrueDBGrid60.TDBGrid grdEntArrec 
         Height          =   1740
         Left            =   90
         OleObjectBlob   =   "userOperCaiCons.frx":0000
         TabIndex        =   16
         Top             =   2550
         Width           =   5820
      End
      Begin TrueDBGrid60.TDBGrid grdEntrada 
         Height          =   1740
         Left            =   90
         OleObjectBlob   =   "userOperCaiCons.frx":4F05
         TabIndex        =   15
         Top             =   810
         Width           =   5820
      End
      Begin TrueDBGrid60.TDBGrid grdRetirada 
         Height          =   1740
         Left            =   5910
         OleObjectBlob   =   "userOperCaiCons.frx":94E9
         TabIndex        =   19
         Top             =   4290
         Width           =   5820
      End
      Begin TrueDBGrid60.TDBGrid grdEntradaAtend 
         Height          =   1740
         Left            =   5910
         OleObjectBlob   =   "userOperCaiCons.frx":DAC2
         TabIndex        =   17
         Top             =   810
         Width           =   5820
      End
      Begin TrueDBGrid60.TDBGrid grdDespesa 
         Height          =   1740
         Left            =   5910
         OleObjectBlob   =   "userOperCaiCons.frx":1266B
         TabIndex        =   18
         Top             =   2550
         Width           =   5820
      End
      Begin VB.Label Label12 
         Caption         =   "L/P"
         Height          =   165
         Left            =   8430
         TabIndex        =   59
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label11 
         Caption         =   "Saldo Ate."
         Height          =   195
         Left            =   3660
         TabIndex        =   58
         Top             =   210
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "Empr."
         Height          =   225
         Left            =   8430
         TabIndex        =   57
         Top             =   510
         Width           =   765
      End
      Begin VB.Label Label7 
         Caption         =   "Desp.Total"
         Height          =   225
         Left            =   3660
         TabIndex        =   56
         Top             =   510
         Width           =   765
      End
      Begin VB.Label Label4 
         Caption         =   "Em Caixa"
         Height          =   165
         Left            =   10020
         TabIndex        =   55
         Top             =   510
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Total"
         Height          =   165
         Left            =   6870
         TabIndex        =   54
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label8 
         Caption         =   "Sangria"
         Height          =   225
         Left            =   6870
         TabIndex        =   53
         Top             =   510
         Width           =   765
      End
      Begin VB.Label Label9 
         Caption         =   "Saldo"
         Height          =   165
         Left            =   5430
         TabIndex        =   52
         Top             =   510
         Width           =   525
      End
      Begin VB.Label Label6 
         Caption         =   "Despesa"
         Height          =   165
         Left            =   1830
         TabIndex        =   51
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Saldo Arrec."
         Height          =   195
         Left            =   1830
         TabIndex        =   50
         Top             =   210
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Atendente"
         Height          =   165
         Left            =   120
         TabIndex        =   49
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "Caixa"
         Height          =   165
         Left            =   120
         TabIndex        =   48
         Top             =   210
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdSairSelecao 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   855
      Left            =   11010
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6540
      Width           =   900
   End
   Begin VB.Frame fraImpressao 
      Caption         =   "Impressão"
      Height          =   2085
      Left            =   7770
      TabIndex        =   38
      Top             =   6420
      Width           =   2355
      Begin VB.Label Label72 
         Caption         =   "CTRL + G - Detalhar Emp./Troca"
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
         TabIndex        =   47
         Top             =   1470
         Width           =   2145
      End
      Begin VB.Label Label72 
         Caption         =   "CTRL + F - Detalhar Boleto Arrec."
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
         TabIndex        =   46
         Top             =   1260
         Width           =   2145
      End
      Begin VB.Label Label72 
         Caption         =   "CTRL + E - Detalhar Boleto Atend."
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
         TabIndex        =   45
         Top             =   1050
         Width           =   2145
      End
      Begin VB.Label Label72 
         Caption         =   "CTRL + D - Detalhar Entrada Atend."
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
         TabIndex        =   44
         Top             =   840
         Width           =   2205
      End
      Begin VB.Label Label72 
         Caption         =   "CTRL + C - Detalhar Retirada"
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
         TabIndex        =   43
         Top             =   630
         Width           =   2145
      End
      Begin VB.Label Label72 
         Caption         =   "CTRL + B - Detalhar Entrada"
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
         TabIndex        =   42
         Top             =   420
         Width           =   2145
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
         TabIndex        =   39
         Top             =   210
         Width           =   2205
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Selecione a opção"
      Height          =   2085
      Left            =   60
      TabIndex        =   37
      Top             =   6420
      Width           =   7665
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&M - Leit. Especial    "
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
         Index           =   12
         Left            =   2700
         TabIndex        =   32
         ToolTipText     =   "Detalhar Arrecadador"
         Top             =   1020
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&L - Detalhe Arrec     "
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
         Index           =   11
         Left            =   1380
         TabIndex        =   31
         ToolTipText     =   "Detalhar Arrecadador"
         Top             =   1020
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&J - Atualizar               "
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
         Index           =   9
         Left            =   5280
         TabIndex        =   29
         ToolTipText     =   "Atualizar Tela"
         Top             =   630
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&K - Detalhe Atend    "
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
         Index           =   10
         Left            =   60
         TabIndex        =   30
         ToolTipText     =   "Detalhar Atendente"
         Top             =   1020
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&B - Retirada               "
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
         Left            =   1380
         TabIndex        =   21
         ToolTipText     =   "Retirada"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&A - Entrada                "
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
         TabIndex        =   20
         ToolTipText     =   "Entrada"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&I - Fech. Caixa          "
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
         Left            =   4020
         TabIndex        =   28
         ToolTipText     =   "Cancelar GR outros prestadores"
         Top             =   630
         Width           =   1245
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&H - Fech. Arrec.      "
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
         Left            =   2700
         TabIndex        =   27
         ToolTipText     =   "Comprovante Recebimento"
         Top             =   630
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&F - Emp./Troca        "
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
         TabIndex        =   25
         ToolTipText     =   "Consultar GR"
         Top             =   630
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&G - Fech. Atend.      "
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
         TabIndex        =   26
         ToolTipText     =   "impressão da GR"
         Top             =   630
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&C - Ent. Atendente"
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
         Left            =   2700
         TabIndex        =   22
         ToolTipText     =   "Entrada do atendente"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&D - Boleto Atend.   "
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
         Left            =   3960
         TabIndex        =   23
         ToolTipText     =   "Boleto do atendente"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&E - Boleto Arrec.     "
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
         Left            =   5280
         TabIndex        =   24
         ToolTipText     =   "Boleto do arrecadador"
         Top             =   240
         Width           =   1305
      End
      Begin ComctlLib.StatusBar StatusBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1740
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
               TextSave        =   "30/1/2011"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   5
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   1235
               MinWidth        =   1235
               TextSave        =   "20:52"
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
      TabIndex        =   2
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
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "txtUsuario"
      Top             =   30
      Width           =   1695
   End
   Begin MSMask.MaskEdBox mskDataPrinc 
      Height          =   255
      Left            =   3990
      TabIndex        =   1
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
   Begin VB.Label Label16 
      Caption         =   "Data"
      Height          =   255
      Left            =   3150
      TabIndex        =   36
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label21 
      Caption         =   "Turno Corrente"
      Height          =   255
      Left            =   5280
      TabIndex        =   35
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label Label22 
      Caption         =   "Usuário Logado"
      Height          =   255
      Left            =   30
      TabIndex        =   34
      Top             =   30
      Width           =   1215
   End
End
Attribute VB_Name = "frmUserOperCaiCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public nGrupo                         As Integer
Public Status                   As tpStatus
Public blnRetorno                     As Boolean
Public blnFechar                      As Boolean
Public lngTURNOATENDEPESQ             As Long
Public lngTURNOARRECEPESQ             As Long
'''
'''Public objUserGRInc             As SisMaq.frmUserGRInc
'''Public objUserContaCorrente     As SisMaq.frmUserContaCorrente
'''
Public blnPrimeiraVez                 As Boolean 'Propósito: Preencher lista no combo

'Entrada
Private ENTR_COLUNASMATRIZ            As Long
Private ENTR_LINHASMATRIZ             As Long
Private ENTR_Matriz()                 As String
'Retirada
Private RET_COLUNASMATRIZ             As Long
Private RET_LINHASMATRIZ              As Long
Private RET_Matriz()                  As String
'Entrada Atendente
Private ENTRAT_COLUNASMATRIZ            As Long
Private ENTRAT_LINHASMATRIZ             As Long
Private ENTRAT_Matriz()                 As String

'Entrada Arrecadador
Private ENTARR_COLUNASMATRIZ            As Long
Private ENTARR_LINHASMATRIZ             As Long
Private ENTARR_Matriz()                 As String
'Saída Atendente
Private DESP_COLUNASMATRIZ            As Long
Private DESP_LINHASMATRIZ             As Long
Private DESP_Matriz()                 As String


Public Sub Clique_botao(intIndice As Integer)
  If cmdSelecao(intIndice).Enabled = True And cmdSelecao(intIndice).Visible = True Then
    cmdSelecao_Click intIndice
  End If
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  Dim blnReload As Boolean
  'Tratamento de tecla para verificação de chamada de Outras telas
  'verifica se tem permissão
  'Tudo ok, faz chamada
  blnReload = False
  Select Case KeyAscii
  Case 1
    'TURNO - ABERTURA/REIMPRESSÃO
    frmUserTurnoInc.Show vbModal 'Turno
    Form_Load
    blnReload = True
'''  Case 2
'''    'TURNO - FECHAMENTO
'''    FechamentoTurno
'''    Form_Load
'''    blnReload = True
  Case 2
    'DETALHAR ENTRADA
    frmUserEntradaLis.Show vbModal
    Form_Load
    blnReload = True
  Case 3
    'DETALHAR RETIRADA
    frmUserRetiradaLis.Show vbModal
    Form_Load
    blnReload = True
  Case 4
    'DETALHAR ENTRADA ATENDENTE
    frmUserEntradaAtendLis.Show vbModal
    Form_Load
    blnReload = True
  Case 5
    'DETALHAR BOLETO ATENDENTE
    frmUserBoletoAtendLis.Show vbModal
    Form_Load
    blnReload = True
  Case 6
    'DETALHAR BOLETO ARRECADADOR
    frmUserBoletoArrecLis.Show vbModal
    Form_Load
    blnReload = True
  Case 7
    'DETALHAR EMPRÉSTIMO/TROCA
    frmUserEmpTrocaLis.Show vbModal
    Form_Load
    blnReload = True
  End Select
  '
  If blnReload = True Then
    Trata_Matrizes_Totais
    SetarFoco grdEntrada
  Else
  End If
  '
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserOperCaiCons.Form_KeyPress]"
End Sub

'''Private Sub cmdInfFinanc_Click()
'''  On Error GoTo trata
'''  'Chamar o form de Consulta/Visualização das Informações Financeiras.
'''  frmUserInfFinancLis.Show vbModal
'''  SetarFoco grdEntrada
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, _
'''             Err.Description, _
'''             Err.Source
'''  AmpN
'''End Sub
'''
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
  Trata_Matrizes_Totais
  SetarFoco grdEntrada
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[cmdSelecao_Click]"
  frmMDI.tmrUnidade.Enabled = True
End Sub


Public Sub VerificaQuemChamou()
  Dim objUserEntradaInc As SisMaq.frmUserEntradaInc
  Dim objUserRetiradaInc As SisMaq.frmUserRetiradaInc
  Dim objUserEntradaAtendInc As SisMaq.frmUserEntradaAtendInc
  Dim objUserBoletoAtendInc As SisMaq.frmUserBoletoAtendInc
  Dim objUserBoletoArrecInc As SisMaq.frmUserBoletoArrecInc
  Dim objUserEmpTrocaInc As SisMaq.frmUserEmpTrocaInc
  Dim objUserFechaAteCons As SisMaq.frmUserFechaAteCons
  Dim objUserFechaArrCons As SisMaq.frmUserFechaArrCons
  Dim objUserDetalheAtend As SisMaq.frmUserDetalhamentoAtend
  Dim objUserDetalheArrec As SisMaq.frmUserDetalhamentoArrec
  Dim objUserLeituraFechaInc As SisMaq.frmUserLeituraFechaInc
  '
  Dim strMsg As String
  On Error GoTo trata
  '
  Select Case nGrupo

  Case 0
    'Entrada
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "Não há turno aberto. Por favor abra o turno antes de iniciar as atividades.", vbExclamation, TITULOSISTEMA
      SetarFoco grdEntrada
      Exit Sub
    End If

    Set objUserEntradaInc = New SisMaq.frmUserEntradaInc
    objUserEntradaInc.Status = tpStatus_Incluir
    objUserEntradaInc.lngPKID = 0
    objUserEntradaInc.Show vbModal
    Set objUserEntradaInc = Nothing
  Case 1
    'Retirada
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "Não há turno aberto. Por favor abra o turno antes de iniciar as atividades.", vbExclamation, TITULOSISTEMA
      SetarFoco grdEntrada
      Exit Sub
    End If

    Set objUserRetiradaInc = New SisMaq.frmUserRetiradaInc
    objUserRetiradaInc.Status = tpStatus_Incluir
    objUserRetiradaInc.curSaldo = CCur(txtEmCaixa)
    objUserRetiradaInc.lngPKID = 0
    objUserRetiradaInc.Show vbModal
    Set objUserRetiradaInc = Nothing
  Case 2
    'Caixa Atend
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "Não há turno aberto. Por favor abra o turno antes de iniciar as atividades.", vbExclamation, TITULOSISTEMA
      SetarFoco grdEntrada
      Exit Sub
    End If

    Set objUserEntradaAtendInc = New SisMaq.frmUserEntradaAtendInc
    'objUserEntradaAtendInc.Status = tpStatus_Incluir
    objUserEntradaAtendInc.Status = Status
    objUserEntradaAtendInc.lngPKID = 0
    objUserEntradaAtendInc.lngTURNOATENDEPESQ = lngTURNOATENDEPESQ
    objUserEntradaAtendInc.curSaldo = CCur(txtEmCaixa)
    objUserEntradaAtendInc.Show vbModal
    Set objUserEntradaAtendInc = Nothing
  Case 3
    'boleto Atend
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "Não há turno aberto. Por favor abra o turno antes de iniciar as atividades.", vbExclamation, TITULOSISTEMA
      SetarFoco grdEntrada
      Exit Sub
    End If

    Set objUserBoletoAtendInc = New SisMaq.frmUserBoletoAtendInc
    'objUserBoletoAtendInc.Status = tpStatus_Incluir
    objUserBoletoAtendInc.Status = Status
    objUserBoletoAtendInc.lngPKID = 0
    objUserBoletoAtendInc.lngTURNOATENDEPESQ = lngTURNOATENDEPESQ
    objUserBoletoAtendInc.Show vbModal
    Set objUserBoletoAtendInc = Nothing
  Case 4
    'frmUserBoletoDebInc
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "Não há turno aberto. Por favor abra o turno antes de iniciar as atividades.", vbExclamation, TITULOSISTEMA
      SetarFoco grdEntrada
      Exit Sub
    End If

    Set objUserBoletoArrecInc = New SisMaq.frmUserBoletoArrecInc
    objUserBoletoArrecInc.Status = tpStatus_Incluir
    objUserBoletoArrecInc.lngTURNOARRECEPESQ = 0
    objUserBoletoArrecInc.lngPKID = 0
    objUserBoletoArrecInc.Show vbModal
    Set objUserBoletoArrecInc = Nothing
  Case 5
    'Empréstimo/troca
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "Não há turno aberto. Por favor abra o turno antes de iniciar as atividades.", vbExclamation, TITULOSISTEMA
      SetarFoco grdEntrada
      Exit Sub
    End If

    Set objUserEmpTrocaInc = New SisMaq.frmUserEmpTrocaInc
    objUserEmpTrocaInc.Status = tpStatus_Incluir
    objUserEmpTrocaInc.lngPKID = 0
    objUserEmpTrocaInc.Show vbModal
    Set objUserEmpTrocaInc = Nothing
  Case 6
    'Fechamento Atendente
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "Não há turno aberto. Por favor abra o turno antes de iniciar as atividades.", vbExclamation, TITULOSISTEMA
      SetarFoco grdEntrada
      Exit Sub
    End If

    Set objUserFechaAteCons = New SisMaq.frmUserFechaAteCons
    objUserFechaAteCons.Status = Status
    objUserFechaAteCons.lngTURNOATENDEPESQ = lngTURNOATENDEPESQ
    objUserFechaAteCons.Show vbModal
    Set objUserFechaAteCons = Nothing
  Case 7
    'Fechamento Arrecadador
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "Não há turno aberto. Por favor abra o turno antes de iniciar as atividades.", vbExclamation, TITULOSISTEMA
      SetarFoco grdEntrada
      Exit Sub
    End If

    Set objUserFechaArrCons = New SisMaq.frmUserFechaArrCons
    objUserFechaArrCons.Status = tpStatus_Consultar
    objUserFechaArrCons.lngTURNOARRECEPESQ = 0
    objUserFechaArrCons.Show vbModal
    Set objUserFechaArrCons = Nothing
  Case 8
    'Fechamento do caixa
    FechamentoTurno
  Case 9
    'Atualizar Tela
  Case 10
    'Detalhamento Atendente
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "Não há turno aberto. Por favor abra o turno antes de iniciar as atividades.", vbExclamation, TITULOSISTEMA
      SetarFoco grdEntrada
      Exit Sub
    End If

    Set objUserDetalheAtend = New SisMaq.frmUserDetalhamentoAtend
    objUserDetalheAtend.Show vbModal
    Set objUserDetalheAtend = Nothing
  Case 11
    'Detalhamento Arrecadador
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "Não há turno aberto. Por favor abra o turno antes de iniciar as atividades.", vbExclamation, TITULOSISTEMA
      SetarFoco grdEntrada
      Exit Sub
    End If

    Set objUserDetalheArrec = New SisMaq.frmUserDetalhamentoArrec
    objUserDetalheArrec.Show vbModal
    Set objUserDetalheArrec = Nothing
  Case 12
    'Leitura Especial
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "Não há turno aberto. Por favor abra o turno antes de iniciar as atividades.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
    '
    Set objUserLeituraFechaInc = New SisMaq.frmUserLeituraFechaInc
    objUserLeituraFechaInc.Status = tpStatus_Incluir
    objUserLeituraFechaInc.lngLEITURAFECHAID = 0
    objUserLeituraFechaInc.Show vbModal
    Set objUserLeituraFechaInc = Nothing
    
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
    TratarErroPrevisto "Não há turnos em aberto, favor abrir um turno antes de iniciair o dia", "Form_Load"
  ElseIf RetornaCodTurnoCorrente(datDataTurno) = -1 Then
    TratarErroPrevisto "Há mais de 1 turno turnos em aberto, favor abrir um turno antes de iniciair o dia", "Form_Load"
  Else
    'OK Para turno
'''    datDataIniAtual = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now) & " 00:00:00")
'''    datDataFimAtual = DateAdd("d", 1, datDataIniAtual)
'''    If datDataTurno < datDataIniAtual Or datDataTurno >= datDataFimAtual Then
'''      TratarErroPrevisto "ATENÇÃO" & vbCrLf & vbCrLf & "A data do turno atual aberto não corresponde a data de hoje:" & vbCrLf & vbCrLf & "Data do turno --> " & Format(datDataTurno, "DD/MM/YYYY") & vbCrLf & "Data Atual --> " & Format(datDataIniAtual, "DD/MM/YYYY") & vbCrLf & vbCrLf & "Por favor, feche o turno e abra-o novamente.", "Form_Load"
'''    End If
  End If

  blnFechar = False
  blnRetorno = False
  blnPrimeiraVez = True
  AmpS
'''  If Me.ActiveControl Is Nothing Then
'''    Me.Top = 580
'''    Me.Left = 1
'''    Me.WindowState = 2 'Maximizado
'''  End If
  Me.Height = 9195
  Me.Width = 12090
  CenterForm Me
  '
  LerFigurasAvulsas cmdSairSelecao, "Sair.ico", "SairDown.ico", "Sair"
  '
  txtTurno.Text = RetornaDescTurnoCorrente
  txtUsuario.Text = gsNomeUsu
  mskDataPrinc.Text = Format(Date, "DD/MM/YYYY")
  'Trata nível de acesso
  If gsNivel = gsAdmin Or gsNivel = gsGerente Then
    cmdSelecao(8).Enabled = True
    cmdSelecao(0).Enabled = True
    cmdSelecao(1).Enabled = True
    cmdSelecao(12).Enabled = True
  Else
    cmdSelecao(8).Enabled = False
    cmdSelecao(0).Enabled = False
    cmdSelecao(1).Enabled = False
    cmdSelecao(12).Enabled = False
  End If
  'NOVO BOTÕES NOVOS
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
End Sub

Private Sub grdDespesa_UnboundReadDataEx( _
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
               Offset + intI, DESP_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, DESP_COLUNASMATRIZ, DESP_LINHASMATRIZ, DESP_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, DESP_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserOperCaiCons.gdbSaidaAtend_UnboundReadDataEx]"
End Sub

Private Sub grdEntArrec_UnboundReadDataEx( _
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
               Offset + intI, ENTARR_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, ENTARR_COLUNASMATRIZ, ENTARR_LINHASMATRIZ, ENTARR_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ENTARR_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserOperCaiCons.grdEntArrec_UnboundReadDataEx]"
End Sub

Private Sub grdEntrada_UnboundReadDataEx( _
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
               Offset + intI, ENTR_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, ENTR_COLUNASMATRIZ, ENTR_LINHASMATRIZ, ENTR_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ENTR_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserOperCaiCons.grdEntrada_UnboundReadDataEx]"
End Sub


Public Sub Trata_Matrizes_Totais()
  On Error GoTo trata
  Dim strSql          As String
  Dim objRs           As ADODB.Recordset
  Dim objGeral        As busSisMaq.clsGeral
  Dim curEntradaCaixa As Currency
  Dim curEntradaArrec As Currency
  Dim curSaldoAtend   As Currency
  Dim curSaldoAtendAtua   As Currency
  Dim curEntradaTotal As Currency
  Dim curLP           As Currency
  Dim curDespesaTotal As Currency
  '
  Dim curSaiAtend       As Currency
  Dim curDespesa  As Currency
  Dim curSaiSaldo       As Currency
  '
  Dim curSangria        As Currency
  Dim curEmpTroca       As Currency
  Dim curEmCaixa        As Currency
  
  'MONTA MATRIZES
  'Entrada
  ENTR_COLUNASMATRIZ = grdEntrada.Columns.Count
  ENTR_LINHASMATRIZ = 0
  MontaENTR_Matriz
  grdEntrada.Bookmark = Null
  grdEntrada.ReBind
  grdEntrada.ApproxCount = ENTR_LINHASMATRIZ
  blnPrimeiraVez = False
  'Entrada Arrecadador
  ENTARR_COLUNASMATRIZ = grdEntArrec.Columns.Count
  ENTARR_LINHASMATRIZ = 0
  MontaENTARR_Matriz
  grdEntArrec.Bookmark = Null
  grdEntArrec.ReBind
  grdEntArrec.ApproxCount = ENTARR_LINHASMATRIZ
  blnPrimeiraVez = False
  'Entrada Atendente
  ENTRAT_COLUNASMATRIZ = grdEntradaAtend.Columns.Count
  ENTRAT_LINHASMATRIZ = 0
  MontaENTRAT_Matriz
  grdEntradaAtend.Bookmark = Null
  grdEntradaAtend.ReBind
  grdEntradaAtend.ApproxCount = ENTRAT_LINHASMATRIZ
  blnPrimeiraVez = False
  'Despesas
  DESP_COLUNASMATRIZ = grdDespesa.Columns.Count
  DESP_LINHASMATRIZ = 0
  MontaDESP_Matriz
  grdDespesa.Bookmark = Null
  grdDespesa.ReBind
  grdDespesa.ApproxCount = DESP_LINHASMATRIZ
  blnPrimeiraVez = False
  '
  'Retirada
  RET_COLUNASMATRIZ = grdRetirada.Columns.Count
  RET_LINHASMATRIZ = 0
  MontaRET_Matriz
  grdRetirada.Bookmark = Null
  grdRetirada.ReBind
  grdRetirada.ApproxCount = RET_LINHASMATRIZ
  blnPrimeiraVez = False
  'Monta saldo
  curEntradaCaixa = 0
  curEntradaArrec = 0
  curSaldoAtend = 0
  curSaldoAtendAtua = 0
  curEntradaTotal = 0
  curLP = 0
  '
  curSaiAtend = 0
  curDespesa = 0
  curDespesaTotal = 0
  curSaiSaldo = 0
  '
  curSangria = 0
  curEmpTroca = 0
  curEmCaixa = 0
  '
  Set objGeral = New busSisMaq.clsGeral
  strSql = "SELECT ISNULL(SUM(ENTRADA.VALOR),0) AS TOTAL " & _
            "FROM ENTRADA " & _
            " WHERE ENTRADA.TURNOID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    curEntradaCaixa = objRs.Fields("TOTAL").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  strSql = "SELECT ISNULL(SUM(DEBITO.VALORPAGO),0) AS TOTAL "
  strSql = strSql & " FROM DEBITO " & _
          " INNER JOIN BOLETOATEND ON BOLETOATEND.PKID = DEBITO.BOLETOATENDID " & _
          " INNER JOIN CAIXAATEND ON CAIXAATEND.PKID = BOLETOATEND.CAIXAATENDID " & _
          " WHERE CAIXAATEND.TURNOENTRADAID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo)

  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    curSaldoAtendAtua = objRs.Fields("TOTAL").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  strSql = "SELECT ISNULL(SUM(CREDITO.VALORPAGO),0) AS TOTAL "
  strSql = strSql & " FROM CREDITO " & _
          " INNER JOIN BOLETOARREC ON BOLETOARREC.PKID = CREDITO.BOLETOARRECID " & _
          " INNER JOIN CAIXAARREC ON CAIXAARREC.PKID = BOLETOARREC.CAIXAARRECID " & _
          " WHERE CAIXAARREC.TURNOENTRADAID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo)

  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    curEntradaArrec = objRs.Fields("TOTAL").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  strSql = "SELECT ISNULL(SUM(ENTRADAATEND.VALOR),0) AS TOTAL " & _
            "FROM ENTRADAATEND " & _
            " INNER JOIN CAIXAATEND ON CAIXAATEND.PKID = ENTRADAATEND.CAIXAATENDID " & _
            " WHERE CAIXAATEND.TURNOENTRADAID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    curSaiAtend = objRs.Fields("TOTAL").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  strSql = "SELECT ISNULL(SUM(CAIXAATEND.VALORDEVOL),0) AS TOTAL "
  strSql = strSql & " FROM CAIXAATEND " & _
          " WHERE CAIXAATEND.TURNOFECHAID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    curSaldoAtend = objRs.Fields("TOTAL").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  strSql = "SELECT ISNULL(SUM(ENTRADAATEND.VALOR),0) AS TOTAL " & _
            "FROM ENTRADAATEND " & _
            " INNER JOIN CAIXAATEND ON CAIXAATEND.PKID = ENTRADAATEND.CAIXAATENDID " & _
            " WHERE CAIXAATEND.TURNOENTRADAID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    curSaiAtend = objRs.Fields("TOTAL").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  strSql = "SELECT ISNULL(SUM(DESPESA.VR_PAGAR), 0) AS TOTAL "
  strSql = strSql & " FROM DESPESA " & _
          " WHERE DESPESA.TURNOID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    curDespesa = objRs.Fields("TOTAL").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  strSql = "SELECT ISNULL(SUM(RETIRADA.VALOR),0) AS TOTAL " & _
            "FROM RETIRADA " & _
            " WHERE RETIRADA.TURNOID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    curSangria = objRs.Fields("TOTAL").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  strSql = "SELECT ISNULL(SUM(EMPTROCA.VALOR),0) AS TOTAL " & _
            "FROM EMPTROCA " & _
            " INNER JOIN TIPOPGTO ON TIPOPGTO.PKID = EMPTROCA.TIPOPGTOID " & _
            " WHERE EMPTROCA.TURNOID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo) '& _
            " AND TIPOPGTO.TIPOPGTO = " & Formata_Dados(gsDescEmprestimo, tpDados_Texto)
            
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    curEmpTroca = objRs.Fields("TOTAL").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  
  Set objGeral = Nothing
  '
  'Clacula total entrada
  curEntradaTotal = curEntradaCaixa + curEntradaArrec + curSaldoAtend
  curLP = curEntradaArrec - curSaldoAtendAtua
  curDespesaTotal = curSaiAtend + curDespesa
  curSaiSaldo = curEntradaTotal - curDespesaTotal
  curEmCaixa = curSaiSaldo - curSangria - curEmpTroca
  'Atualiza caixas de texto
  txtEntradaCaixa = Format(curEntradaCaixa, "###,##0.00")
  txtEntradaArrec = Format(curEntradaArrec, "###,##0.00")
  txtSaldoAtend = Format(curSaldoAtend, "###,##0.00")
  txtEntradaTotal = Format(curEntradaTotal, "###,##0.00")
  txtLP.Text = Format(curLP, "###,##0.00")
  If curLP >= 0 Then
    txtLP.BackColor = &H80000005
  Else
    txtLP.BackColor = &H8080FF
  End If
  '
  txtSaiAtendente = Format(curSaiAtend, "###,##0.00")
  txtDespesa = Format(curDespesa, "###,##0.00")
  txtDespesaTotal = Format(curDespesaTotal, "###,##0.00")
  txtSaiSaldo = Format(curSaiSaldo, "###,##0.00")
  '
  txtSangria = Format(curSangria, "###,##0.00")
  txtEmpTroca = Format(curEmpTroca, "###,##0.00")
  txtEmCaixa = Format(curEmCaixa, "###,##0.00")
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Activate()
  If blnPrimeiraVez Then
    DoEvents
    '
    Trata_Matrizes_Totais
    SetarFoco grdEntrada
  End If
End Sub

Public Sub MontaRET_Matriz()

  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMaq.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMaq.clsGeral
  '
  strSql = "SELECT RETIRADA.PKID, RETIRADA.DATA, RETIRADA.VALOR " & _
            "FROM RETIRADA " & _
            " WHERE RETIRADA.TURNOID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo) & _
            "ORDER BY RETIRADA.DATA"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    RET_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim RET_Matriz(0 To RET_COLUNASMATRIZ - 1, 0 To RET_LINHASMATRIZ - 1)
  Else
    ReDim RET_Matriz(0 To RET_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To RET_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To RET_COLUNASMATRIZ - 1  'varre as colunas
          RET_Matriz(intJ, intI) = objRs(intJ) & ""
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

Public Sub MontaENTRAT_Matriz()

  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMaq.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMaq.clsGeral
  '
  strSql = "SELECT ENTRADAATEND.PKID, PESSOA.NOME, ENTRADAATEND.DATA, ENTRADAATEND.VALOR " & _
            "FROM ENTRADAATEND " & _
            " INNER JOIN CAIXAATEND ON CAIXAATEND.PKID = ENTRADAATEND.CAIXAATENDID " & _
            " INNER JOIN PESSOA ON PESSOA.PKID = CAIXAATEND.ATENDENTEID " & _
            " WHERE CAIXAATEND.TURNOENTRADAID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo) & _
            "ORDER BY PESSOA.NOME, ENTRADAATEND.DATA"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    ENTRAT_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim ENTRAT_Matriz(0 To ENTRAT_COLUNASMATRIZ - 1, 0 To ENTRAT_LINHASMATRIZ - 1)
  Else
    ReDim ENTRAT_Matriz(0 To ENTRAT_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To ENTRAT_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To ENTRAT_COLUNASMATRIZ - 1  'varre as colunas
          ENTRAT_Matriz(intJ, intI) = objRs(intJ) & ""
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


Public Sub MontaDESP_Matriz()

  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMaq.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMaq.clsGeral
  '
  strSql = "SELECT DESPESA.PKID, GRUPODESPESA.DESCRICAO + ' - ' + SUBGRUPODESPESA.DESCRICAO, DESPESA.DESCRICAO, DESPESA.VR_PAGAR "
  strSql = strSql & " FROM DESPESA " & _
          " LEFT JOIN SUBGRUPODESPESA ON SUBGRUPODESPESA.PKID = DESPESA.SUBGRUPODESPESAID " & _
          " LEFT JOIN GRUPODESPESA ON GRUPODESPESA.PKID = SUBGRUPODESPESA.GRUPODESPESAID " & _
          " WHERE DESPESA.TURNOID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo) & _
          " ORDER BY GRUPODESPESA.DESCRICAO, SUBGRUPODESPESA.DESCRICAO;"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    DESP_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim DESP_Matriz(0 To DESP_COLUNASMATRIZ - 1, 0 To DESP_LINHASMATRIZ - 1)
  Else
    ReDim DESP_Matriz(0 To DESP_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To DESP_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To DESP_COLUNASMATRIZ - 1  'varre as colunas
          DESP_Matriz(intJ, intI) = objRs(intJ) & ""
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


Public Sub MontaENTARR_Matriz()

  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMaq.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMaq.clsGeral
  '
  strSql = "SELECT CREDITO.PKID, PESSOA.NOME, EQUIPAMENTO.NUMERO, CREDITO.MEDICAO, CREDITO.VALORPAGO "
  strSql = strSql & " FROM CREDITO " & _
          " INNER JOIN BOLETOARREC ON BOLETOARREC.PKID = CREDITO.BOLETOARRECID " & _
          " INNER JOIN MAQUINA ON MAQUINA.PKID = CREDITO.MAQUINAID " & _
          " INNER JOIN EQUIPAMENTO ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
          " INNER JOIN CAIXAARREC ON CAIXAARREC.PKID = BOLETOARREC.CAIXAARRECID " & _
          " INNER JOIN PESSOA ON PESSOA.PKID = CAIXAARREC.ARRECADADORID " & _
          " WHERE CAIXAARREC.TURNOENTRADAID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo) & _
          " ORDER BY PESSOA.NOME, EQUIPAMENTO.NUMERO;"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    ENTARR_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim ENTARR_Matriz(0 To ENTARR_COLUNASMATRIZ - 1, 0 To ENTARR_LINHASMATRIZ - 1)
  Else
    ReDim ENTARR_Matriz(0 To ENTARR_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To ENTARR_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To ENTARR_COLUNASMATRIZ - 1  'varre as colunas
          ENTARR_Matriz(intJ, intI) = objRs(intJ) & ""
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

Public Sub MontaENTR_Matriz()

  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMaq.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMaq.clsGeral
  '
  strSql = "SELECT ENTRADA.PKID, ENTRADA.DATA, ENTRADA.VALOR " & _
            "FROM ENTRADA " & _
            " WHERE ENTRADA.TURNOID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo) & _
            "ORDER BY ENTRADA.DATA"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    ENTR_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim ENTR_Matriz(0 To ENTR_COLUNASMATRIZ - 1, 0 To ENTR_LINHASMATRIZ - 1)
  Else
    ReDim ENTR_Matriz(0 To ENTR_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To ENTR_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To ENTR_COLUNASMATRIZ - 1  'varre as colunas
          ENTR_Matriz(intJ, intI) = objRs(intJ) & ""
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

Private Sub grdEntradaAtend_UnboundReadDataEx( _
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
               Offset + intI, ENTRAT_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, ENTRAT_COLUNASMATRIZ, ENTRAT_LINHASMATRIZ, ENTRAT_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ENTRAT_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserOperCaiCons.grdEntradaAtend_UnboundReadDataEx]"
End Sub

Private Sub grdRetirada_UnboundReadDataEx( _
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
               Offset + intI, RET_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, RET_COLUNASMATRIZ, RET_LINHASMATRIZ, RET_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, RET_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserOperCaiCons.grdRetirada_UnboundReadDataEx]"
End Sub

