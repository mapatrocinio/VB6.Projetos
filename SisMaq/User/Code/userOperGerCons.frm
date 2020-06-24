VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserOperGerCons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operacional Gerente"
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
      Caption         =   "Operacional Gerente"
      Height          =   6015
      Left            =   60
      TabIndex        =   20
      Top             =   360
      Width           =   11835
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   9210
         TabIndex        =   42
         Top             =   420
         Width           =   1275
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   6180
         TabIndex        =   40
         Top             =   450
         Width           =   1275
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3570
         TabIndex        =   38
         Top             =   480
         Width           =   1275
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3570
         TabIndex        =   36
         Top             =   210
         Width           =   1275
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   990
         TabIndex        =   34
         Top             =   480
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   990
         TabIndex        =   32
         Top             =   210
         Width           =   1275
      End
      Begin TrueDBGrid60.TDBGrid grdGeral 
         Height          =   1740
         Left            =   90
         OleObjectBlob   =   "userOperGerCons.frx":0000
         TabIndex        =   0
         Top             =   4290
         Width           =   5820
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid1 
         Height          =   1740
         Left            =   5910
         OleObjectBlob   =   "userOperGerCons.frx":4F06
         TabIndex        =   29
         Top             =   4290
         Width           =   5820
      End
      Begin TrueDBGrid60.TDBGrid grdEntrada 
         Height          =   1740
         Left            =   90
         OleObjectBlob   =   "userOperGerCons.frx":9E0C
         TabIndex        =   30
         Top             =   810
         Width           =   5820
      End
      Begin TrueDBGrid60.TDBGrid grdRetirada 
         Height          =   1740
         Left            =   5910
         OleObjectBlob   =   "userOperGerCons.frx":E3E4
         TabIndex        =   43
         Top             =   810
         Width           =   5820
      End
      Begin TrueDBGrid60.TDBGrid grdEntradaAtend 
         Height          =   1740
         Left            =   90
         OleObjectBlob   =   "userOperGerCons.frx":129C1
         TabIndex        =   44
         Top             =   2550
         Width           =   5820
      End
      Begin VB.Label Label8 
         Caption         =   "Saldo"
         Height          =   165
         Left            =   8340
         TabIndex        =   41
         Top             =   420
         Width           =   765
      End
      Begin VB.Label Label9 
         Caption         =   "Crédito"
         Height          =   165
         Left            =   5310
         TabIndex        =   39
         Top             =   450
         Width           =   765
      End
      Begin VB.Label Label6 
         Caption         =   "Débito"
         Height          =   165
         Left            =   2700
         TabIndex        =   37
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label5 
         Caption         =   "Entrada"
         Height          =   165
         Left            =   2700
         TabIndex        =   35
         Top             =   210
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Retirada"
         Height          =   165
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "Entrada"
         Height          =   165
         Left            =   120
         TabIndex        =   31
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
      TabIndex        =   10
      Top             =   6540
      Width           =   900
   End
   Begin VB.Frame fraImpressao 
      Caption         =   "Impressão"
      Height          =   2085
      Left            =   7770
      TabIndex        =   18
      Top             =   6420
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
         TabIndex        =   28
         Top             =   1680
         Width           =   2145
      End
      Begin VB.Label Label72 
         Caption         =   "CTRL + G - Detalhar Boleto Arrec."
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
         TabIndex        =   27
         Top             =   1470
         Width           =   2145
      End
      Begin VB.Label Label72 
         Caption         =   "CTRL + F - Detalhar Boleto Atend."
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
         TabIndex        =   26
         Top             =   1260
         Width           =   2145
      End
      Begin VB.Label Label72 
         Caption         =   "CTRL + E - Detalhar Entrada Atend."
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
         TabIndex        =   25
         Top             =   1050
         Width           =   2205
      End
      Begin VB.Label Label72 
         Caption         =   "CTRL + D - Detalhar Retirada"
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   420
         Width           =   2175
      End
      Begin VB.Label Label72 
         Caption         =   "CTRL + C - Detalhar Entrada"
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
         TabIndex        =   22
         Top             =   630
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
         TabIndex        =   19
         Top             =   210
         Width           =   2205
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Selecione a opção"
      Height          =   2085
      Left            =   60
      TabIndex        =   17
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
         Caption         =   "&G - Impressão GR  "
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
         Left            =   1350
         TabIndex        =   7
         ToolTipText     =   "impressão da GR"
         Top             =   630
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
         TabIndex        =   1
         ToolTipText     =   "Incluir GR"
         Top             =   240
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
         Left            =   1350
         TabIndex        =   2
         ToolTipText     =   "Alterar GR"
         Top             =   240
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
         Left            =   2640
         TabIndex        =   3
         ToolTipText     =   "ítens da GR"
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
         Left            =   3930
         TabIndex        =   4
         ToolTipText     =   "Recebimento da GR"
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
         Left            =   5250
         TabIndex        =   5
         ToolTipText     =   "Cancelar GR"
         Top             =   240
         Width           =   1305
      End
      Begin ComctlLib.StatusBar StatusBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   21
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
               TextSave        =   "15/8/2010"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   5
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   1235
               MinWidth        =   1235
               TextSave        =   "20:45"
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
      TabIndex        =   13
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
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "txtUsuario"
      Top             =   30
      Width           =   1695
   End
   Begin MSMask.MaskEdBox mskDataPrinc 
      Height          =   255
      Left            =   3990
      TabIndex        =   12
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
      TabIndex        =   16
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label21 
      Caption         =   "Turno Corrente"
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label Label22 
      Caption         =   "Usuário Logado"
      Height          =   255
      Left            =   30
      TabIndex        =   14
      Top             =   30
      Width           =   1215
   End
End
Attribute VB_Name = "frmUserOperGerCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public nGrupo                         As Integer
'''Public Status                   As tpStatus
Public blnRetorno                     As Boolean
Public blnFechar                      As Boolean
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


Public Sub Clique_botao(intIndice As Integer)
  If cmdSelecao(intIndice).Enabled = True And cmdSelecao(intIndice).Visible = True Then
    cmdSelecao_Click intIndice
  End If
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  'Tratamento de tecla para verificação de chamada de Outras telas
  'verifica se tem permissão
  'Tudo ok, faz chamada
  Select Case KeyAscii
  Case 1
    'TURNO - ABERTURA/REIMPRESSÃO
    frmUserTurnoInc.Show vbModal 'Turno
    Form_Load
  Case 2
    'TURNO - FECHAMENTO
    FechamentoTurno
    Form_Load
  Case 3
    'DETALHAR ENTRADA
    frmUserEntradaLis.Show vbModal
    Form_Load
  Case 4
    'DETALHAR RETIRADA
    frmUserRetiradaLis.Show vbModal
    Form_Load
  Case 5
    'DETALHAR ENTRADA ATENDENTE
    frmUserEntradaAtendLis.Show vbModal
    Form_Load
  Case 6
    'DETALHAR BOLETO ATENDENTE
    frmUserBoletoAtendLis.Show vbModal
    Form_Load
  Case 7
    'DETALHAR BOLETO ARRECADADOR
    frmUserBoletoArrecLis.Show vbModal
    Form_Load
'''  Case 4
'''    'ATUALIZAR
'''    Form_Load
'''  Case 5
'''    'CONSULTAR PRONTUÁRIO
'''    frmUserProntuarioGRCons.Show vbModal
'''    Form_Load
'''  Case 6
'''    'ZERAR SENHA
'''    frmUserZerarSenhaLis.Show vbModal
'''    Form_Load
'''  Case 7
'''    'CONSULTAR PROCEDIMENTO
'''    frmUserProcedimentoCons.indOrigem = 1
'''    frmUserProcedimentoCons.lngPRESTADORID = 0
'''    frmUserProcedimentoCons.Show vbModal
'''    Form_Load
'''  Case 8
'''    'CONSULTAR GR
'''    frmUserGRFinancCons.Show vbModal
'''    Form_Load
  End Select
  '
  Trata_Matrizes_Totais
  SetarFoco grdEntrada
  '
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserOperGerCons.Form_KeyPress]"
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
'  Dim objGRTotalPrestCons As SisMaq.frmUserGRTotalPrestCons
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
    objUserEntradaAtendInc.Status = tpStatus_Incluir
    objUserEntradaAtendInc.lngPKID = 0
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
    objUserBoletoAtendInc.Status = tpStatus_Incluir
    objUserBoletoAtendInc.lngPKID = 0
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
    objUserBoletoArrecInc.lngPKID = 0
    objUserBoletoArrecInc.Show vbModal
    Set objUserBoletoArrecInc = Nothing
    
    
    
    
'''  Case 1
'''    'Alteração da GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "Não há turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdEntrada.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para alterá-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Trim(RetornaDescAtende(grdEntrada.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
'''      MsgBox "Apenas o atendente que lançou a GR pode alterá-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Trim(grdEntrada.Columns("Status").Value & "") = "C" Then
'''      MsgBox "Não pode haver alteração em uma GR cancelada.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Trim(grdEntrada.Columns("Status").Value & "") = "F" Then
'''      'Pedir senha superior para alterar uma GR já fechada
'''      '----------------------------
'''      '----------------------------
'''      'Pede Senha Superior (Diretor, Gerente ou Administrador
'''      If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
'''        'Só pede senha superior se quem estiver logado não for superior
'''        gsNomeUsuLib = ""
'''        gsNivelUsuLib = ""
'''        frmUserLoginSup.Show vbModal
'''
'''        If Len(Trim(gsNomeUsuLib)) = 0 Then
'''          strMsg = "É necessário a confirmação com senha superior para alterar uma GR."
'''          TratarErroPrevisto strMsg, "cmdConfirmar_Click"
'''          SetarFoco grdEntrada
'''          Exit Sub
'''        Else
'''          'Capturou Nome do Usuário, continua com processo
'''        End If
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''    End If
'''
'''    Set objUserGRInc = New SisMaq.frmUserGRInc
'''    objUserGRInc.Status = tpStatus_Alterar
'''    objUserGRInc.IcEstadoGR = tpIcEstadoGR_Inic
'''    objUserGRInc.lngGRID = grdEntrada.Columns("ID").Value
'''    objUserGRInc.Show vbModal
'''    Set objUserGRInc = Nothing
'''  Case 2
'''    'Itens da GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "Não há turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdEntrada.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para alterar seus ítens.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Trim(RetornaDescAtende(grdEntrada.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
'''      MsgBox "Apenas o atendente que lançou a GR pode alterá-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Trim(grdEntrada.Columns("Status").Value & "") = "C" Then
'''      MsgBox "Não pode haver alteração de ítens de uma GR cancelada.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Trim(grdEntrada.Columns("Status").Value & "") = "F" Then
'''      'Pedir senha superior para alterar uma GR já fechada
'''      '----------------------------
'''      '----------------------------
'''      'Pede Senha Superior (Diretor, Gerente ou Administrador
'''      If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
'''        'Só pede senha superior se quem estiver logado não for superior
'''        gsNomeUsuLib = ""
'''        gsNivelUsuLib = ""
'''        frmUserLoginSup.Show vbModal
'''
'''        If Len(Trim(gsNomeUsuLib)) = 0 Then
'''          strMsg = "É necessário a confirmação com senha superior para alterar uma GR."
'''          TratarErroPrevisto strMsg, "cmdConfirmar_Click"
'''          SetarFoco grdEntrada
'''          Exit Sub
'''        Else
'''          'Capturou Nome do Usuário, continua com processo
'''        End If
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''    End If
'''    Set objUserGRInc = New SisMaq.frmUserGRInc
'''    objUserGRInc.Status = tpStatus_Alterar
'''    objUserGRInc.IcEstadoGR = tpIcEstadoGR_Proc
'''    objUserGRInc.lngGRID = grdEntrada.Columns("ID").Value
'''    objUserGRInc.Show vbModal
'''    Set objUserGRInc = Nothing
'''  Case 3
'''    'Alteração de pagamento da GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "Não há turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdEntrada.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para alterar seus dados de pagamento.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Trim(RetornaDescAtende(grdEntrada.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
'''      If (gsNivel <> gsLaboratorio) And (Trim(RetornaNivelAtende(grdEntrada.Columns("Atendente").Value & "")) <> gsLaboratorio) Then
'''        MsgBox "Apenas o atendente que lançou a GR pode efetuar o seu pagamento ou uma GR lançada pelo Laboratório.", vbExclamation, TITULOSISTEMA
'''        SetarFoco grdEntrada
'''        Exit Sub
'''      End If
'''    End If
'''    If Trim(grdEntrada.Columns("Status").Value & "") = "C" Then
'''      MsgBox "Não pode haver pagamento de uma GR cancelada.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Trim(grdEntrada.Columns("Status").Value & "") = "F" Then
'''      'Pedir senha superior para alterar uma GR já fechada
'''      '----------------------------
'''      '----------------------------
'''      'Pede Senha Superior (Diretor, Gerente ou Administrador
'''      If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
'''        'Só pede senha superior se quem estiver logado não for superior
'''        gsNomeUsuLib = ""
'''        gsNivelUsuLib = ""
'''        frmUserLoginSup.Show vbModal
'''
'''        If Len(Trim(gsNomeUsuLib)) = 0 Then
'''          strMsg = "É necessário a confirmação com senha superior para alterar pagamento de uma GR fechada."
'''          TratarErroPrevisto strMsg, "cmdConfirmar_Click"
'''          SetarFoco grdEntrada
'''          Exit Sub
'''        Else
'''          'Capturou Nome do Usuário, continua com processo
'''        End If
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''    End If
'''    Set objUserContaCorrente = New frmUserContaCorrente
'''    objUserContaCorrente.lngCCID = 0
'''    objUserContaCorrente.lngGRID = grdEntrada.Columns("ID").Value
'''    objUserContaCorrente.intGrupo = 0
'''    objUserContaCorrente.strFuncionarioNome = gsNomeUsuCompleto
'''    objUserContaCorrente.Status = tpStatus_Incluir
'''    objUserContaCorrente.strStatusLanc = "RC"
'''    objUserContaCorrente.strNivelAcesso = Trim(RetornaNivelAtende(grdEntrada.Columns("Atendente").Value & ""))
'''    objUserContaCorrente.Show vbModal
'''    Set objUserContaCorrente = Nothing
'''  Case 4
'''    'Cancelamento da GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "Não há turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdEntrada.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para excluí-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Trim(RetornaDescAtende(grdEntrada.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
'''      If Mid(grdEntrada.Columns("Atendente").Value & "", 2, 3) <> gsLaboratorio Then
'''        MsgBox "Apenas o atendente que lançou a GR pode excluí-la.", vbExclamation, TITULOSISTEMA
'''        SetarFoco grdEntrada
'''        Exit Sub
'''      End If
'''    End If
'''    'If Trim(grdEntrada.Columns("Status").Value & "") <> "F" Then
'''    '  MsgBox "Apenas pode de excluida uma GR fechada.", vbExclamation, TITULOSISTEMA
'''    '  SetarFoco grdEntrada
'''    '  Exit Sub
'''    'End If
'''    'If Trim(grdEntrada.Columns("Status").Value & "") = "F" Then
'''      'Pedir senha superior para alterar uma GR já fechada
'''      '----------------------------
'''      '----------------------------
'''      'Pede Senha Superior (Diretor, Gerente ou Administrador
'''      gsNomeUsuLib = ""
'''      gsNivelUsuLib = ""
'''      If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
'''        'Só pede senha superior se quem estiver logado não for superior
'''        frmUserLoginSup.Show vbModal
'''
'''        If Len(Trim(gsNomeUsuLib)) = 0 Then
'''          strMsg = "É necessário a confirmação com senha superior para cancelar uma GR."
'''          TratarErroPrevisto strMsg, "cmdConfirmar_Click"
'''          SetarFoco grdEntrada
'''          Exit Sub
'''        Else
'''          'Capturou Nome do Usuário, continua com processo
'''        End If
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''    'End If
'''    'Confirmação
'''    If MsgBox("Confirma cancelamento da GR " & grdEntrada.Columns("Seq.").Value & " de " & grdEntrada.Columns("Prontuário").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''
'''    Set objGR = New busSisMaq.clsGR
'''    objGR.AlterarStatusGR grdEntrada.Columns("ID").Value, _
'''                          "C", _
'''                          "", _
'''                          RetornaCodTurnoCorrente
'''    Set objGR = Nothing
'''    IMP_COMP_CANC_GR grdEntrada.Columns("ID").Value, gsNomeEmpresa, 1
'''  Case 5
'''    'consultar GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "Não há turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdEntrada.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para alterá-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Trim(RetornaDescAtende(grdEntrada.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
'''      MsgBox "Apenas o atendente que lançou a GR pode consultá-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    Set objUserGRInc = New SisMaq.frmUserGRInc
'''    objUserGRInc.Status = tpStatus_Consultar
'''    objUserGRInc.IcEstadoGR = tpIcEstadoGR_Con
'''    objUserGRInc.lngGRID = grdEntrada.Columns("ID").Value
'''    objUserGRInc.Show vbModal
'''    Set objUserGRInc = Nothing
'''  Case 6
'''    'Imprimir GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "Não há turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdEntrada.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para imprimí-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Trim(grdEntrada.Columns("Status").Value & "") <> "F" Then
'''      If Trim(RetornaNivelAtende(grdEntrada.Columns("Atendente").Value & "")) <> gsLaboratorio Then
'''        MsgBox "Não pode haver impressão de uma GR que não esteja fechada ou seja lançada pelo Laboratório.", vbExclamation, TITULOSISTEMA
'''        SetarFoco grdEntrada
'''        Exit Sub
'''      End If
'''    End If
'''    'Confirmação
'''    If MsgBox("Confirma impressão da GR " & grdEntrada.Columns("Seq.").Value & " de " & grdEntrada.Columns("Prontuário").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Trim(grdEntrada.Columns("Imp").Value & "") = "S" Then
'''      'Pedir senha superior para imprimir uma GR já impressa
'''      '----------------------------
'''      '----------------------------
'''      'Pede Senha Superior (Diretor, Gerente ou Administrador
'''      If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
'''        'Só pede senha superior se quem estiver logado não for superior
'''        gsNomeUsuLib = ""
'''        gsNivelUsuLib = ""
'''        frmUserLoginSup.Show vbModal
'''
'''        If Len(Trim(gsNomeUsuLib)) = 0 Then
'''          strMsg = "É necessário a confirmação com senha superior para imprimir uma GR já impressa."
'''          TratarErroPrevisto strMsg, "cmdConfirmar_Click"
'''          SetarFoco grdEntrada
'''          Exit Sub
'''        Else
'''          'Capturou Nome do Usuário, continua com processo
'''        End If
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''    End If
'''
'''    IMP_COMP_GR grdEntrada.Columns("ID").Value, gsNomeEmpresa, 1, IIf(Trim(grdEntrada.Columns("Imp").Value & "") = "S", True, False)
'''    'Após impressão altera status para impressa
'''    Set objGR = New busSisMaq.clsGR
'''    objGR.AlterarStatusGR grdEntrada.Columns("ID").Value, _
'''                          "", _
'''                          "S"
'''
'''
'''    Set objGR = Nothing
'''
'''  Case 7
'''    'Imprimir Comprovante de Recebimento
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "Não há turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdEntrada.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para imprimí-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    If Trim(grdEntrada.Columns("Status").Value & "") <> "F" Then
'''      MsgBox "Não pode haver impressão de uma GR que não esteja fechada.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''    'Confirmação
'''    If MsgBox("Confirma impressão da GR " & grdEntrada.Columns("Seq.").Value & " de " & grdEntrada.Columns("Prontuário").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco grdEntrada
'''      Exit Sub
'''    End If
'''
'''    'Fecou GR do laboratório, emitir comprovante de pagamento
'''    IMP_COMPROV_REC grdEntrada.Columns("ID").Value, gsNomeEmpresa, 1
'''  Case 8
'''    'Canlelar GR outros prestadores
'''    Set objGRTotalPrestCons = New SisMaq.frmUserGRTotalPrestCons
'''    objGRTotalPrestCons.Show vbModal
'''    Set objGRTotalPrestCons = Nothing
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

  'NOVO BOTÕES NOVOS
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
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
  TratarErro Err.Number, Err.Description, "[frmUserOperGerCons.grdEntrada_UnboundReadDataEx]"
End Sub


Public Sub Trata_Matrizes_Totais()
  On Error GoTo trata
  'Entrada
  ENTR_COLUNASMATRIZ = grdEntrada.Columns.Count
  ENTR_LINHASMATRIZ = 0
  MontaENTR_Matriz
  grdEntrada.Bookmark = Null
  grdEntrada.ReBind
  grdEntrada.ApproxCount = ENTR_LINHASMATRIZ
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
  'Entrada Atendente
  ENTRAT_COLUNASMATRIZ = grdEntradaAtend.Columns.Count
  ENTRAT_LINHASMATRIZ = 0
  MontaENTRAT_Matriz
  grdEntradaAtend.Bookmark = Null
  grdEntradaAtend.ReBind
  grdEntradaAtend.ApproxCount = ENTRAT_LINHASMATRIZ
  blnPrimeiraVez = False
  
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
  TratarErro Err.Number, Err.Description, "[frmUserOperGerCons.grdEntradaAtend_UnboundReadDataEx]"
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
  TratarErro Err.Number, Err.Description, "[frmUserOperGerCons.grdRetirada_UnboundReadDataEx]"
End Sub


