VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserOperAteCons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operacional Atendente"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraUnidade 
      Caption         =   "Operacional Atendente"
      Height          =   6015
      Left            =   60
      TabIndex        =   13
      Top             =   360
      Width           =   11835
      Begin VB.TextBox txtSenha 
         Height          =   312
         IMEMode         =   3  'DISABLE
         Left            =   1410
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   585
         Width           =   1452
      End
      Begin VB.TextBox txtUsuario 
         Height          =   312
         Left            =   1410
         MaxLength       =   30
         TabIndex        =   0
         Top             =   270
         Width           =   1452
      End
      Begin TrueDBGrid60.TDBGrid grdEntradaAtend 
         Height          =   1740
         Left            =   90
         OleObjectBlob   =   "userOperAteCons.frx":0000
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1170
         Width           =   5820
      End
      Begin TrueDBGrid60.TDBGrid grdAtendente 
         Height          =   2940
         Left            =   90
         OleObjectBlob   =   "userOperAteCons.frx":4BB1
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2910
         Width           =   11640
      End
      Begin MSMask.MaskEdBox mskEntrada 
         Height          =   255
         Left            =   4320
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   270
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPago 
         Height          =   255
         Left            =   4320
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   570
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskSaldo 
         Height          =   255
         Left            =   4320
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   870
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin TrueDBGrid60.TDBGrid grdBoleto 
         Height          =   1740
         Left            =   5910
         OleObjectBlob   =   "userOperAteCons.frx":E18F
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1170
         Width           =   5820
      End
      Begin VB.Label Label5 
         Caption         =   "Saldo"
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
         Index           =   4
         Left            =   3060
         TabIndex        =   25
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Pago"
         Height          =   285
         Index           =   3
         Left            =   3060
         TabIndex        =   23
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Entrada"
         Height          =   285
         Index           =   0
         Left            =   3060
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Usu�rio"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   18
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Senha"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   17
         Top             =   615
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSairSelecao 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   855
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6570
      Width           =   900
   End
   Begin VB.Frame fraImpressao 
      Caption         =   "Impress�o"
      Height          =   1035
      Left            =   7770
      TabIndex        =   11
      Top             =   6420
      Visible         =   0   'False
      Width           =   2355
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
         TabIndex        =   15
         Top             =   420
         Width           =   2175
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
         TabIndex        =   12
         Top             =   210
         Width           =   2205
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Selecione a op��o"
      Height          =   1065
      Left            =   60
      TabIndex        =   10
      Top             =   6420
      Width           =   7665
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&C - Deslogar             "
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
         TabIndex        =   4
         ToolTipText     =   "Deslogar"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&B - Alterar �tem        "
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
         TabIndex        =   3
         ToolTipText     =   "Deslogar"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&A - Boleto                  "
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
         TabIndex        =   2
         ToolTipText     =   "Boleto"
         Top             =   240
         Width           =   1305
      End
      Begin ComctlLib.StatusBar StatusBar1 
         Height          =   255
         Left            =   2670
         TabIndex        =   14
         Top             =   750
         Width           =   4320
         _ExtentX        =   7620
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
               TextSave        =   "20/11/2010"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   5
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   1235
               MinWidth        =   1235
               TextSave        =   "17:22"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   1
               Alignment       =   1
               Bevel           =   2
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
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "txtTurno"
      Top             =   30
      Width           =   4815
   End
   Begin MSMask.MaskEdBox mskDataPrinc 
      Height          =   255
      Left            =   900
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   60
      Width           =   1245
      _ExtentX        =   2196
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
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   765
   End
   Begin VB.Label Label21 
      Caption         =   "Turno Corrente"
      Height          =   255
      Left            =   2190
      TabIndex        =   8
      Top             =   60
      Width           =   1245
   End
End
Attribute VB_Name = "frmUserOperAteCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public nGrupo                         As Integer
Public Status                         As tpStatus
Public blnRetorno                     As Boolean
Public blnFechar                      As Boolean
Private lngFUNCIONARIOID              As Long
Private lngTURNOATENDID               As Long
Public lngTURNOATENDEPESQ             As Long
'''
'''Public objUserGRInc             As SisMaq.frmUserGRInc
'''Public objUserContaCorrente     As SisMaq.frmUserContaCorrente
'''
Public blnPrimeiraVez                 As Boolean 'Prop�sito: Preencher lista no combo

'Entrada
Private ATEND_COLUNASMATRIZ            As Long
Private ATEND_LINHASMATRIZ             As Long
Private ATEND_Matriz()                 As String
'Resumo Boleto
Private RESBOL_COLUNASMATRIZ            As Long
Private RESBOL_LINHASMATRIZ             As Long
Private RESBOL_Matriz()                 As String
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
  'Tratamento de tecla para verifica��o de chamada de Outras telas
  'verifica se tem permiss�o
  'Tudo ok, faz chamada
  If KeyAscii = 13 Then
    SendKeys "{tab}"
    Exit Sub
  End If
'''  Select Case KeyAscii
'''  Case 1
'''    'TURNO - ABERTURA/REIMPRESS�O
'''    frmUserTurnoInc.Show vbModal 'Turno
'''    Form_Load
'''  Case 2
'''    'TURNO - FECHAMENTO
'''    FechamentoTurno
'''    Form_Load
'''  Case 3
'''    'DETALHAR ENTRADA
'''    frmUserEntradaLis.Show vbModal
'''    Form_Load
'''  Case 4
'''    'DETALHAR RETIRADA
'''    frmUserRetiradaLis.Show vbModal
'''    Form_Load
'''  Case 5
'''    'DETALHAR ENTRADA ATENDENTE
'''    frmUserEntradaAtendLis.Show vbModal
'''    Form_Load
'''  Case 6
'''    'DETALHAR BOLETO ATENDENTE
'''    frmUserBoletoAtendLis.Show vbModal
'''    Form_Load
'''  Case 4
'''    'ATUALIZAR
'''    Form_Load
'''  Case 5
'''    'CONSULTAR PRONTU�RIO
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
'''  End Select
  '
  'Trata_Matrizes_Totais
  'SetarFoco txtUsuario
  '
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserOperAteCons.Form_KeyPress]"
End Sub

'''Private Sub cmdInfFinanc_Click()
'''  On Error GoTo trata
'''  'Chamar o form de Consulta/Visualiza��o das Informa��es Financeiras.
'''  frmUserInfFinancLis.Show vbModal
'''  SetarFoco grdAtendente
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
  SetarFoco txtUsuario
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[cmdSelecao_Click]"
  frmMDI.tmrUnidade.Enabled = True
End Sub


Public Sub VerificaQuemChamou()
  Dim objUserBoletoDebInc As SisMaq.frmUserBoletoDebInc
  Dim objDebitoInc As SisMaq.frmUserDebitoInc
  Dim strMsg As String
  On Error GoTo trata
  '
  Select Case nGrupo

  Case 0
    'Boleto
    If RetornaCodTurnoCorrente = 0 Then
      MsgBox "N�o h� turno aberto. Por favor abra o turno antes de iniciar as atividades.", vbExclamation, TITULOSISTEMA
      SetarFoco txtUsuario
      Exit Sub
    End If
    If Not Valida_String(txtUsuario, TpObrigatorio, True) Then
      MsgBox "Preencher o usu�rio.", vbExclamation, TITULOSISTEMA
      SetarFoco txtUsuario
      Exit Sub
    End If
    If Not Valida_String(txtSenha, TpObrigatorio, True) Then
      MsgBox "Preencher a senha.", vbExclamation, TITULOSISTEMA
      SetarFoco txtSenha
      Exit Sub
    End If
    '
    Set objUserBoletoDebInc = New SisMaq.frmUserBoletoDebInc
    objUserBoletoDebInc.Status = tpStatus_Incluir
    objUserBoletoDebInc.lngBOLETOATENDID = 0
    objUserBoletoDebInc.txtUsuario = txtUsuario.Text
    objUserBoletoDebInc.txtSenha = txtSenha.Text
    objUserBoletoDebInc.Show vbModal
    Set objUserBoletoDebInc = Nothing
  Case 1
    If Not IsNumeric(grdAtendente.Columns("DEBITOID").Value) Then
      MsgBox "Selecione um lan�amento !", vbExclamation, TITULOSISTEMA
      Exit Sub
    End If
    '
    Set objDebitoInc = New SisMaq.frmUserDebitoInc
    objDebitoInc.Status = tpStatus_Alterar
    objDebitoInc.lngPKID = grdAtendente.Columns("DEBITOID").Value
    objDebitoInc.Show vbModal
    Set objDebitoInc = Nothing
  Case 2
    'Deslogar
    lngFUNCIONARIOID = 0
    txtUsuario.Text = ""
    txtSenha.Text = ""
    Trata_Matrizes_Totais
    SetarFoco txtUsuario
    
'''  Case 1
'''    'Altera��o da GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdAtendente.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para alter�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Trim(RetornaDescAtende(grdAtendente.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
'''      MsgBox "Apenas o atendente que lan�ou a GR pode alter�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Trim(grdAtendente.Columns("Status").Value & "") = "C" Then
'''      MsgBox "N�o pode haver altera��o em uma GR cancelada.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Trim(grdAtendente.Columns("Status").Value & "") = "F" Then
'''      'Pedir senha superior para alterar uma GR j� fechada
'''      '----------------------------
'''      '----------------------------
'''      'Pede Senha Superior (Diretor, Gerente ou Administrador
'''      If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
'''        'S� pede senha superior se quem estiver logado n�o for superior
'''        gsNomeUsuLib = ""
'''        gsNivelUsuLib = ""
'''        frmUserLoginSup.Show vbModal
'''
'''        If Len(Trim(gsNomeUsuLib)) = 0 Then
'''          strMsg = "� necess�rio a confirma��o com senha superior para alterar uma GR."
'''          TratarErroPrevisto strMsg, "cmdConfirmar_Click"
'''          SetarFoco grdAtendente
'''          Exit Sub
'''        Else
'''          'Capturou Nome do Usu�rio, continua com processo
'''        End If
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''    End If
'''
'''    Set objUserGRInc = New SisMaq.frmUserGRInc
'''    objUserGRInc.Status = tpStatus_Alterar
'''    objUserGRInc.IcEstadoGR = tpIcEstadoGR_Inic
'''    objUserGRInc.lngGRID = grdAtendente.Columns("ID").Value
'''    objUserGRInc.Show vbModal
'''    Set objUserGRInc = Nothing
'''  Case 2
'''    'Itens da GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdAtendente.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para alterar seus �tens.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Trim(RetornaDescAtende(grdAtendente.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
'''      MsgBox "Apenas o atendente que lan�ou a GR pode alter�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Trim(grdAtendente.Columns("Status").Value & "") = "C" Then
'''      MsgBox "N�o pode haver altera��o de �tens de uma GR cancelada.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Trim(grdAtendente.Columns("Status").Value & "") = "F" Then
'''      'Pedir senha superior para alterar uma GR j� fechada
'''      '----------------------------
'''      '----------------------------
'''      'Pede Senha Superior (Diretor, Gerente ou Administrador
'''      If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
'''        'S� pede senha superior se quem estiver logado n�o for superior
'''        gsNomeUsuLib = ""
'''        gsNivelUsuLib = ""
'''        frmUserLoginSup.Show vbModal
'''
'''        If Len(Trim(gsNomeUsuLib)) = 0 Then
'''          strMsg = "� necess�rio a confirma��o com senha superior para alterar uma GR."
'''          TratarErroPrevisto strMsg, "cmdConfirmar_Click"
'''          SetarFoco grdAtendente
'''          Exit Sub
'''        Else
'''          'Capturou Nome do Usu�rio, continua com processo
'''        End If
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''    End If
'''    Set objUserGRInc = New SisMaq.frmUserGRInc
'''    objUserGRInc.Status = tpStatus_Alterar
'''    objUserGRInc.IcEstadoGR = tpIcEstadoGR_Proc
'''    objUserGRInc.lngGRID = grdAtendente.Columns("ID").Value
'''    objUserGRInc.Show vbModal
'''    Set objUserGRInc = Nothing
'''  Case 3
'''    'Altera��o de pagamento da GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdAtendente.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para alterar seus dados de pagamento.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Trim(RetornaDescAtende(grdAtendente.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
'''      If (gsNivel <> gsLaboratorio) And (Trim(RetornaNivelAtende(grdAtendente.Columns("Atendente").Value & "")) <> gsLaboratorio) Then
'''        MsgBox "Apenas o atendente que lan�ou a GR pode efetuar o seu pagamento ou uma GR lan�ada pelo Laborat�rio.", vbExclamation, TITULOSISTEMA
'''        SetarFoco grdAtendente
'''        Exit Sub
'''      End If
'''    End If
'''    If Trim(grdAtendente.Columns("Status").Value & "") = "C" Then
'''      MsgBox "N�o pode haver pagamento de uma GR cancelada.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Trim(grdAtendente.Columns("Status").Value & "") = "F" Then
'''      'Pedir senha superior para alterar uma GR j� fechada
'''      '----------------------------
'''      '----------------------------
'''      'Pede Senha Superior (Diretor, Gerente ou Administrador
'''      If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
'''        'S� pede senha superior se quem estiver logado n�o for superior
'''        gsNomeUsuLib = ""
'''        gsNivelUsuLib = ""
'''        frmUserLoginSup.Show vbModal
'''
'''        If Len(Trim(gsNomeUsuLib)) = 0 Then
'''          strMsg = "� necess�rio a confirma��o com senha superior para alterar pagamento de uma GR fechada."
'''          TratarErroPrevisto strMsg, "cmdConfirmar_Click"
'''          SetarFoco grdAtendente
'''          Exit Sub
'''        Else
'''          'Capturou Nome do Usu�rio, continua com processo
'''        End If
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''    End If
'''    Set objUserContaCorrente = New frmUserContaCorrente
'''    objUserContaCorrente.lngCCID = 0
'''    objUserContaCorrente.lngGRID = grdAtendente.Columns("ID").Value
'''    objUserContaCorrente.intGrupo = 0
'''    objUserContaCorrente.strFuncionarioNome = gsNomeUsuCompleto
'''    objUserContaCorrente.Status = tpStatus_Incluir
'''    objUserContaCorrente.strStatusLanc = "RC"
'''    objUserContaCorrente.strNivelAcesso = Trim(RetornaNivelAtende(grdAtendente.Columns("Atendente").Value & ""))
'''    objUserContaCorrente.Show vbModal
'''    Set objUserContaCorrente = Nothing
'''  Case 4
'''    'Cancelamento da GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdAtendente.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para exclu�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Trim(RetornaDescAtende(grdAtendente.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
'''      If Mid(grdAtendente.Columns("Atendente").Value & "", 2, 3) <> gsLaboratorio Then
'''        MsgBox "Apenas o atendente que lan�ou a GR pode exclu�-la.", vbExclamation, TITULOSISTEMA
'''        SetarFoco grdAtendente
'''        Exit Sub
'''      End If
'''    End If
'''    'If Trim(grdAtendente.Columns("Status").Value & "") <> "F" Then
'''    '  MsgBox "Apenas pode de excluida uma GR fechada.", vbExclamation, TITULOSISTEMA
'''    '  SetarFoco grdAtendente
'''    '  Exit Sub
'''    'End If
'''    'If Trim(grdAtendente.Columns("Status").Value & "") = "F" Then
'''      'Pedir senha superior para alterar uma GR j� fechada
'''      '----------------------------
'''      '----------------------------
'''      'Pede Senha Superior (Diretor, Gerente ou Administrador
'''      gsNomeUsuLib = ""
'''      gsNivelUsuLib = ""
'''      If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
'''        'S� pede senha superior se quem estiver logado n�o for superior
'''        frmUserLoginSup.Show vbModal
'''
'''        If Len(Trim(gsNomeUsuLib)) = 0 Then
'''          strMsg = "� necess�rio a confirma��o com senha superior para cancelar uma GR."
'''          TratarErroPrevisto strMsg, "cmdConfirmar_Click"
'''          SetarFoco grdAtendente
'''          Exit Sub
'''        Else
'''          'Capturou Nome do Usu�rio, continua com processo
'''        End If
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''    'End If
'''    'Confirma��o
'''    If MsgBox("Confirma cancelamento da GR " & grdAtendente.Columns("Seq.").Value & " de " & grdAtendente.Columns("Prontu�rio").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''
'''    Set objGR = New busSisMaq.clsGR
'''    objGR.AlterarStatusGR grdAtendente.Columns("ID").Value, _
'''                          "C", _
'''                          "", _
'''                          RetornaCodTurnoCorrente
'''    Set objGR = Nothing
'''    IMP_COMP_CANC_GR grdAtendente.Columns("ID").Value, gsNomeEmpresa, 1
'''  Case 5
'''    'consultar GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdAtendente.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para alter�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Trim(RetornaDescAtende(grdAtendente.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
'''      MsgBox "Apenas o atendente que lan�ou a GR pode consult�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    Set objUserGRInc = New SisMaq.frmUserGRInc
'''    objUserGRInc.Status = tpStatus_Consultar
'''    objUserGRInc.IcEstadoGR = tpIcEstadoGR_Con
'''    objUserGRInc.lngGRID = grdAtendente.Columns("ID").Value
'''    objUserGRInc.Show vbModal
'''    Set objUserGRInc = Nothing
'''  Case 6
'''    'Imprimir GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdAtendente.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para imprim�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Trim(grdAtendente.Columns("Status").Value & "") <> "F" Then
'''      If Trim(RetornaNivelAtende(grdAtendente.Columns("Atendente").Value & "")) <> gsLaboratorio Then
'''        MsgBox "N�o pode haver impress�o de uma GR que n�o esteja fechada ou seja lan�ada pelo Laborat�rio.", vbExclamation, TITULOSISTEMA
'''        SetarFoco grdAtendente
'''        Exit Sub
'''      End If
'''    End If
'''    'Confirma��o
'''    If MsgBox("Confirma impress�o da GR " & grdAtendente.Columns("Seq.").Value & " de " & grdAtendente.Columns("Prontu�rio").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Trim(grdAtendente.Columns("Imp").Value & "") = "S" Then
'''      'Pedir senha superior para imprimir uma GR j� impressa
'''      '----------------------------
'''      '----------------------------
'''      'Pede Senha Superior (Diretor, Gerente ou Administrador
'''      If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
'''        'S� pede senha superior se quem estiver logado n�o for superior
'''        gsNomeUsuLib = ""
'''        gsNivelUsuLib = ""
'''        frmUserLoginSup.Show vbModal
'''
'''        If Len(Trim(gsNomeUsuLib)) = 0 Then
'''          strMsg = "� necess�rio a confirma��o com senha superior para imprimir uma GR j� impressa."
'''          TratarErroPrevisto strMsg, "cmdConfirmar_Click"
'''          SetarFoco grdAtendente
'''          Exit Sub
'''        Else
'''          'Capturou Nome do Usu�rio, continua com processo
'''        End If
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''    End If
'''
'''    IMP_COMP_GR grdAtendente.Columns("ID").Value, gsNomeEmpresa, 1, IIf(Trim(grdAtendente.Columns("Imp").Value & "") = "S", True, False)
'''    'Ap�s impress�o altera status para impressa
'''    Set objGR = New busSisMaq.clsGR
'''    objGR.AlterarStatusGR grdAtendente.Columns("ID").Value, _
'''                          "", _
'''                          "S"
'''
'''
'''    Set objGR = Nothing
'''
'''  Case 7
'''    'Imprimir Comprovante de Recebimento
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdAtendente.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para imprim�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    If Trim(grdAtendente.Columns("Status").Value & "") <> "F" Then
'''      MsgBox "N�o pode haver impress�o de uma GR que n�o esteja fechada.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''    'Confirma��o
'''    If MsgBox("Confirma impress�o da GR " & grdAtendente.Columns("Seq.").Value & " de " & grdAtendente.Columns("Prontu�rio").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco grdAtendente
'''      Exit Sub
'''    End If
'''
'''    'Fecou GR do laborat�rio, emitir comprovante de pagamento
'''    IMP_COMPROV_REC grdAtendente.Columns("ID").Value, gsNomeEmpresa, 1
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
  Dim objGeral          As busSisMaq.clsGeral
  Dim objRs             As ADODB.Recordset
  Dim datDataTurno      As Date
  Dim datDataIniAtual   As Date
  Dim datDataFimAtual   As Date
  '
  If RetornaCodTurnoCorrente(datDataTurno) = 0 Then
    TratarErroPrevisto "N�o h� turnos em aberto, favor informar ao Gerente para abrir o turno.", "Form_Load"
    End
  Else
    'OK Para turno
'''    datDataIniAtual = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now) & " 00:00:00")
'''    datDataFimAtual = DateAdd("d", 1, datDataIniAtual)
'''    If datDataTurno < datDataIniAtual Or datDataTurno >= datDataFimAtual Then
'''      TratarErroPrevisto "ATEN��O" & vbCrLf & vbCrLf & "A data do turno atual aberto n�o corresponde a data de hoje:" & vbCrLf & vbCrLf & "Data do turno --> " & Format(datDataTurno, "DD/MM/YYYY") & vbCrLf & "Data Atual --> " & Format(datDataIniAtual, "DD/MM/YYYY") & vbCrLf & vbCrLf & "Por favor, feche o turno e abra-o novamente.", "Form_Load"
'''    End If
  End If

  blnFechar = False
  blnRetorno = False
  blnPrimeiraVez = True
  lngFUNCIONARIOID = 0
  AmpS
'''  If Me.ActiveControl Is Nothing Then
'''    Me.Top = 580
'''    Me.Left = 1
'''    Me.WindowState = 2 'Maximizado
'''  End If
  Me.Height = 8010
  Me.Width = 12090
  CenterForm Me
  '
  LerFigurasAvulsas cmdSairSelecao, "Sair.ico", "SairDown.ico", "Sair"
  '
  txtTurno.Text = RetornaDescTurnoCorrente
  mskDataPrinc.Text = Format(Date, "DD/MM/YYYY")
  If Status = tpStatus_Incluir Then
    Label5(1).Enabled = True
    Label5(2).Enabled = True
    txtUsuario.Enabled = True
    txtSenha.Enabled = True
    cmdSelecao(1).Enabled = False
    cmdSelecao(2).Enabled = True
  ElseIf Status = tpStatus_Consultar Then
    Label5(1).Enabled = False
    Label5(2).Enabled = False
    txtUsuario.Enabled = False
    txtSenha.Enabled = False
    cmdSelecao(1).Enabled = True
    cmdSelecao(2).Enabled = False
    'Obter usu�rio e senha do turno corrente
    Set objGeral = New busSisMaq.clsGeral
    strSql = "SELECT FUNCIONARIO.USUARIO, FUNCIONARIO.SENHA FROM FUNCIONARIO " & _
          " INNER JOIN CAIXAATEND ON FUNCIONARIO.PESSOAID = CAIXAATEND.ATENDENTEID " & _
          " WHERE CAIXAATEND.PKID = " & Formata_Dados(lngTURNOATENDEPESQ, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If objRs.EOF Then
      txtUsuario.Text = ""
      txtSenha.Text = ""
    Else
      txtUsuario.Text = objRs.Fields("USUARIO").Value & ""
      txtSenha.Text = Encripta(UCase$(objRs.Fields("SENHA").Value & ""))
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
  End If
  'NOVO BOT�ES NOVOS
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
End Sub

Private Sub grdAtendente_UnboundReadDataEx( _
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
               Offset + intI, ATEND_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, ATEND_COLUNASMATRIZ, ATEND_LINHASMATRIZ, ATEND_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ATEND_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserOperAteCons.grdAtendente_UnboundReadDataEx]"
End Sub


Public Sub Trata_Matrizes_Totais()
  On Error GoTo trata
  Dim strSql          As String
  Dim objRs           As ADODB.Recordset
  Dim objGeral        As busSisMaq.clsGeral
  Dim curSaldoEntrada As Currency
  Dim curSaldoPago    As Currency
  'Entrada Atendente
  ENTRAT_COLUNASMATRIZ = grdEntradaAtend.Columns.Count
  ENTRAT_LINHASMATRIZ = 0
  MontaENTRAT_Matriz
  grdEntradaAtend.Bookmark = Null
  grdEntradaAtend.ReBind
  grdEntradaAtend.ApproxCount = ENTRAT_LINHASMATRIZ
  blnPrimeiraVez = False
  'Boleto Atendente
  RESBOL_COLUNASMATRIZ = grdBoleto.Columns.Count
  RESBOL_LINHASMATRIZ = 0
  MontaRESBOL_Matriz
  grdBoleto.Bookmark = Null
  grdBoleto.ReBind
  grdBoleto.ApproxCount = RESBOL_LINHASMATRIZ
  blnPrimeiraVez = False
  'Atendente
  ATEND_COLUNASMATRIZ = grdAtendente.Columns.Count
  ATEND_LINHASMATRIZ = 0
  MontaATEND_Matriz
  grdAtendente.Bookmark = Null
  grdAtendente.ReBind
  grdAtendente.ApproxCount = ATEND_LINHASMATRIZ
  blnPrimeiraVez = False
  'Monta saldo
  curSaldoEntrada = 0
  curSaldoPago = 0
  '
  Set objGeral = New busSisMaq.clsGeral
  strSql = "SELECT ISNULL(SUM(ENTRADAATEND.VALOR), 0) AS VALOR " & _
            "FROM ENTRADAATEND " & _
            " INNER JOIN CAIXAATEND ON CAIXAATEND.PKID = ENTRADAATEND.CAIXAATENDID " & _
            " INNER JOIN PESSOA ON PESSOA.PKID = CAIXAATEND.ATENDENTEID " & _
            " WHERE CAIXAATEND.PKID = " & Formata_Dados(RetornaCodTurnoCorrenteAtend(lngFUNCIONARIOID, lngTURNOATENDEPESQ), tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    curSaldoEntrada = objRs.Fields("VALOR").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  strSql = "SELECT ISNULL(SUM(DEBITO.VALORPAGO), 0) AS VALOR "
  strSql = strSql & " FROM DEBITO " & _
          " INNER JOIN BOLETOATEND ON BOLETOATEND.PKID = DEBITO.BOLETOATENDID " & _
          " INNER JOIN MAQUINA ON MAQUINA.PKID = DEBITO.MAQUINAID " & _
          " INNER JOIN EQUIPAMENTO ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
          " INNER JOIN CAIXAATEND ON CAIXAATEND.PKID = BOLETOATEND.CAIXAATENDID " & _
          " INNER JOIN PESSOA ON PESSOA.PKID = CAIXAATEND.ATENDENTEID " & _
          " WHERE CAIXAATEND.PKID = " & Formata_Dados(RetornaCodTurnoCorrenteAtend(lngFUNCIONARIOID, lngTURNOATENDEPESQ), tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    curSaldoPago = objRs.Fields("VALOR").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set objGeral = Nothing
  '
  INCLUIR_VALOR_NO_MASK mskEntrada, curSaldoEntrada, TpMaskMoeda
  INCLUIR_VALOR_NO_MASK mskPago, curSaldoPago, TpMaskMoeda
  INCLUIR_VALOR_NO_MASK mskSaldo, curSaldoEntrada - curSaldoPago, TpMaskMoeda
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
    SetarFoco txtUsuario
  End If
End Sub

Public Sub MontaRESBOL_Matriz()

  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMaq.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMaq.clsGeral
  '
  strSql = "SELECT BOLETOATEND.PKID, MAX(PESSOA.NOME), BOLETOATEND.NUMERO, SUM(DEBITO.VALORPAGO) AS TOTAL, MAX(BOLETOATEND.STATUS), ISNULL(COUNT(DEBITO.PKID),0) AS LANCADO, 10 - ISNULL(COUNT(DEBITO.PKID),0) AS ALANC "
  strSql = strSql & " FROM " & _
          " BOLETOATEND " & _
          " INNER JOIN CAIXAATEND ON CAIXAATEND.PKID = BOLETOATEND.CAIXAATENDID " & _
          " INNER JOIN PESSOA ON PESSOA.PKID = CAIXAATEND.ATENDENTEID " & _
          " LEFT JOIN DEBITO ON BOLETOATEND.PKID = DEBITO.BOLETOATENDID " & _
          " WHERE CAIXAATEND.PKID = " & Formata_Dados(RetornaCodTurnoCorrenteAtend(lngFUNCIONARIOID, lngTURNOATENDEPESQ), tpDados_Longo) & _
          " GROUP BY BOLETOATEND.PKID, PESSOA.NOME, BOLETOATEND.NUMERO " & _
          " ORDER BY PESSOA.NOME, BOLETOATEND.NUMERO;"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    RESBOL_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim RESBOL_Matriz(0 To RESBOL_COLUNASMATRIZ - 1, 0 To RESBOL_LINHASMATRIZ - 1)
  Else
    ReDim RESBOL_Matriz(0 To RESBOL_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se j� houver algum item
    For intI = 0 To RESBOL_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda n�o se atingiu fim do recordset
        For intJ = 0 To RESBOL_COLUNASMATRIZ - 1  'varre as colunas
          RESBOL_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'pr�xima linha matriz
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
            " WHERE CAIXAATEND.PKID = " & Formata_Dados(RetornaCodTurnoCorrenteAtend(lngFUNCIONARIOID, lngTURNOATENDEPESQ), tpDados_Longo) & _
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
  If Not objRs.EOF Then   'se j� houver algum item
    For intI = 0 To ENTRAT_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda n�o se atingiu fim do recordset
        For intJ = 0 To ENTRAT_COLUNASMATRIZ - 1  'varre as colunas
          ENTRAT_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'pr�xima linha matriz
  End If
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub
Public Sub MontaATEND_Matriz()

  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMaq.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMaq.clsGeral
  '
  strSql = "SELECT DEBITO.MAQUINAID, DEBITO.BOLETOATENDID, DEBITO.PKID, PESSOA.NOME, BOLETOATEND.NUMERO, DEBITO.NUMERO, EQUIPAMENTO.NUMERO, DEBITO.MEDICAO, DEBITO.VALORPAGO, DEBITO.COEFICIENTE, (ISNULL(DEBITO.VALORPAGO,0) / ISNULL(DEBITO.COEFICIENTE,0)) AS CREDITO, DEBITO.DATA "
  strSql = strSql & " FROM DEBITO " & _
          " INNER JOIN BOLETOATEND ON BOLETOATEND.PKID = DEBITO.BOLETOATENDID " & _
          " INNER JOIN MAQUINA ON MAQUINA.PKID = DEBITO.MAQUINAID " & _
          " INNER JOIN EQUIPAMENTO ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
          " INNER JOIN CAIXAATEND ON CAIXAATEND.PKID = BOLETOATEND.CAIXAATENDID " & _
          " INNER JOIN PESSOA ON PESSOA.PKID = CAIXAATEND.ATENDENTEID " & _
          " WHERE CAIXAATEND.PKID = " & Formata_Dados(RetornaCodTurnoCorrenteAtend(lngFUNCIONARIOID, lngTURNOATENDEPESQ), tpDados_Longo) & _
          " ORDER BY BOLETOATEND.NUMERO, DEBITO.NUMERO;"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    ATEND_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim ATEND_Matriz(0 To ATEND_COLUNASMATRIZ - 1, 0 To ATEND_LINHASMATRIZ - 1)
  Else
    ReDim ATEND_Matriz(0 To ATEND_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se j� houver algum item
    For intI = 0 To ATEND_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda n�o se atingiu fim do recordset
        For intJ = 0 To ATEND_COLUNASMATRIZ - 1  'varre as colunas
          ATEND_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'pr�xima linha matriz
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
  TratarErro Err.Number, Err.Description, "[frmUserOperAteCons.grdAtendenteAtend_UnboundReadDataEx]"
End Sub



Private Sub txtSenha_GotFocus()
  Seleciona_Conteudo_Controle txtUsuario
End Sub
Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strSql                As String
  Dim objRs                 As ADODB.Recordset
  Dim objGeral              As busSisMaq.clsGeral
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCampos = False
  lngFUNCIONARIOID = 0
  lngTURNOATENDID = 0
  If Not Valida_String(txtUsuario, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o usu�rio" & vbCrLf
  End If
  If Not Valida_String(txtSenha, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a senha" & vbCrLf
  End If
  If Len(strMsg) = 0 Then
    'Ok
    'Valida usu�rio
    Set objGeral = New busSisMaq.clsGeral
    strSql = "Select FUNCIONARIO.USUARIO, FUNCIONARIO.SENHA, FUNCIONARIO.NIVEL, FUNCIONARIO.PESSOAID, PESSOA.NOME "
    strSql = strSql & " FROM FUNCIONARIO INNER JOIN PESSOA ON PESSOA.PKID = FUNCIONARIO.PESSOAID "
    strSql = strSql & " INNER JOIN ATENDENTE ON PESSOA.PKID = ATENDENTE.PESSOAID "
    strSql = strSql & " WHERE FUNCIONARIO.SENHA =  " & Formata_Dados(Encripta(UCase$(txtSenha.Text)), tpDados_Texto)
    strSql = strSql & " AND FUNCIONARIO.USUARIO =  " & Formata_Dados(txtUsuario.Text, tpDados_Texto)
    strSql = strSql & " AND FUNCIONARIO.INDEXCLUIDO =  " & Formata_Dados("N", tpDados_Texto)
  
    Set objRs = objGeral.ExecutarSQL(strSql)
    'Verifica se o usu�rio existe
    If objRs.EOF Then
      strMsg = strMsg & "Senha/usu�rio n�o encontrado"
      Pintar_Controle txtSenha, tpCorContr_Erro
      SetarFoco txtSenha
    Else
      lngFUNCIONARIOID = objRs.Fields("PESSOAID").Value & ""
    End If
    '
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
  End If
  If Len(strMsg) = 0 Then
    lngTURNOATENDID = RetornaCodTurnoCorrenteAtend(lngFUNCIONARIOID, lngTURNOATENDEPESQ)
    If lngTURNOATENDID = 0 Then
      strMsg = strMsg & "N�o h� turno aberto para o atendente"
      Pintar_Controle txtSenha, tpCorContr_Erro
      SetarFoco txtSenha
    ElseIf lngTURNOATENDID = -1 Then
      strMsg = strMsg & "H� mais de um turno aberto para o atendente, entre em contato com o administrador do sistema."
      Pintar_Controle txtSenha, tpCorContr_Erro
      SetarFoco txtSenha
    End If
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserOperAteCons.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserOperAteCons.ValidaCampos]", _
            Err.Description
End Function

Private Sub txtSenha_LostFocus()
  On Error GoTo trata
  Pintar_Controle txtSenha, tpCorContr_Normal
  If Me.ActiveControl.Name <> "cmdSelecao" Then Exit Sub
  If Not ValidaCampos Then
    Exit Sub
  End If
  'MsgBox "ok"
  'Montar RecordSet
  Trata_Matrizes_Totais
  '
  SetarFoco cmdSelecao(0)
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserOperAteCons.txtSenha_LostFocus]"
End Sub

Private Sub txtUsuario_GotFocus()
  Seleciona_Conteudo_Controle txtUsuario
End Sub
Private Sub txtUsuario_LostFocus()
  Pintar_Controle txtUsuario, tpCorContr_Normal
End Sub


Private Sub grdBoleto_UnboundReadDataEx( _
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
               Offset + intI, RESBOL_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, RESBOL_COLUNASMATRIZ, RESBOL_LINHASMATRIZ, RESBOL_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, RESBOL_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserOperAteCons.grdBoleto_UnboundReadDataEx]"
End Sub


