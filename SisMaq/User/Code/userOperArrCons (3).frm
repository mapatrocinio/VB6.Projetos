VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserOperArrCons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operacional Arrecadador"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraUnidade 
      Caption         =   "Operacional Arrecadador"
      Height          =   6015
      Left            =   60
      TabIndex        =   11
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
      Begin TrueDBGrid60.TDBGrid grdArrecadador 
         Height          =   2940
         Left            =   90
         OleObjectBlob   =   "userOperArrCons.frx":0000
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2910
         Width           =   11640
      End
      Begin MSMask.MaskEdBox mskArrecadado 
         Height          =   255
         Left            =   4230
         TabIndex        =   2
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
      Begin TrueDBGrid60.TDBGrid grdBoleto 
         Height          =   1740
         Left            =   120
         OleObjectBlob   =   "userOperArrCons.frx":9018
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1170
         Width           =   5820
      End
      Begin VB.Label Label5 
         Caption         =   "Arrecadado"
         Height          =   285
         Index           =   3
         Left            =   2970
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Usu�rio"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   15
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Senha"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   14
         Top             =   615
         Width           =   855
      End
   End
   Begin VB.Frame fraImpressao 
      Caption         =   "Impress�o"
      Height          =   1065
      Left            =   7770
      TabIndex        =   9
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
         TabIndex        =   13
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
         TabIndex        =   10
         Top             =   210
         Width           =   2205
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Selecione a op��o"
      Height          =   1065
      Left            =   60
      TabIndex        =   8
      Top             =   6420
      Width           =   7665
      Begin VB.CommandButton cmdSairSelecao 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   855
         Left            =   6690
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   150
         Width           =   900
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&B - Deslogar             "
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
         TabIndex        =   16
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
         TabIndex        =   3
         ToolTipText     =   "Boleto"
         Top             =   240
         Width           =   1305
      End
      Begin ComctlLib.StatusBar StatusBar1 
         Height          =   255
         Left            =   2670
         TabIndex        =   12
         Top             =   750
         Width           =   4140
         _ExtentX        =   7303
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
               TextSave        =   "19/9/2010"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   5
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   1235
               MinWidth        =   1235
               TextSave        =   "15:33"
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
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "txtTurno"
      Top             =   30
      Width           =   4815
   End
   Begin MSMask.MaskEdBox mskDataPrinc 
      Height          =   255
      Left            =   900
      TabIndex        =   4
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
      TabIndex        =   7
      Top             =   60
      Width           =   765
   End
   Begin VB.Label Label21 
      Caption         =   "Turno Corrente"
      Height          =   255
      Left            =   2190
      TabIndex        =   6
      Top             =   60
      Width           =   1245
   End
End
Attribute VB_Name = "frmUserOperArrCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public nGrupo                         As Integer
'''Public Status                   As tpStatus
Public blnRetorno                     As Boolean
Public blnFechar                      As Boolean
Private lngFUNCIONARIOID              As Long
Private lngTURNOARRECID               As Long
'''
'''Public objUserGRInc             As SisMaq.frmUserGRInc
'''Public objUserContaCorrente     As SisMaq.frmUserContaCorrente
'''
Public blnPrimeiraVez                 As Boolean 'Prop�sito: Preencher lista no combo

'Resumo Boleto
Private RESBOL_COLUNASMATRIZ            As Long
Private RESBOL_LINHASMATRIZ             As Long
Private RESBOL_Matriz()                 As String
'Entrada Arrecadador
Private ARREC_COLUNASMATRIZ            As Long
Private ARREC_LINHASMATRIZ             As Long
Private ARREC_Matriz()                 As String


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
'''    'DETALHAR ENTRADA ARRECENTE
'''    frmUserEntradaAtendLis.Show vbModal
'''    Form_Load
'''  Case 6
'''    'DETALHAR BOLETO ARRECENTE
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
             "[frmUserOperArrCons.Form_KeyPress]"
End Sub

'''Private Sub cmdInfFinanc_Click()
'''  On Error GoTo trata
'''  'Chamar o form de Consulta/Visualiza��o das Informa��es Financeiras.
'''  frmUserInfFinancLis.Show vbModal
'''  SetarFoco grdArrecadador
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
  Dim objUserBoletoCredInc As SisMaq.frmUserBoletoCredInc
'  Dim objGRTotalPrestCons As SisMaq.frmUserGRTotalPrestCons
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
    Set objUserBoletoCredInc = New SisMaq.frmUserBoletoCredInc
    objUserBoletoCredInc.Status = tpStatus_Incluir
    objUserBoletoCredInc.lngBOLETOARRECID = 0
    objUserBoletoCredInc.txtUsuario = txtUsuario.Text
    objUserBoletoCredInc.txtSenha = txtSenha.Text
    objUserBoletoCredInc.Show vbModal
    Set objUserBoletoCredInc = Nothing
  Case 1
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
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdArrecadador.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para alter�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Trim(RetornaDescAtende(grdArrecadador.Columns("Arrecadador").Value & "")) <> gsNomeUsuCompleto Then
'''      MsgBox "Apenas o atendente que lan�ou a GR pode alter�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Trim(grdArrecadador.Columns("Status").Value & "") = "C" Then
'''      MsgBox "N�o pode haver altera��o em uma GR cancelada.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Trim(grdArrecadador.Columns("Status").Value & "") = "F" Then
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
'''          SetarFoco grdArrecadador
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
'''    objUserGRInc.lngGRID = grdArrecadador.Columns("ID").Value
'''    objUserGRInc.Show vbModal
'''    Set objUserGRInc = Nothing
'''  Case 2
'''    'Itens da GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdArrecadador.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para alterar seus �tens.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Trim(RetornaDescAtende(grdArrecadador.Columns("Arrecadador").Value & "")) <> gsNomeUsuCompleto Then
'''      MsgBox "Apenas o atendente que lan�ou a GR pode alter�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Trim(grdArrecadador.Columns("Status").Value & "") = "C" Then
'''      MsgBox "N�o pode haver altera��o de �tens de uma GR cancelada.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Trim(grdArrecadador.Columns("Status").Value & "") = "F" Then
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
'''          SetarFoco grdArrecadador
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
'''    objUserGRInc.lngGRID = grdArrecadador.Columns("ID").Value
'''    objUserGRInc.Show vbModal
'''    Set objUserGRInc = Nothing
'''  Case 3
'''    'Altera��o de pagamento da GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdArrecadador.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para alterar seus dados de pagamento.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Trim(RetornaDescAtende(grdArrecadador.Columns("Arrecadador").Value & "")) <> gsNomeUsuCompleto Then
'''      If (gsNivel <> gsLaboratorio) And (Trim(RetornaNivelAtende(grdArrecadador.Columns("Arrecadador").Value & "")) <> gsLaboratorio) Then
'''        MsgBox "Apenas o atendente que lan�ou a GR pode efetuar o seu pagamento ou uma GR lan�ada pelo Laborat�rio.", vbExclamation, TITULOSISTEMA
'''        SetarFoco grdArrecadador
'''        Exit Sub
'''      End If
'''    End If
'''    If Trim(grdArrecadador.Columns("Status").Value & "") = "C" Then
'''      MsgBox "N�o pode haver pagamento de uma GR cancelada.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Trim(grdArrecadador.Columns("Status").Value & "") = "F" Then
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
'''          SetarFoco grdArrecadador
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
'''    objUserContaCorrente.lngGRID = grdArrecadador.Columns("ID").Value
'''    objUserContaCorrente.intGrupo = 0
'''    objUserContaCorrente.strFuncionarioNome = gsNomeUsuCompleto
'''    objUserContaCorrente.Status = tpStatus_Incluir
'''    objUserContaCorrente.strStatusLanc = "RC"
'''    objUserContaCorrente.strNivelAcesso = Trim(RetornaNivelAtende(grdArrecadador.Columns("Arrecadador").Value & ""))
'''    objUserContaCorrente.Show vbModal
'''    Set objUserContaCorrente = Nothing
'''  Case 4
'''    'Cancelamento da GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdArrecadador.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para exclu�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Trim(RetornaDescAtende(grdArrecadador.Columns("Arrecadador").Value & "")) <> gsNomeUsuCompleto Then
'''      If Mid(grdArrecadador.Columns("Arrecadador").Value & "", 2, 3) <> gsLaboratorio Then
'''        MsgBox "Apenas o atendente que lan�ou a GR pode exclu�-la.", vbExclamation, TITULOSISTEMA
'''        SetarFoco grdArrecadador
'''        Exit Sub
'''      End If
'''    End If
'''    'If Trim(grdArrecadador.Columns("Status").Value & "") <> "F" Then
'''    '  MsgBox "Apenas pode de excluida uma GR fechada.", vbExclamation, TITULOSISTEMA
'''    '  SetarFoco grdArrecadador
'''    '  Exit Sub
'''    'End If
'''    'If Trim(grdArrecadador.Columns("Status").Value & "") = "F" Then
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
'''          SetarFoco grdArrecadador
'''          Exit Sub
'''        Else
'''          'Capturou Nome do Usu�rio, continua com processo
'''        End If
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''    'End If
'''    'Confirma��o
'''    If MsgBox("Confirma cancelamento da GR " & grdArrecadador.Columns("Seq.").Value & " de " & grdArrecadador.Columns("Prontu�rio").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''
'''    Set objGR = New busSisMaq.clsGR
'''    objGR.AlterarStatusGR grdArrecadador.Columns("ID").Value, _
'''                          "C", _
'''                          "", _
'''                          RetornaCodTurnoCorrente
'''    Set objGR = Nothing
'''    IMP_COMP_CANC_GR grdArrecadador.Columns("ID").Value, gsNomeEmpresa, 1
'''  Case 5
'''    'consultar GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdArrecadador.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para alter�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Trim(RetornaDescAtende(grdArrecadador.Columns("Arrecadador").Value & "")) <> gsNomeUsuCompleto Then
'''      MsgBox "Apenas o atendente que lan�ou a GR pode consult�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    Set objUserGRInc = New SisMaq.frmUserGRInc
'''    objUserGRInc.Status = tpStatus_Consultar
'''    objUserGRInc.IcEstadoGR = tpIcEstadoGR_Con
'''    objUserGRInc.lngGRID = grdArrecadador.Columns("ID").Value
'''    objUserGRInc.Show vbModal
'''    Set objUserGRInc = Nothing
'''  Case 6
'''    'Imprimir GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdArrecadador.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para imprim�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Trim(grdArrecadador.Columns("Status").Value & "") <> "F" Then
'''      If Trim(RetornaNivelAtende(grdArrecadador.Columns("Arrecadador").Value & "")) <> gsLaboratorio Then
'''        MsgBox "N�o pode haver impress�o de uma GR que n�o esteja fechada ou seja lan�ada pelo Laborat�rio.", vbExclamation, TITULOSISTEMA
'''        SetarFoco grdArrecadador
'''        Exit Sub
'''      End If
'''    End If
'''    'Confirma��o
'''    If MsgBox("Confirma impress�o da GR " & grdArrecadador.Columns("Seq.").Value & " de " & grdArrecadador.Columns("Prontu�rio").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Trim(grdArrecadador.Columns("Imp").Value & "") = "S" Then
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
'''          SetarFoco grdArrecadador
'''          Exit Sub
'''        Else
'''          'Capturou Nome do Usu�rio, continua com processo
'''        End If
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''    End If
'''
'''    IMP_COMP_GR grdArrecadador.Columns("ID").Value, gsNomeEmpresa, 1, IIf(Trim(grdArrecadador.Columns("Imp").Value & "") = "S", True, False)
'''    'Ap�s impress�o altera status para impressa
'''    Set objGR = New busSisMaq.clsGR
'''    objGR.AlterarStatusGR grdArrecadador.Columns("ID").Value, _
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
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdArrecadador.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para imprim�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    If Trim(grdArrecadador.Columns("Status").Value & "") <> "F" Then
'''      MsgBox "N�o pode haver impress�o de uma GR que n�o esteja fechada.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''    'Confirma��o
'''    If MsgBox("Confirma impress�o da GR " & grdArrecadador.Columns("Seq.").Value & " de " & grdArrecadador.Columns("Prontu�rio").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco grdArrecadador
'''      Exit Sub
'''    End If
'''
'''    'Fecou GR do laborat�rio, emitir comprovante de pagamento
'''    IMP_COMPROV_REC grdArrecadador.Columns("ID").Value, gsNomeEmpresa, 1
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
  Me.Height = 8025
  Me.Width = 12090
  CenterForm Me
  '
  LerFigurasAvulsas cmdSairSelecao, "Sair.ico", "SairDown.ico", "Sair"
  '
  txtTurno.Text = RetornaDescTurnoCorrente
  mskDataPrinc.Text = Format(Date, "DD/MM/YYYY")

  'NOVO BOT�ES NOVOS
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
End Sub

Private Sub grdArrecadador_UnboundReadDataEx( _
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
               Offset + intI, ARREC_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, ARREC_COLUNASMATRIZ, ARREC_LINHASMATRIZ, ARREC_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ARREC_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserOperArrCons.grdArrecadador_UnboundReadDataEx]"
End Sub


Public Sub Trata_Matrizes_Totais()
  On Error GoTo trata
  Dim strSql          As String
  Dim objRs           As ADODB.Recordset
  Dim objGeral        As busSisMaq.clsGeral
  Dim curSaldoArrec   As Currency
  'Boleto Arrecadador
  RESBOL_COLUNASMATRIZ = grdBoleto.Columns.Count
  RESBOL_LINHASMATRIZ = 0
  MontaRESBOL_Matriz
  grdBoleto.Bookmark = Null
  grdBoleto.ReBind
  grdBoleto.ApproxCount = RESBOL_LINHASMATRIZ
  blnPrimeiraVez = False
  'Arrecadador
  ARREC_COLUNASMATRIZ = grdArrecadador.Columns.Count
  ARREC_LINHASMATRIZ = 0
  MontaARREC_Matriz
  grdArrecadador.Bookmark = Null
  grdArrecadador.ReBind
  grdArrecadador.ApproxCount = ARREC_LINHASMATRIZ
  blnPrimeiraVez = False
  'Monta saldo
  curSaldoArrec = 0
  '
  Set objGeral = New busSisMaq.clsGeral
  '
  strSql = "SELECT ISNULL(SUM(CREDITO.VALORPAGO), 0) AS VALOR "
  strSql = strSql & " FROM CREDITO " & _
          " INNER JOIN BOLETOARREC ON BOLETOARREC.PKID = CREDITO.BOLETOARRECID " & _
          " INNER JOIN MAQUINA ON MAQUINA.PKID = CREDITO.MAQUINAID " & _
          " INNER JOIN EQUIPAMENTO ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
          " INNER JOIN CAIXAARREC ON CAIXAARREC.PKID = BOLETOARREC.CAIXAARRECID " & _
          " INNER JOIN PESSOA ON PESSOA.PKID = CAIXAARREC.ARRECADADORID " & _
          " WHERE CAIXAARREC.PKID = " & Formata_Dados(RetornaCodTurnoCorrenteArrec(lngFUNCIONARIOID), tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    curSaldoArrec = objRs.Fields("VALOR").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set objGeral = Nothing
  '
  INCLUIR_VALOR_NO_MASK mskArrecadado, curSaldoArrec, TpMaskMoeda
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
  strSql = "SELECT BOLETOARREC.PKID, MAX(PESSOA.NOME), BOLETOARREC.NUMERO, MAX(BOLETOARREC.STATUS), ISNULL(COUNT(CREDITO.PKID),0) AS LANCADO, 10 - ISNULL(COUNT(CREDITO.PKID),0) AS ALANC "
  strSql = strSql & " FROM " & _
          " BOLETOARREC " & _
          " INNER JOIN CAIXAARREC ON CAIXAARREC.PKID = BOLETOARREC.CAIXAARRECID " & _
          " INNER JOIN PESSOA ON PESSOA.PKID = CAIXAARREC.ARRECADADORID " & _
          " LEFT JOIN CREDITO ON BOLETOARREC.PKID = CREDITO.BOLETOARRECID " & _
          " WHERE CAIXAARREC.PKID = " & Formata_Dados(RetornaCodTurnoCorrenteArrec(lngFUNCIONARIOID), tpDados_Longo) & _
          " GROUP BY BOLETOARREC.PKID, PESSOA.NOME, BOLETOARREC.NUMERO " & _
          " ORDER BY PESSOA.NOME, BOLETOARREC.NUMERO;"
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
Public Sub MontaARREC_Matriz()

  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMaq.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMaq.clsGeral
  '
  strSql = "SELECT CREDITO.MAQUINAID, CREDITO.BOLETOARRECID, PESSOA.NOME, BOLETOARREC.NUMERO, CREDITO.NUMERO, EQUIPAMENTO.NUMERO, CREDITO.MEDICAO, CREDITO.VALORPAGO, CREDITO.COEFICIENTE, (ISNULL(CREDITO.VALORPAGO,0) * ISNULL(CREDITO.COEFICIENTE,0)) AS CREDITO, CREDITO.DATA "
  strSql = strSql & " FROM CREDITO " & _
          " INNER JOIN BOLETOARREC ON BOLETOARREC.PKID = CREDITO.BOLETOARRECID " & _
          " INNER JOIN MAQUINA ON MAQUINA.PKID = CREDITO.MAQUINAID " & _
          " INNER JOIN EQUIPAMENTO ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
          " INNER JOIN CAIXAARREC ON CAIXAARREC.PKID = BOLETOARREC.CAIXAARRECID " & _
          " INNER JOIN PESSOA ON PESSOA.PKID = CAIXAARREC.ARRECADADORID " & _
          " WHERE CAIXAARREC.PKID = " & Formata_Dados(RetornaCodTurnoCorrenteArrec(lngFUNCIONARIOID), tpDados_Longo) & _
          " ORDER BY BOLETOARREC.NUMERO, CREDITO.NUMERO;"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    ARREC_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim ARREC_Matriz(0 To ARREC_COLUNASMATRIZ - 1, 0 To ARREC_LINHASMATRIZ - 1)
  Else
    ReDim ARREC_Matriz(0 To ARREC_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se j� houver algum item
    For intI = 0 To ARREC_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda n�o se atingiu fim do recordset
        For intJ = 0 To ARREC_COLUNASMATRIZ - 1  'varre as colunas
          ARREC_Matriz(intJ, intI) = objRs(intJ) & ""
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
  lngTURNOARRECID = 0
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
    strSql = strSql & " INNER JOIN ARRECADADOR ON PESSOA.PKID = ARRECADADOR.PESSOAID "
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
    lngTURNOARRECID = RetornaCodTurnoCorrenteArrec(lngFUNCIONARIOID)
    If lngTURNOARRECID = 0 Then
      strMsg = strMsg & "N�o h� turno aberto para o arrecadador"
      Pintar_Controle txtSenha, tpCorContr_Erro
      SetarFoco txtSenha
    ElseIf lngTURNOARRECID = -1 Then
      strMsg = strMsg & "H� mais de um turno aberto para o arrecadador, entre em contato com o administrador do sistema."
      Pintar_Controle txtSenha, tpCorContr_Erro
      SetarFoco txtSenha
    End If
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserOperArrCons.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserOperArrCons.ValidaCampos]", _
            Err.Description
End Function

Private Sub txtSenha_LostFocus()
  On Error GoTo trata
  Pintar_Controle txtUsuario, tpCorContr_Normal
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
  TratarErro Err.Number, Err.Description, "[frmUserOperArrCons.txtSenha_LostFocus]"
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
  TratarErro Err.Number, Err.Description, "[frmUserOperArrCons.grdBoleto_UnboundReadDataEx]"
End Sub


