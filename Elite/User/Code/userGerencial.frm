VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmGerencial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gerencial de Pedido"
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
      Caption         =   "Pedidos"
      Height          =   6015
      Left            =   60
      TabIndex        =   17
      Top             =   330
      Width           =   11835
      Begin TrueDBGrid60.TDBGrid grdGeral 
         Height          =   5730
         Left            =   90
         OleObjectBlob   =   "userGerencial.frx":0000
         TabIndex        =   0
         Top             =   180
         Width           =   11580
      End
   End
   Begin VB.CommandButton cmdInfFinanc 
      Caption         =   "&Z"
      Height          =   855
      Left            =   11070
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7770
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton cmdSairSelecao 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   855
      Left            =   11070
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6600
      Width           =   900
   End
   Begin VB.Frame fraImpressao 
      Caption         =   "Impress�o"
      Height          =   2085
      Left            =   8460
      TabIndex        =   16
      Top             =   6510
      Width           =   2565
      Begin VB.Label Label72 
         Caption         =   "CTRL + A - Pedido"
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
         TabIndex        =   29
         Top             =   240
         Width           =   2205
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Selecione a op��o"
      Height          =   1335
      Left            =   90
      TabIndex        =   15
      Top             =   6420
      Width           =   8145
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&G - Entrega Direta  "
         Enabled         =   0   'False
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
         Left            =   150
         TabIndex        =   7
         ToolTipText     =   "Gerenciar entrega direta"
         Top             =   630
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&H - Ajustes                "
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
         Left            =   1440
         TabIndex        =   8
         ToolTipText     =   "Gerenciar ajustes no estoque"
         Top             =   630
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&C - Cons. Pacote   "
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
         Left            =   2760
         TabIndex        =   3
         ToolTipText     =   "Consultar Pacote"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&E - Obs. Servi�o      "
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
         Left            =   5460
         TabIndex        =   5
         ToolTipText     =   "Observa��o do Servi�o"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&F - Tranf. Servi�o    "
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
         Left            =   6780
         TabIndex        =   6
         ToolTipText     =   "Transferir Servi�o"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&A - Incluir Pacote    "
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
         Left            =   150
         TabIndex        =   1
         ToolTipText     =   "Incluir Pacote"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&B - Alterar Pacote  "
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
         Left            =   1440
         TabIndex        =   2
         ToolTipText     =   "Alterar Pacote"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&D - Final. Servi�o    "
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
         Left            =   4140
         TabIndex        =   4
         ToolTipText     =   "Finalizar Servi�o"
         Top             =   240
         Width           =   1305
      End
      Begin ComctlLib.StatusBar StatusBar1 
         Height          =   255
         Left            =   60
         TabIndex        =   21
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
               TextSave        =   "10/3/2016"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   5
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   1235
               MinWidth        =   1235
               TextSave        =   "04:54"
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
   Begin Crystal.CrystalReport Report1 
      Left            =   7950
      Top             =   7890
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
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
      Index           =   13
      Left            =   3900
      TabIndex        =   28
      Top             =   8100
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Baixa Total"
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
      Left            =   2580
      TabIndex        =   27
      Top             =   8100
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Baixa Parcial"
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
      TabIndex        =   26
      Top             =   8100
      Visible         =   0   'False
      Width           =   1245
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
      TabIndex        =   25
      Top             =   8340
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Pago"
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
      Left            =   2760
      TabIndex        =   24
      Top             =   7830
      Width           =   975
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
      Left            =   3780
      TabIndex        =   23
      Top             =   7830
      Visible         =   0   'False
      Width           =   2625
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
      TabIndex        =   22
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
      Left            =   6450
      TabIndex        =   20
      Top             =   7830
      Visible         =   0   'False
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
      TabIndex        =   19
      Top             =   7830
      Width           =   525
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Concluido"
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
      TabIndex        =   18
      Top             =   7830
      Width           =   915
   End
   Begin VB.Label Label16 
      Caption         =   "Data"
      Height          =   255
      Left            =   3150
      TabIndex        =   14
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label22 
      Caption         =   "Usu�rio Logado"
      Height          =   255
      Left            =   30
      TabIndex        =   13
      Top             =   30
      Width           =   1215
   End
End
Attribute VB_Name = "frmGerencial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intGrupo                 As Integer
'''Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean
Public blnPrimeiraVez           As Boolean 'Prop�sito: Preencher lista no combo
'''
''''''Public objUserGRInc             As SisMed.frmUserGRInc
''''''Public objUserContaCorrente     As SisMed.frmUserContaCorrente


Private COLUNASMATRIZ           As Long
Private LINHASMATRIZ            As Long
Private Matriz()                As String






Private Sub Form_Load()
  On Error GoTo trata
'''  Dim strSql            As String
'''  Dim datDataTurno      As Date
'''  Dim datDataIniAtual   As Date
'''  Dim datDataFimAtual   As Date
  '
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
  txtUsuario.Text = gsNomeUsu
  mskDataPrinc.Text = Format(Date, "DD/MM/YYYY")

  'NOVO BOT�ES NOVOS
'''  ConcederAcessoFnc
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
End Sub
'''
''''''Public Sub Clique_botao(intIndice As Integer)
''''''  If cmdSelecao(intIndice).Enabled = True And cmdSelecao(intIndice).Visible = True Then
''''''    cmdSelecao_Click intIndice
''''''  End If
''''''End Sub
''''''
''''''
''''''
'''Private Sub Form_KeyPress(KeyAscii As Integer)
'''  On Error GoTo trata
'''  'Tratamento de tecla para verifica��o de chamada de Outras telas
'''  'verifica se tem permiss�o
'''  'Tudo ok, faz chamada
'''  Select Case KeyAscii
'''  Case 1
'''    'NOVO - IMPRIME PEDIDO EM TELA
'''    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione um Pedido para imprimir o Pedido.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    Report1.Connect = ConnectRpt
'''    Report1.ReportFileName = gsReportPath & "Pedido.rpt"
'''    '
'''    'If optSai1.Value Then
'''      Report1.Destination = 0 'Video
'''    'ElseIf optSai2.Value Then
'''    '  Report1.Destination = 1   'Impressora
'''    'End If
'''    Report1.CopiesToPrinter = 1
'''    Report1.WindowState = crptMaximized
'''    '
'''    Report1.Formulas(0) = "PEDIDOID = " & Formata_Dados(grdGeral.Columns("ID").Value, tpDados_Longo)
'''    '
'''    Report1.Action = 1
'''    '
'''  End Select
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, _
'''             Err.Description, _
'''             "[frmGerencial.Form_KeyPress]"
'''End Sub
'''
''''''Private Sub cmdInfFinanc_Click()
''''''  On Error GoTo trata
''''''  'Chamar o form de Consulta/Visualiza��o das Informa��es Financeiras.
''''''  frmUserInfFinancLis.Show vbModal
''''''  SetarFoco grdGeral
''''''  Exit Sub
''''''trata:
''''''  TratarErro Err.Number, _
''''''             Err.Description, _
''''''             Err.Source
''''''  AmpN
''''''End Sub
''''''
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
  intGrupo = Index
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
  If Index <> 4 Then
    SetarFoco grdGeral
  End If
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[cmdSelecao_Click]"
  frmMDI.tmrUnidade.Enabled = True
End Sub


Public Sub VerificaQuemChamou()
'''  Dim objPacoteInc As Elite.frmPacoteInc
  '
  On Error GoTo trata
  '
  Select Case intGrupo

  Case 0
    'Inclus�o de Pacote
    frmPacoteInc.Status = tpStatus_Incluir
    frmPacoteInc.lngPKID = 0
    frmPacoteInc.Show vbModal
  Case 1
    'Altera��o do Pacote
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione um Pacote para alter�-lo.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
'''    If Trim(grdGeral.Columns("Status").Value & "") <> "I" And Trim(grdGeral.Columns("Status").Value & "") <> "C" Then
'''      MsgBox "Somente pedidos com status [INICIAL] ou [COMPRADOR] podem ser alterados.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
    frmPacoteInc.Status = tpStatus_Alterar
    frmPacoteInc.lngPKID = grdGeral.Columns("ID").Value
    frmPacoteInc.Show vbModal
  Case 2
    'Consultar Pacote
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione um Pacote para consult�-lo.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
'''    If Trim(grdGeral.Columns("Status").Value & "") <> "I" And Trim(grdGeral.Columns("Status").Value & "") <> "C" Then
'''      MsgBox "Somente pedidos com status [INICIAL] ou [COMPRADOR] podem ser alterados.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
    frmPacoteInc.Status = tpStatus_Consultar
    frmPacoteInc.lngPKID = grdGeral.Columns("ID").Value
    frmPacoteInc.Show vbModal
  Case 3
    'Finalizar Servi�o do Pacote
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione um Pacote para finalizar um servi�o.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    frmServicoCons.intQuemChamou = 0
    frmServicoCons.lngPACOTEID = grdGeral.Columns("ID").Value
    frmServicoCons.strPacote = grdGeral.Columns(1).Value & _
                              " a " & grdGeral.Columns(2).Value & _
                              " - " & grdGeral.Columns(3).Value & _
                              " - " & grdGeral.Columns(4).Value
                              
    frmServicoCons.Show vbModal
  Case 4
    'Observar Servi�o do Pacote
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione um Pacote para lan�ar uma observa��o no servi�o.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    frmServicoCons.intQuemChamou = 1
    frmServicoCons.lngPACOTEID = grdGeral.Columns("ID").Value
    frmServicoCons.strPacote = grdGeral.Columns(1).Value & _
                              " a " & grdGeral.Columns(2).Value & _
                              " - " & grdGeral.Columns(3).Value & _
                              " - " & grdGeral.Columns(4).Value
                              
    frmServicoCons.Show vbModal

  Case 5
    'Observar Servi�o do Pacote
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione um Pacote para transferir um servi�o.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    frmServicoCons.intQuemChamou = 2
    frmServicoCons.lngPACOTEID = grdGeral.Columns("ID").Value
    frmServicoCons.strPacote = grdGeral.Columns(1).Value & _
                              " a " & grdGeral.Columns(2).Value & _
                              " - " & grdGeral.Columns(3).Value & _
                              " - " & grdGeral.Columns(4).Value
                              
    frmServicoCons.Show vbModal




'''  Case 3
'''    'Alterar status do pedido para fornecedor
'''    If Trim(grdGeral.Columns("Status").Value & "") <> "C" And Trim(grdGeral.Columns("Status").Value & "") <> "F" Then
'''      MsgBox "Apenas um pedido no estado comprador pode ser encaminhado para o fornecedor.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    'Encaminhar para fornecedor
'''    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione um Pedido para encaminh�-lo para o fornecedor.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    'Valida��o dos �tens do pedido
'''    If Not ValidaCamposEncFornecedor Then
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    If MsgBox("Confirma envio do pedido " & grdGeral.Columns("Ano-OS").Value & " para o fornecedor?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    Set objPedido = New busElite.clsPedido
'''    If objPedido.ValidaPedidoFechado(grdGeral.Columns("ID").Value) = False Then
'''      Set objPedido = Nothing
'''      MsgBox "Pedido possui perfis n�o distribuidos para anodiza��o e/ou entrega direta.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    objPedido.AlterarStatusFornecedor grdGeral.Columns("ID").Value
'''    Set objPedido = Nothing
'''    'NOVO - IMPRIME PEDIDO EM TELA
'''    Report1.Connect = ConnectRpt
'''    Report1.ReportFileName = gsReportPath & "Pedido.rpt"
'''    '
'''    'If optSai1.Value Then
'''      Report1.Destination = 0 'Video
'''    'ElseIf optSai2.Value Then
'''    '  Report1.Destination = 1   'Impressora
'''    'End If
'''    Report1.CopiesToPrinter = 1
'''    Report1.WindowState = crptMaximized
'''    '
'''    Report1.Formulas(0) = "PEDIDOID = " & Formata_Dados(grdGeral.Columns("ID").Value, tpDados_Longo)
'''    '
'''    Report1.Action = 1
'''    '
'''  Case 4
'''    'Cancelamento de pedido
'''    If Not ValidaCamposExclusao Then
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    'Confirma��o
'''    If Trim(grdGeral.Columns("Status").Value & "") = "E" Then
'''      'Cancelado --> reativar
'''      If MsgBox("Confirma ativa��o do pedido " & grdGeral.Columns("Ano-OS").Value & " para o fornecedor " & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''        SetarFoco grdGeral
'''        Exit Sub
'''      End If
'''    Else
'''      'Ativo --> cancelar
'''      If MsgBox("Confirma cancelamento do pedido " & grdGeral.Columns("Ano-OS").Value & " para o fornecedor " & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''        SetarFoco grdGeral
'''        Exit Sub
'''      End If
'''    End If
'''    Set objPedido = New busElite.clsPedido
'''    objPedido.ExcluirPedido grdGeral.Columns("ID").Value, _
'''                            Trim(grdGeral.Columns("Status").Value & "")
'''
'''    Set objPedido = Nothing
'''
'''  Case 5
'''    'Gerenciar OS
'''    Set objOSLis = New Elite.frmOSLis
'''    objOSLis.Show vbModal
'''    Set objOSLis = Nothing
'''  Case 6
'''    'Gerenciar Entrega Direta
'''    Set objEntregaDiretaLis = New Elite.frmEntregaDiretaLis
'''    objEntregaDiretaLis.Show vbModal
'''    Set objEntregaDiretaLis = Nothing
'''  Case 7
'''    'Gerenciar Ajustes
'''    Set objAjusteLis = New Elite.frmAjusteLis
'''    objAjusteLis.Show vbModal
'''    Set objAjusteLis = Nothing


'''  Case 5
'''    'consultar GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para alter�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    If Trim(RetornaDescAtende(grdGeral.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
'''      MsgBox "Apenas o atendente que lan�ou a GR pode consult�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    Set objUserGRInc = New SisMed.frmUserGRInc
'''    objUserGRInc.Status = tpStatus_Consultar
'''    objUserGRInc.IcEstadoGR = tpIcEstadoGR_Con
'''    objUserGRInc.lngGRID = grdGeral.Columns("ID").Value
'''    objUserGRInc.Show vbModal
'''    Set objUserGRInc = Nothing
'''  Case 6
'''    'Imprimir GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para imprim�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    If Trim(RetornaNivelAtende(grdGeral.Columns("Atendente").Value & "")) = gsLaboratorio Then
'''      MsgBox "N�o pode haver impress�o de uma GR lan�ada pelo Laborat�rio.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    If Trim(grdGeral.Columns("Status").Value & "") <> "F" Then
'''      If Trim(grdGeral.Columns("Status").Value & "") <> "A" Then
'''        MsgBox "N�o pode haver impress�o de uma GR que n�o esteja fechada ou atendida.", vbExclamation, TITULOSISTEMA
'''        SetarFoco grdGeral
'''        Exit Sub
'''      End If
'''    End If
'''    'Confirma��o
'''    If MsgBox("Confirma impress�o da GR " & grdGeral.Columns("Seq.").Value & " de " & grdGeral.Columns("Prontu�rio").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    If Trim(grdGeral.Columns("Imp").Value & "") = "S" Then
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
'''          SetarFoco grdGeral
'''          Exit Sub
'''        Else
'''          'Capturou Nome do Usu�rio, continua com processo
'''        End If
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''    End If
'''
'''    IMP_COMP_GR grdGeral.Columns("ID").Value, gsNomeEmpresa, 1, IIf(Trim(grdGeral.Columns("Imp").Value & "") = "S", True, False)
'''    'Ap�s impress�o altera status para impressa
'''    Set objGR = New busElite.clsGR
'''    objGR.AlterarStatusGR grdGeral.Columns("ID").Value, _
'''                          "", _
'''                          "S"



'''    Set objGR = Nothing
'''
'''  Case 7
'''    'Imprimir Comprovante de Recebimento
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "N�o h� turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para imprim�-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    If Trim(grdGeral.Columns("Status").Value & "") <> "F" Then
'''      If Trim(grdGeral.Columns("Status").Value & "") <> "A" Then
'''        MsgBox "N�o pode haver impress�o de uma GR que n�o esteja fechada ou atendida.", vbExclamation, TITULOSISTEMA
'''        SetarFoco grdGeral
'''        Exit Sub
'''      End If
'''    End If
'''    'Confirma��o
'''    If MsgBox("Confirma impress�o da GR " & grdGeral.Columns("Seq.").Value & " de " & grdGeral.Columns("Prontu�rio").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''
'''    'Fecou GR do laborat�rio, emitir comprovante de pagamento
'''    IMP_COMPROV_REC grdGeral.Columns("ID").Value, gsNomeEmpresa, 1
'''  Case 8
'''    'Canlelar GR outros prestadores
'''    Set objGRTotalPrestCons = New SisMed.frmUserGRTotalPrestCons
'''    objGRTotalPrestCons.Show vbModal
'''    Set objGRTotalPrestCons = Nothing
  End Select
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  End
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
  TratarErro Err.Number, Err.Description, "[frmGerencial.grdGeral_UnboundReadDataEx]"
End Sub
'''
'''Public Sub ConcederAcessoFnc()
'''  On Error GoTo trata
'''  Select Case gsNivel
'''  Case gsAdmin
'''    cmdSelecao(0).Enabled = True
'''  Case gsDiretor
'''    cmdSelecao(0).Enabled = True
'''  Case gsGerente
'''    cmdSelecao(0).Enabled = True
'''  Case gsCompra
'''    cmdSelecao(0).Enabled = False
'''  End Select
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[frmGerencial.ValidaCamposExclusao]", _
'''            Err.Description
'''End Sub
'''
'''Private Function ValidaCamposExclusao() As Boolean
'''  On Error GoTo trata
'''  Dim strMsg                As String
'''  Dim blnSetarFocoControle  As Boolean
'''  Dim objGer        As busElite.clsGeral
'''  Dim objRs         As ADODB.Recordset
'''  Dim strSql        As String
'''  '
'''  '
'''  blnSetarFocoControle = True
'''  ValidaCamposExclusao = False
'''  '
'''  On Error GoTo trata
'''  Set objGer = New busElite.clsGeral
'''  'ITEM_PEDIDO
'''  strSql = "Select * from ITEM_PEDIDO WHERE PEDIDOID = " & grdGeral.Columns("ID").Value
'''  Set objRs = objGer.ExecutarSQL(strSql)
'''  If Not objRs.EOF Then
'''    strMsg = strMsg & "Pedido n�o pode ser excluido pois j� possui itens lan�ados." & vbCrLf
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
'''  '
'''  Set objGer = Nothing
'''  '
'''  If Len(strMsg) <> 0 Then
'''    TratarErroPrevisto strMsg, "[frmGerencial.ValidaCamposExclusao]"
'''    ValidaCamposExclusao = False
'''  Else
'''    ValidaCamposExclusao = True
'''  End If
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[frmGerencial.ValidaCamposExclusao]", _
'''            Err.Description
'''End Function
'''
'''Private Function ValidaCamposEncFornecedor() As Boolean
'''  On Error GoTo trata
'''  Dim strMsg                As String
'''  Dim blnSetarFocoControle  As Boolean
'''  Dim objGer        As busElite.clsGeral
'''  Dim objRs         As ADODB.Recordset
'''  Dim strSql        As String
'''  '
'''  '
'''  blnSetarFocoControle = True
'''  ValidaCamposEncFornecedor = False
'''  '
'''  On Error GoTo trata
'''  Set objGer = New busElite.clsGeral
'''  'ITEM_PEDIDO
'''  strSql = "Select * from ITEM_PEDIDO WHERE PEDIDOID = " & grdGeral.Columns("ID").Value & _
'''      " AND ISNULL(PESO_INI,0) <>  ISNULL(PESO,0) + ISNULL(PESO_FAB,0) "
'''  Set objRs = objGer.ExecutarSQL(strSql)
'''  If Not objRs.EOF Then
'''    strMsg = strMsg & "Pedido n�o pode ser encaminhado pois ainda possui �tens n�o distribu�dos." & vbCrLf
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
'''  '
'''  Set objGer = Nothing
'''  '
'''  If Len(strMsg) <> 0 Then
'''    TratarErroPrevisto strMsg, "[frmGerencial.ValidaCamposEncFornecedor]"
'''    ValidaCamposEncFornecedor = False
'''  Else
'''    ValidaCamposEncFornecedor = True
'''  End If
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[frmGerencial.ValidaCamposEncFornecedor]", _
'''            Err.Description
'''End Function


Private Sub Form_Activate()
  On Error GoTo trata
  '
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
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Public Sub MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGeral  As busElite.clsGeral
  Dim strDataCanc As String
  '
  AmpS
  On Error GoTo trata
  '
  strDataCanc = Format(DateAdd("d", -10, Now), "DD/MM/YYYY hh:mm")
  '
  Set objGeral = New busElite.clsGeral
  '
  strSql = "SELECT PACOTE.PKID, PACOTE.DATAINICIO, PACOTE.DATATERMINO, PESSOA.NOME, PACOTE.VALOR, " & _
        " PACOTE.STATUS " & _
        "FROM PACOTE LEFT JOIN MOTORISTA ON MOTORISTA.PESSOAID = PACOTE.MOTORISTAID " & _
        " LEFT JOIN PESSOA ON PESSOA.PKID = MOTORISTA.PESSOAID " & _
        " ORDER BY PACOTE.DATAINICIO DESC, PACOTE.DATATERMINO DESC;"
  '
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
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
  Set objGeral = Nothing
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub
