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
      Caption         =   "Impressão"
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
      Caption         =   "Selecione a opção"
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
         Caption         =   "&C - Cons. Pedido   "
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
         ToolTipText     =   "Visualização do Pedido"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&E - Cancelar/Ativar"
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
         Left            =   5370
         TabIndex        =   5
         ToolTipText     =   "Cancelar/ativar pedido"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&F - Gerenciar OS    "
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
         Left            =   6690
         TabIndex        =   6
         ToolTipText     =   "Gerenciar OS"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&A - Incluir Pedido    "
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
         ToolTipText     =   "Incluir Pedido"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&B - Alterar Pedido  "
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
         ToolTipText     =   "Alterar Pedido"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&D - Fornecedor        "
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
         Left            =   4050
         TabIndex        =   4
         ToolTipText     =   "Enviar para fornecedor"
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
               TextSave        =   "7/8/2014"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   5
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   1235
               MinWidth        =   1235
               TextSave        =   "04:30"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   1
               Alignment       =   1
               Bevel           =   2
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
      Width           =   1245
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Confirmação de Expiração"
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
      Caption         =   "Fornecedor"
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
      Caption         =   "Movimento após o fechamento"
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
      Caption         =   "Comprador"
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
      Caption         =   "Usuário Logado"
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
Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean
Public blnPrimeiraVez           As Boolean 'Propósito: Preencher lista no combo

'''Public objUserGRInc             As SisMed.frmUserGRInc
'''Public objUserContaCorrente     As SisMed.frmUserContaCorrente


Private COLUNASMATRIZ           As Long
Private LINHASMATRIZ            As Long
Private Matriz()                As String


'''Public Sub Clique_botao(intIndice As Integer)
'''  If cmdSelecao(intIndice).Enabled = True And cmdSelecao(intIndice).Visible = True Then
'''    cmdSelecao_Click intIndice
'''  End If
'''End Sub
'''
'''
'''
Private Sub Form_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  'Tratamento de tecla para verificação de chamada de Outras telas
  'verifica se tem permissão
  'Tudo ok, faz chamada
  Select Case KeyAscii
  Case 1
    'NOVO - IMPRIME PEDIDO EM TELA
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione um Pedido para imprimir o Pedido.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    Report1.Connect = ConnectRpt
    Report1.ReportFileName = gsReportPath & "Pedido.rpt"
    '
    'If optSai1.Value Then
      Report1.Destination = 0 'Video
    'ElseIf optSai2.Value Then
    '  Report1.Destination = 1   'Impressora
    'End If
    Report1.CopiesToPrinter = 1
    Report1.WindowState = crptMaximized
    '
    Report1.Formulas(0) = "PEDIDOID = " & Formata_Dados(grdGeral.Columns("ID").Value, tpDados_Longo)
    '
    Report1.Action = 1
    '
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmGerencial.Form_KeyPress]"
End Sub

'''Private Sub cmdInfFinanc_Click()
'''  On Error GoTo trata
'''  'Chamar o form de Consulta/Visualização das Informações Financeiras.
'''  frmUserInfFinancLis.Show vbModal
'''  SetarFoco grdGeral
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
  Dim objPedidoInc As SisMetal.frmPedidoInc
  Dim objFornecedorSt As SisMetal.frmFornecedorSt
  Dim objPedido As busSisMetal.clsPedido
  Dim objOSLis As SisMetal.frmOSLis
  Dim objEntregaDiretaLis As SisMetal.frmEntregaDiretaLis
  Dim objAjusteLis As SisMetal.frmAjusteLis
  
'''  Dim objGRTotalPrestCons As SisMed.frmUserGRTotalPrestCons
'''  Dim strMsg As String
'''  Dim objGeral      As busSisMetal.clsGeral
'''  Dim objRs         As ADODB.Recordset
'''  Dim strSql        As String
  On Error GoTo trata
'''  '
  Select Case intGrupo

  Case 0
    'Inclusão de Pedido
    frmPedidoItemPedidoInc.Status = tpStatus_Incluir
    frmPedidoItemPedidoInc.lngPEDIDOID = 0
    frmPedidoItemPedidoInc.Show vbModal
  Case 1
    'Alteração do pedido
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione um Pedido para alterá-lo.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") <> "I" And Trim(grdGeral.Columns("Status").Value & "") <> "C" Then
      MsgBox "Somente pedidos com status [INICIAL] ou [COMPRADOR] podem ser alterados.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    frmPedidoItemPedidoInc.Status = tpStatus_Alterar
    frmPedidoItemPedidoInc.lngPEDIDOID = grdGeral.Columns("ID").Value
    frmPedidoItemPedidoInc.Show vbModal
  Case 2
    'Visualização do pedido
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione um Pedido para visualizá-lo.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    frmPedidoItemPedidoInc.Status = tpStatus_Consultar
    frmPedidoItemPedidoInc.lngPEDIDOID = grdGeral.Columns("ID").Value
    frmPedidoItemPedidoInc.Show vbModal
  
  Case 3
    'Alterar status do pedido para fornecedor
    If Trim(grdGeral.Columns("Status").Value & "") <> "C" And Trim(grdGeral.Columns("Status").Value & "") <> "F" Then
      MsgBox "Apenas um pedido no estado comprador pode ser encaminhado para o fornecedor.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    'Encaminhar para fornecedor
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione um Pedido para encaminhá-lo para o fornecedor.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    'Validação dos ítens do pedido
    If Not ValidaCamposEncFornecedor Then
      SetarFoco grdGeral
      Exit Sub
    End If
    If MsgBox("Confirma envio do pedido " & grdGeral.Columns("Ano-OS").Value & " para o fornecedor?", vbYesNo, TITULOSISTEMA) = vbNo Then
      SetarFoco grdGeral
      Exit Sub
    End If
    Set objPedido = New busSisMetal.clsPedido
    If objPedido.ValidaPedidoFechado(grdGeral.Columns("ID").Value) = False Then
      Set objPedido = Nothing
      MsgBox "Pedido possui perfis não distribuidos para anodização e/ou entrega direta.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    objPedido.AlterarStatusFornecedor grdGeral.Columns("ID").Value
    Set objPedido = Nothing
    'NOVO - IMPRIME PEDIDO EM TELA
    Report1.Connect = ConnectRpt
    Report1.ReportFileName = gsReportPath & "Pedido.rpt"
    '
    'If optSai1.Value Then
      Report1.Destination = 0 'Video
    'ElseIf optSai2.Value Then
    '  Report1.Destination = 1   'Impressora
    'End If
    Report1.CopiesToPrinter = 1
    Report1.WindowState = crptMaximized
    '
    Report1.Formulas(0) = "PEDIDOID = " & Formata_Dados(grdGeral.Columns("ID").Value, tpDados_Longo)
    '
    Report1.Action = 1
    '
  Case 4
    'Cancelamento de pedido
    If Not ValidaCamposExclusao Then
      SetarFoco grdGeral
      Exit Sub
    End If
    'Confirmação
    If Trim(grdGeral.Columns("Status").Value & "") = "E" Then
      'Cancelado --> reativar
      If MsgBox("Confirma ativação do pedido " & grdGeral.Columns("Ano-OS").Value & " para o fornecedor " & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
        SetarFoco grdGeral
        Exit Sub
      End If
    Else
      'Ativo --> cancelar
      If MsgBox("Confirma cancelamento do pedido " & grdGeral.Columns("Ano-OS").Value & " para o fornecedor " & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
        SetarFoco grdGeral
        Exit Sub
      End If
    End If
    Set objPedido = New busSisMetal.clsPedido
    objPedido.ExcluirPedido grdGeral.Columns("ID").Value, _
                            Trim(grdGeral.Columns("Status").Value & "")

    Set objPedido = Nothing

  Case 5
    'Gerenciar OS
    Set objOSLis = New SisMetal.frmOSLis
    objOSLis.Show vbModal
    Set objOSLis = Nothing
  Case 6
    'Gerenciar Entrega Direta
    Set objEntregaDiretaLis = New SisMetal.frmEntregaDiretaLis
    objEntregaDiretaLis.Show vbModal
    Set objEntregaDiretaLis = Nothing
'''  Case 7
'''    'Gerenciar Ajustes
'''    Set objAjusteLis = New SisMetal.frmAjusteLis
'''    objAjusteLis.Show vbModal
'''    Set objAjusteLis = Nothing


'''  Case 5
'''    'consultar GR
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "Não há turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para alterá-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    If Trim(RetornaDescAtende(grdGeral.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
'''      MsgBox "Apenas o atendente que lançou a GR pode consultá-la.", vbExclamation, TITULOSISTEMA
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
'''      MsgBox "Não há turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para imprimí-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    If Trim(RetornaNivelAtende(grdGeral.Columns("Atendente").Value & "")) = gsLaboratorio Then
'''      MsgBox "Não pode haver impressão de uma GR lançada pelo Laboratório.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    If Trim(grdGeral.Columns("Status").Value & "") <> "F" Then
'''      If Trim(grdGeral.Columns("Status").Value & "") <> "A" Then
'''        MsgBox "Não pode haver impressão de uma GR que não esteja fechada ou atendida.", vbExclamation, TITULOSISTEMA
'''        SetarFoco grdGeral
'''        Exit Sub
'''      End If
'''    End If
'''    'Confirmação
'''    If MsgBox("Confirma impressão da GR " & grdGeral.Columns("Seq.").Value & " de " & grdGeral.Columns("Prontuário").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    If Trim(grdGeral.Columns("Imp").Value & "") = "S" Then
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
'''          SetarFoco grdGeral
'''          Exit Sub
'''        Else
'''          'Capturou Nome do Usuário, continua com processo
'''        End If
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''    End If
'''
'''    IMP_COMP_GR grdGeral.Columns("ID").Value, gsNomeEmpresa, 1, IIf(Trim(grdGeral.Columns("Imp").Value & "") = "S", True, False)
'''    'Após impressão altera status para impressa
'''    Set objGR = New busSisMetal.clsGR
'''    objGR.AlterarStatusGR grdGeral.Columns("ID").Value, _
'''                          "", _
'''                          "S"



'''    Set objGR = Nothing
'''
'''  Case 7
'''    'Imprimir Comprovante de Recebimento
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "Não há turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
'''      MsgBox "Selecione uma GR para imprimí-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    If Trim(grdGeral.Columns("Status").Value & "") <> "F" Then
'''      If Trim(grdGeral.Columns("Status").Value & "") <> "A" Then
'''        MsgBox "Não pode haver impressão de uma GR que não esteja fechada ou atendida.", vbExclamation, TITULOSISTEMA
'''        SetarFoco grdGeral
'''        Exit Sub
'''      End If
'''    End If
'''    'Confirmação
'''    If MsgBox("Confirma impressão da GR " & grdGeral.Columns("Seq.").Value & " de " & grdGeral.Columns("Prontuário").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''
'''    'Fecou GR do laboratório, emitir comprovante de pagamento
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
  LerFigurasAvulsas cmdInfFinanc, "InfFinanc.ico", "InfFinancDown.ico", "Informações financeiras do turno"
  '
  txtUsuario.Text = gsNomeUsu
  mskDataPrinc.Text = Format(Date, "DD/MM/YYYY")

  'NOVO BOTÕES NOVOS
  ConcederAcessoFnc
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
  TratarErro Err.Number, Err.Description, "[frmGerencial.grdGeral_UnboundReadDataEx]"
End Sub

Public Sub ConcederAcessoFnc()
  On Error GoTo trata
  Select Case gsNivel
  Case gsAdmin
    cmdSelecao(0).Enabled = True
  Case gsDiretor
    cmdSelecao(0).Enabled = True
  Case gsGerente
    cmdSelecao(0).Enabled = True
  Case gsCompra
    cmdSelecao(0).Enabled = False
  End Select
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmGerencial.ValidaCamposExclusao]", _
            Err.Description
End Sub

Private Function ValidaCamposExclusao() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  Dim objGer        As busSisMetal.clsGeral
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  '
  '
  blnSetarFocoControle = True
  ValidaCamposExclusao = False
  '
  On Error GoTo trata
  Set objGer = New busSisMetal.clsGeral
  'ITEM_PEDIDO
  strSql = "Select * from ITEM_PEDIDO WHERE PEDIDOID = " & grdGeral.Columns("ID").Value
  Set objRs = objGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    strMsg = strMsg & "Pedido não pode ser excluido pois já possui itens lançados." & vbCrLf
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set objGer = Nothing
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmGerencial.ValidaCamposExclusao]"
    ValidaCamposExclusao = False
  Else
    ValidaCamposExclusao = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmGerencial.ValidaCamposExclusao]", _
            Err.Description
End Function

Private Function ValidaCamposEncFornecedor() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  Dim objGer        As busSisMetal.clsGeral
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  '
  '
  blnSetarFocoControle = True
  ValidaCamposEncFornecedor = False
  '
  On Error GoTo trata
  Set objGer = New busSisMetal.clsGeral
  'ITEM_PEDIDO
  strSql = "Select * from ITEM_PEDIDO WHERE PEDIDOID = " & grdGeral.Columns("ID").Value & _
      " AND ISNULL(PESO_INI,0) <>  ISNULL(PESO,0) + ISNULL(PESO_FAB,0) "
  Set objRs = objGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    strMsg = strMsg & "Pedido não pode ser encaminhado pois ainda possui ítens não distribuídos." & vbCrLf
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set objGer = Nothing
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmGerencial.ValidaCamposEncFornecedor]"
    ValidaCamposEncFornecedor = False
  Else
    ValidaCamposEncFornecedor = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmGerencial.ValidaCamposEncFornecedor]", _
            Err.Description
End Function


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
  Dim objGeral  As busSisMetal.clsGeral
  Dim strDataCanc As String
  '
  AmpS
  On Error GoTo trata
  '
  strDataCanc = Format(DateAdd("d", -10, Now), "DD/MM/YYYY hh:mm")
  '
  Set objGeral = New busSisMetal.clsGeral
  '
  strSql = "SELECT PEDIDO.PKID, CONVERT(CHAR(4), PEDIDO.OS_ANO) + '-' + CONVERT(VARCHAR(50), PEDIDO.OS_NUMERO) , PEDIDO.DATA, LOJA.NOME, PEDIDO.VALOR_ALUMINIO, " & _
        " CASE PEDIDO.CANCELADO WHEN 'S' THEN 'E' ELSE PEDIDO.STATUS END " & _
        "FROM PEDIDO LEFT JOIN LOJA ON PEDIDO.FORNECEDORID = LOJA.PKID " & _
        " WHERE CANCELADO = " & Formata_Dados("N", tpDados_Texto) & _
        " OR (CANCELADO = " & Formata_Dados("S", tpDados_Texto) & _
        " AND DATA_CANCELAMENTO >= " & Formata_Dados(strDataCanc, tpDados_DataHora) & ") " & _
        " ORDER BY PEDIDO.OS_ANO DESC, PEDIDO.OS_NUMERO DESC;"
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
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub
