VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmGerencialPed 
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
         OleObjectBlob   =   "userGerencialPed.frx":0000
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
         Caption         =   "CTRL + B - Comissão"
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
         TabIndex        =   30
         Top             =   420
         Width           =   2205
      End
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
         Left            =   1440
         TabIndex        =   7
         ToolTipText     =   "Gerenciar entrega direta"
         Top             =   630
         Visible         =   0   'False
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
         Left            =   2730
         TabIndex        =   8
         ToolTipText     =   "Gerenciar ajustes no estoque"
         Top             =   630
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&C - Empresa              "
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
         ToolTipText     =   "Incluir Pedido Empresa"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&E - Canc./Ativar       "
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
         Left            =   150
         TabIndex        =   6
         ToolTipText     =   "Gerenciar OS"
         Top             =   630
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&A - Balcão                 "
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
         ToolTipText     =   "Incluir Pedido Balcão"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&B - Cliente                  "
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
         ToolTipText     =   "Incluir Pedido Cliente"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&D - Alterar Pedido  "
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
         ToolTipText     =   "Alterar Pedido"
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
               TextSave        =   "14/7/2015"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   5
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   1235
               MinWidth        =   1235
               TextSave        =   "16:04"
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
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
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
      Index           =   13
      Left            =   4830
      TabIndex        =   28
      Top             =   7830
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
      Caption         =   "Recebimento"
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
      Left            =   1200
      TabIndex        =   26
      Top             =   8070
      Visible         =   0   'False
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
      Caption         =   "Empresa"
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
      Left            =   5010
      TabIndex        =   23
      Top             =   8130
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
      Caption         =   "Excluída"
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
      Left            =   3780
      TabIndex        =   20
      Top             =   7830
      Width           =   1035
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Balcão"
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
      Caption         =   "Cliente"
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
Attribute VB_Name = "frmGerencialPed"
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

Private COLUNASMATRIZ           As Long
Private LINHASMATRIZ            As Long
Private Matriz()                As String

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
    IMP_COMP_PEDIDO grdGeral.Columns("ID").Value, gsNomeEmpresa
    
    '
  Case 2
    'IMPRIME COMISSÃO DE VENDEDORES
    frmRelDemoVendas.Show vbModal
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmGerencial.Form_KeyPress]"
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
  Dim strMsg As String
  Dim objPedidoVenda As busSisMetal.clsPedidoVenda
  
  On Error GoTo trata
  Select Case intGrupo

  Case 0, 1, 2
    'Inclusão de Pedido Balcão, Cliente e Empresa
    '----------------------------
    '----------------------------
    'Pede Senha Superior (Diretor, Gerente ou Administrador
    If gsNivel = "LOJ" Or gsNivel = "ADM" Then
      'Só pede senha superior se quem estiver logado não for superior
      frmLoginSup.Show vbModal
      
      If Len(Trim(gsNomeUsuLib)) = 0 Then
        strMsg = "Para efetuar a venda é necessário se autenticar no sistema."
        TratarErroPrevisto strMsg, "cmdConfirmar_Click"
        Exit Sub
      ElseIf gsNivelUsuLib <> "LOJ" Then
        strMsg = "Para efetuar a venda é necessário ter o perfil de vendedodor."
        TratarErroPrevisto strMsg, "cmdConfirmar_Click"
        Exit Sub
        
      End If
      '
      'Capturou Nome do Usuário, continua processo de Venda
    Else
      MsgBox "Apenas o perfil Caixa-Vendedor pode realizar vendas.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    '--------------------------------
    '--------------------------------
    frmPedidoItemVendaInc.Status = tpStatus_Incluir
    frmPedidoItemVendaInc.StatusItem = tpStatus_Incluir
    Select Case intGrupo
    Case 0
      frmPedidoItemVendaInc.TipoVenda = tpTipoVenda.tpTipoVenda_Balc
      frmPedidoItemVendaInc.intCdastro = 1
    Case 1
      frmPedidoItemVendaInc.TipoVenda = tpTipoVenda.tpTipoVenda_Clie
      frmPedidoItemVendaInc.intCdastro = 0
    Case 2
      frmPedidoItemVendaInc.TipoVenda = tpTipoVenda.tpTipoVenda_Emp
      frmPedidoItemVendaInc.intCdastro = 0
    End Select
    frmPedidoItemVendaInc.lngPEDIDOVENDAID = 0
    frmPedidoItemVendaInc.Show vbModal
  Case 3
    'Alteração de Pedido Balcão, Cliente e Empresa
    'Verifica se selecionou um pedido
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione um Pedido para alterá-lo.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    '----------------------------
    '----------------------------
    'Pede Senha Superior (Diretor, Gerente ou Administrador
    If gsNivel = "LOJ" Or gsNivel = "ADM" Then
      'Só pede senha superior se quem estiver logado não for superior
      frmLoginSup.Show vbModal
      
      If Len(Trim(gsNomeUsuLib)) = 0 Then
        strMsg = "Para efetuar a alteração da venda é necessário se autenticar no sistema."
        TratarErroPrevisto strMsg, "cmdConfirmar_Click"
        Exit Sub
      End If
      '
      'Capturou Nome do Usuário, continua processo de Venda
    Else
      MsgBox "Apenas o perfil Caixa-Vendedor pode alterar pedidos.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    '--------------------------------
    '--------------------------------
    'Verifica status
    If Trim(grdGeral.Columns("Status").Value & "") = "C" Then
      MsgBox "Pedido cancelado.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") = "X" Then
      MsgBox "Pedido excluído.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") <> "B" _
        And Trim(grdGeral.Columns("Status").Value & "") <> "L" _
        And Trim(grdGeral.Columns("Status").Value & "") <> "E" Then
      MsgBox "Pedido já possui pagamento.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    'Validação dos ítens do pedido
    If Not ValidaCamposPedido Then
      SetarFoco grdGeral
      Exit Sub
    End If
    Select Case Trim(grdGeral.Columns("Status").Value & "")
    Case "B": frmPedidoItemVendaInc.TipoVenda = tpTipoVenda.tpTipoVenda_Balc
    Case "L": frmPedidoItemVendaInc.TipoVenda = tpTipoVenda.tpTipoVenda_Clie
    Case "E": frmPedidoItemVendaInc.TipoVenda = tpTipoVenda.tpTipoVenda_Emp
    End Select
    frmPedidoItemVendaInc.Status = tpStatus_Alterar
    frmPedidoItemVendaInc.lngPEDIDOVENDAID = grdGeral.Columns("ID").Value & ""
    frmPedidoItemVendaInc.intCdastro = 0
    frmPedidoItemVendaInc.Show vbModal
  Case 4
    'Cancelamento/ATIVAÇÃO de pedido
    'Confirmação
    '--------------------------------
    '--------------------------------
    'Verifica status
    If Trim(grdGeral.Columns("Status").Value & "") = "X" Then
      MsgBox "Pedido excluído não pode ser ativado.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") <> "B" _
        And Trim(grdGeral.Columns("Status").Value & "") <> "L" _
        And Trim(grdGeral.Columns("Status").Value & "") <> "E" _
        And Trim(grdGeral.Columns("Status").Value & "") <> "C" Then
      MsgBox "Pedido já possui pagamento.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    
    If Trim(grdGeral.Columns("Status").Value & "") = "C" Then
      'Cancelado --> reativar
      If MsgBox("Confirma ativação do pedido " & Format(grdGeral.Columns("Sequencial").Value, "0000") & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
        SetarFoco grdGeral
        Exit Sub
      End If
    Else
      'Ativo --> cancelar
      If MsgBox("Confirma cancelamento do pedido " & Format(grdGeral.Columns("Sequencial").Value, "0000") & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
        SetarFoco grdGeral
        Exit Sub
      End If
    End If
    Set objPedidoVenda = New busSisMetal.clsPedidoVenda
    objPedidoVenda.AtivaInativaPedidoVenda grdGeral.Columns("ID").Value, _
                                           Trim(grdGeral.Columns("Status").Value & "")

    Set objPedidoVenda = Nothing

  End Select
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  End
End Sub


Private Sub Form_Load()
  On Error GoTo trata
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
  TratarErro Err.Number, Err.Description, "[frmGerencialPed.grdGeral_UnboundReadDataEx]"
End Sub

Private Function ValidaCamposPedido() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  Dim objGer        As busSisMetal.clsGeral
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  '
  '
  blnSetarFocoControle = True
  ValidaCamposPedido = False
  '
  On Error GoTo trata
  Set objGer = New busSisMetal.clsGeral
  'ITEM_PEDIDO
  strSql = "Select * from PEDIDOVENDA WHERE PKID = " & grdGeral.Columns("ID").Value
  Set objRs = objGer.ExecutarSQL(strSql)
  If objRs.EOF Then
    strMsg = strMsg & "Pedido não cadatrado." & vbCrLf
  End If
  If strMsg = "" Then
    If objRs.Fields("VENDEDORID").Value <> giFunIdUsuLib Then
      strMsg = strMsg & "Apenas o funcionário que cadastrou o Pedido pode alterá-lo." & vbCrLf
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set objGer = Nothing
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmGerencialPed.ValidaCamposPedido]"
    ValidaCamposPedido = False
  Else
    ValidaCamposPedido = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmGerencial.ValidaCamposPedido]", _
            Err.Description
End Function

Public Sub ConcederAcessoFnc()
  On Error GoTo trata
  Select Case gsNivel
'''  Case gsAdmin
'''    cmdSelecao(0).Enabled = True
'''  Case gsDiretor
'''    cmdSelecao(0).Enabled = True
'''  Case gsGerente
'''    cmdSelecao(0).Enabled = True
'''  Case gsCompra
'''    cmdSelecao(0).Enabled = False
  End Select
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmGerencial.ValidaCamposExclusao]", _
            Err.Description
End Sub

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
  'strDataCanc = Format(DateAdd("d", (giQtdDiasPedido * (-1)), Now), "DD/MM/YYYY hh:mm")
  strDataCanc = Format(Now, "DD/MM/YYYY")
  '
  Set objGeral = New busSisMetal.clsGeral
  '
  strSql = "SELECT PEDIDOVENDA.PKID, PEDIDOVENDA.PED_NUMERO , PEDIDOVENDA.DATA, MIN(PESSOA.NOME), CASE MIN(TIPOVENDA.DESCRICAO) WHEN 'EMPRESA' THEN MIN(LOJA.NOME) WHEN 'CLIENTE' THEN MIN(FICHACLIENTE.NOME) ELSE '' END, SUM(ITEM_PEDIDOVENDA.QUANTIDADE), SUM(ITEM_PEDIDOVENDA.VALOR), " & _
        " MIN(PEDIDOVENDA.STATUS) " & _
        "FROM PEDIDOVENDA LEFT JOIN PESSOA ON PEDIDOVENDA.VENDEDORID = PESSOA.PKID " & _
        " LEFT JOIN TIPOVENDA ON PEDIDOVENDA.TIPOVENDAID = TIPOVENDA.PKID " & _
        " LEFT JOIN FICHACLIENTE ON PEDIDOVENDA.FICHACLIENTEID = FICHACLIENTE.PKID " & _
        " LEFT JOIN LOJA ON LOJA.PKID = PEDIDOVENDA.EMPRESAID " & _
        " LEFT JOIN ITEM_PEDIDOVENDA ON PEDIDOVENDA.PKID = ITEM_PEDIDOVENDA.PEDIDOVENDAID " & _
        " WHERE PEDIDOVENDA.DATA >= " & Formata_Dados(strDataCanc, tpDados_DataHora) & _
        " AND PEDIDOVENDA.STATUS IN ('B', 'L', 'E', 'C', 'X') " & _
        " GROUP BY PEDIDOVENDA.PKID, PEDIDOVENDA.PED_NUMERO , PEDIDOVENDA.DATA " & _
        " ORDER BY PEDIDOVENDA.PED_NUMERO DESC, PEDIDOVENDA.DATA;"
  'B-BALCÃO, L-CLIENTE, E-EMPRESA,C-CANCELADO, X-EXCLUIDA
  'NÃO MOSTRA O QUE JÁ FOI RECEBIDO
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
