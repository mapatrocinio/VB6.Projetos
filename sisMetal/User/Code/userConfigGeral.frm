VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConfigGeral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configura��es do Sistema - M�dulo Geral"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5865
      Left            =   8520
      ScaleHeight     =   5865
      ScaleWidth      =   1860
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   90
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3660
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5595
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9869
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Cart�o Promocional"
      TabPicture(0)   =   "userConfigGeral.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraConfiguracao(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraConfiguracao 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Configura��o"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Index           =   0
         Left            =   150
         TabIndex        =   7
         Top             =   390
         Width           =   7695
         Begin VB.TextBox txtCaminho 
            Height          =   288
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   0
            Top             =   510
            Width           =   5985
         End
         Begin MSMask.MaskEdBox mskQtdDiasVenda 
            Height          =   255
            Left            =   1320
            TabIndex        =   1
            Top             =   1260
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   2
            Format          =   "#,##0;($#,##0)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Qtd. Dias Pedido exibido na venda"
            Height          =   735
            Index           =   6
            Left            =   180
            TabIndex        =   9
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Caminho da Imagem de Compra"
            Height          =   705
            Index           =   5
            Left            =   60
            TabIndex        =   8
            Top             =   330
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmConfigGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                     As tpStatus
Public lngCONFIGID          As Long
Public bRetorno                   As Boolean
Public bFechar                    As Boolean
Private blnPrimeiraVez            As Boolean



Private Sub cmdCancelar_Click()
  bFechar = True
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdOK_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim strMsgErro              As String
  Dim objRs                   As ADODB.Recordset
  Dim objConfiguracao         As busSisMetal.clsConfiguracao
  Dim objGer                  As busSisMetal.clsGeral

  '
  If Not ValidaCampos Then Exit Sub
  '
  Set objConfiguracao = New busSisMetal.clsConfiguracao
  '
  If Status = tpStatus_Alterar Then
    'C�digo para altera��o
    '
    objConfiguracao.AlterarConfiguracaoGeral lngCONFIGID, _
                                             txtCaminho.Text, _
                                             mskQtdDiasVenda.Text

    '
    Captura_Config
    '
    bRetorno = True
  ElseIf Status = tpStatus_Incluir Then
  End If
  Set objConfiguracao = Nothing
  bFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg          As String
  Dim blnSetarFoco    As Boolean
  '
  strMsg = ""
  blnSetarFoco = True
  If Not Valida_Moeda(mskQtdDiasVenda, TpObrigatorio, blnSetarFoco) Then
    tabDetalhes.Tab = 1
    SetarFoco mskQtdDiasVenda
    strMsg = strMsg & "Preencher o Campo quantidade de dias para pedido ser exibido na venda v�lido" & vbCrLf
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserConfigGeral.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    tabDetalhes.Tab = 0
    tabDetalhes_Click 0
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserConfigCortesia.Form_Activate]"
End Sub

Private Sub LimparCampos()
  On Error GoTo trata
  'Configura��o de Cortesia/promo��es
  LimparCampoTexto txtCaminho
  LimparCampoMask mskQtdDiasVenda
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserConfigImpressao.LimparCampos]", _
            Err.Description
            
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs             As ADODB.Recordset
  Dim strSql            As String
  Dim objGeral          As busSisMetal.clsGeral
  Dim objConfiguracao   As busSisMetal.clsConfiguracao
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 6345
  Me.Width = 10470
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  'Capturar configura��es do sistema
  Set objGeral = New busSisMetal.clsGeral
  'strSql = "SELECT PKID FROM CONFIGURACAO"
  'Set objRs = objGeral.ExecutarSQL(strSql)
  'If objRs.EOF Then
  '  'Inclus�o
  '  Err.Raise 999, , "N�o h� registro de configura��o cadastrado!"
  'Else
  '  'Altera��o
  '  Status = tpStatus.tpStatus_Alterar
  '  lngCONFIGID = objRs.Fields("PKID").Value
  'End If
  'objRs.Close
  'Set objRs = Nothing
  'Set objGeral = Nothing
  Status = tpStatus.tpStatus_Alterar
  'Limpar Campos
  LimparCampos
  If Status = tpStatus_Incluir Then
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objConfiguracao = New busSisMetal.clsConfiguracao
    Set objRs = objConfiguracao.ListarConfiguracaoGeral(lngCONFIGID)
    '
    If Not objRs.EOF Then
      txtCaminho.Text = objRs.Fields("CAMINHOIMAGEMCOMPRA").Value & ""
    End If
    INCLUIR_VALOR_NO_MASK mskQtdDiasVenda, _
                          objRs.Fields("QTDDIASVENDA").Value, _
                          TpMaskMoeda
    objRs.Close
    Set objRs = Nothing
    Set objConfiguracao = Nothing
  End If
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not bFechar Then Cancel = True
End Sub

Private Sub mskQtdDiasVenda_GotFocus()
  Seleciona_Conteudo_Controle mskQtdDiasVenda
End Sub
Private Sub mskQtdDiasVenda_LostFocus()
  Pintar_Controle mskQtdDiasVenda, tpCorContr_Normal
End Sub

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    'Desabilitar campos
    fraConfiguracao(0).Enabled = True
    SetarFoco txtCaminho
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtCaminho_GotFocus()
  Seleciona_Conteudo_Controle txtCaminho
End Sub
Private Sub txtCaminho_LostFocus()
  Pintar_Controle txtCaminho, tpCorContr_Normal
End Sub

