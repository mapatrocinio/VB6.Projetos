VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserEstoqueInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de itens no estoque"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5610
      Left            =   8985
      ScaleHeight     =   5610
      ScaleWidth      =   1860
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   90
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3390
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5295
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   180
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do item do estoque"
      TabPicture(0)   =   "userEstoqueInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraAluno"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraAluno 
         Height          =   4455
         Left            =   90
         TabIndex        =   15
         Top             =   420
         Width           =   8565
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   4185
            Index           =   0
            Left            =   90
            ScaleHeight     =   4185
            ScaleWidth      =   8265
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   240
            Width           =   8265
            Begin VB.Frame Frame3 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Informações cadastrais"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4065
               Left            =   90
               TabIndex        =   17
               Top             =   60
               Width           =   8175
               Begin VB.TextBox txtDescricao 
                  Height          =   285
                  Left            =   1200
                  MaxLength       =   50
                  TabIndex        =   1
                  Text            =   "txtDescricao"
                  Top             =   570
                  Width           =   6855
               End
               Begin VB.TextBox mskCodigo 
                  Height          =   285
                  Left            =   1200
                  MaxLength       =   50
                  TabIndex        =   0
                  Top             =   240
                  Width           =   6855
               End
               Begin VB.ComboBox cboUnidade 
                  Height          =   315
                  Left            =   1200
                  Style           =   2  'Dropdown List
                  TabIndex        =   2
                  Top             =   900
                  Width           =   3135
               End
               Begin MSMask.MaskEdBox mskValor 
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   3
                  Top             =   1260
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   450
                  _Version        =   393216
                  BackColor       =   -2147483644
                  Format          =   "#,##0.00;($#,##0.00)"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskQtdEstoque 
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   5
                  Top             =   1560
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   450
                  _Version        =   393216
                  BackColor       =   -2147483644
                  Format          =   "#,##0;(#,##0)"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskQtdMinEst 
                  Height          =   255
                  Left            =   4590
                  TabIndex        =   4
                  Top             =   1290
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   450
                  _Version        =   393216
                  Format          =   "#,##0;($#,##0)"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskValorIndenizacao 
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   6
                  Top             =   1860
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   450
                  _Version        =   393216
                  BackColor       =   -2147483644
                  Format          =   "#,##0.00;($#,##0.00)"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskPeso 
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   7
                  Top             =   2160
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   450
                  _Version        =   393216
                  BackColor       =   -2147483644
                  Format          =   "#,##0.00;($#,##0.00)"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskAltura 
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   8
                  Top             =   2460
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   450
                  _Version        =   393216
                  BackColor       =   -2147483644
                  Format          =   "#,##0.00;($#,##0.00)"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskLargura 
                  Height          =   255
                  Left            =   4620
                  TabIndex        =   9
                  Top             =   2460
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   450
                  _Version        =   393216
                  BackColor       =   -2147483644
                  Format          =   "#,##0.00;($#,##0.00)"
                  PromptChar      =   "_"
               End
               Begin VB.Label lblLargura 
                  Caption         =   "Largura"
                  Height          =   225
                  Left            =   2880
                  TabIndex        =   28
                  Top             =   2490
                  Width           =   1755
               End
               Begin VB.Label lblAltura 
                  Caption         =   "Altura"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   27
                  Top             =   2490
                  Width           =   1095
               End
               Begin VB.Label Label4 
                  Caption         =   "Peso"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   26
                  Top             =   2190
                  Width           =   1095
               End
               Begin VB.Label Label12 
                  Caption         =   "Vr Indenização"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   25
                  Top             =   1890
                  Width           =   1095
               End
               Begin VB.Label Label8 
                  Caption         =   "Valor"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   24
                  Top             =   1290
                  Width           =   1095
               End
               Begin VB.Label Label9 
                  Caption         =   "Descrição"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   23
                  Top             =   570
                  Width           =   1215
               End
               Begin VB.Label Label3 
                  Caption         =   "Código"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   22
                  Top             =   240
                  Width           =   735
               End
               Begin VB.Label Label5 
                  Caption         =   "Qtd. Estoque"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   21
                  Top             =   1590
                  Width           =   1065
               End
               Begin VB.Label Label6 
                  Caption         =   "Qtd. Min. em Estoque"
                  Height          =   255
                  Left            =   2880
                  TabIndex        =   20
                  Top             =   1290
                  Width           =   1695
               End
               Begin VB.Label Label2 
                  Caption         =   "Informar 0, caso não queira controlar este item ou -1 para aceitar estoque negativo"
                  ForeColor       =   &H000000FF&
                  Height          =   825
                  Index           =   1
                  Left            =   6240
                  TabIndex        =   19
                  Top             =   960
                  Width           =   1785
               End
               Begin VB.Label Label10 
                  Caption         =   "Unidade"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   18
                  Top             =   930
                  Width           =   975
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserEstoqueInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                 As tpStatus
Public lngESTOQUEID           As Long
Public bRetorno               As Boolean
Public bFechar                As Boolean
Public sTitulo                As String
Public intQuemChamou          As Integer
Private blnPrimeiraVez        As Boolean
Public lngQtdAnterior         As Long


Private Sub cboUnidade_Click()
  On Error GoTo trata
  '
  Select Case cboUnidade.Text & ""
  Case RectpIcUnidade.tpIcUnidade_M2
    'M2
    lblAltura.Enabled = True
    lblLargura.Enabled = True
    mskAltura.Enabled = True
    mskLargura.Enabled = True
    '
  Case RectpIcUnidade.tpIcUnidade_MLINEAR
    'Mlinear
    lblAltura.Enabled = False
    lblLargura.Enabled = True
    mskAltura.Enabled = False
    mskLargura.Enabled = True
    '
    LimparCampoMask mskAltura
  Case RectpIcUnidade.tpIcUnidade_UNID
    'Unidade
    lblAltura.Enabled = False
    lblLargura.Enabled = False
    mskAltura.Enabled = False
    mskLargura.Enabled = False
    '
    LimparCampoMask mskAltura
    LimparCampoMask mskLargura
  End Select
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "frmUserEstoqueInc.tabDetalhes"
  AmpN
End Sub

Private Sub cboUnidade_LostFocus()
  Pintar_Controle cboUnidade, tpCorContr_Normal
End Sub

Private Sub cboTipo_click()
  On Error GoTo trata
'''  Dim sCod          As String
'''  Dim sCodAux       As String
'''  Dim strSql        As String
'''  Dim objRs         As ADODB.Recordset
'''  Dim strNovocodigo As String
'''  Dim objGeral      As busSisLoc.clsGeral
'''  '
'''  Set objGeral = New busSisLoc.clsGeral
'''  If cboTipo.Text = "CARDÁPIO" Then sCod = "C"
'''  If cboTipo.Text = "CONSUMO INTERNO" Then sCod = "I"
'''  '
'''  sCodAux = sCod
'''  sCod = sCod & mskCodigo.Text
'''  txtCodigo.Text = sCod
'''  '
'''  If sCodAux = "C" Or sCodAux = "I" Then
'''    If Status = tpStatus_Incluir Then
'''      If Len(Trim(mskCodigo.Text)) = 0 Then
'''        'Seleciona último + 1
'''        strSql = "Select max(codigo) as Maximo from estoque where codigo like '" & sCodAux & "*'"
'''        Set objRs = objGeral.ExecutarSQL(strSql)
'''        If objRs.EOF Then
'''          strNovocodigo = "0001"
'''        Else
'''          If Not IsNull(objRs.Fields("Maximo").Value) Then
'''            If IsNumeric(Right(objRs.Fields("Maximo").Value, Len(objRs.Fields("Maximo").Value) - 1)) Then
'''              If CLng(Right(objRs.Fields("Maximo").Value, Len(objRs.Fields("Maximo").Value) - 1)) >= 9999 Then
'''                'MsgBox "O último item  do " & cboTipo.Text & " atingiu o limite de 9999", vbExclamation, TITULOSISTEMA
'''              Else
'''                mskCodigo.Text = Format(CLng(Right(objRs.Fields("Maximo").Value, Len(objRs.Fields("Maximo").Value) - 1)) + 1, "0000")
'''              End If
'''            End If
'''          End If
'''        End If
'''        '
'''        objRs.Close
'''        Set objRs = Nothing
'''        '
'''      End If
'''    End If
'''  End If
'''  Set objGeral = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub


Private Sub cmdCancelar_Click()
  '
  On Error GoTo trata
  '
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
  Dim strSql                    As String
  Dim objRs                     As ADODB.Recordset
  Dim clsEst                    As busSisLoc.clsEstoque
  Dim lngUNIDADEID              As Long
  '
  Dim strMsgErro                As String
  Dim objGer                    As busSisLoc.clsGeral
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração de CARDÁPIO
    If Not ValidaCampos Then Exit Sub
    Set clsEst = New busSisLoc.clsEstoque
    '
    'Obter campos
    Set objGer = New busSisLoc.clsGeral
    strSql = "SELECT PKID FROM UNIDADE WHERE UNIDADE.UNIDADE = " & _
        Formata_Dados(cboUnidade.Text, tpDados_Texto)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngUNIDADEID = objRs.Fields("PKID").Value & ""
    Else
      lngUNIDADEID = 0
    End If
    'Verifica duplicidade
    'DE CÓDIGO
    strSql = "Select * From ESTOQUE WHERE CODIGO = " & Formata_Dados(mskCodigo.Text, tpDados_Texto) & _
      " AND PKID <> " & Formata_Dados(lngESTOQUEID, tpDados_Longo)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGer = Nothing
      TratarErroPrevisto "Código já cadastrado", "cmdOK_Click"
      Pintar_Controle mskCodigo, tpCorContr_Erro
      SetarFoco mskCodigo
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    'DE DESCRICAO
    strSql = "Select * From ESTOQUE WHERE DESCRICAO = " & Formata_Dados(txtDescricao.Text, tpDados_Texto) & _
      " AND PKID <> " & Formata_Dados(lngESTOQUEID, tpDados_Longo)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGer = Nothing
      TratarErroPrevisto "Descrição já cadastrada", "cmdOK_Click"
      Pintar_Controle txtDescricao, tpCorContr_Erro
      SetarFoco txtDescricao
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGer = Nothing
    '
    '
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      clsEst.AlterarEstoque lngESTOQUEID, _
                            mskCodigo.Text, _
                            txtDescricao.Text, _
                            lngUNIDADEID & "", _
                            mskQtdEstoque.Text, _
                            mskValor.Text, _
                            mskValorIndenizacao.Text, _
                            mskPeso.Text, _
                            mskAltura.Text, _
                            mskLargura, _
                            mskQtdMinEst.Text

    ElseIf Status = tpStatus_Incluir Then
      '
      clsEst.InserirEstoque lngESTOQUEID, _
                            mskCodigo.Text, _
                            txtDescricao.Text, _
                            lngUNIDADEID & "", _
                            mskQtdEstoque.Text, _
                            mskValor.Text, _
                            mskValorIndenizacao.Text, _
                            mskPeso.Text, _
                            mskAltura.Text, _
                            mskLargura, _
                            mskQtdMinEst.Text
    End If
    Set clsEst = Nothing
    bRetorno = True
  End Select
  bFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub


Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCampos = False
  If Not Valida_String(mskCodigo, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o código" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(txtDescricao, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a descrição" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(cboUnidade, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar a unidade" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskValor, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o valor válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskQtdMinEst, TpObrigatorio, blnSetarFocoControle, True) Then
    strMsg = strMsg & "Preencher quantidade mínima válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskQtdEstoque, TpObrigatorio, blnSetarFocoControle, True) Then
    strMsg = strMsg & "Preencher quantidade válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserEstoqueInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserEstoqueInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    If Status = tpStatus_Incluir Then
      tabDetalhes.Tab = 0
      SetarFoco mskCodigo
    Else
      tabDetalhes.Tab = 0
    End If
    blnPrimeiraVez = False
    
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserEstoqueInc.Form_Activate]"
End Sub


Private Sub mskAltura_GotFocus()
  Selecionar_Conteudo mskAltura
End Sub

Private Sub mskAltura_LostFocus()
  Pintar_Controle mskAltura, tpCorContr_Normal
End Sub

Private Sub mskCodigo_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  KeyAscii = TRANSFORMA_MAIUSCULA(KeyAscii)
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub mskCodigo_LostFocus()
  On Error GoTo trata
  Pintar_Controle mskCodigo, tpCorContr_Normal
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Load()
On Error GoTo trata
  Dim objRs     As ADODB.Recordset
  Dim strSql    As String
  Dim clsEst    As busSisLoc.clsEstoque
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 5985
  Me.Width = 10935
  CenterForm Me
  blnPrimeiraVez = True
  '
  '----------------------------
  '----------------------------
  'Desabilitação dos botões para ajuste e exclusão de estoque (Diretor, Gerente ou Administrador)
  'If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
  'Else
  'mskQtdEstoque.Enabled = True
  'mskQtdEstoque.BackColor = vbWhite
  'End If
  '--------------------------------
  '--------------------------------
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  LimparCampos
  '
  strSql = "Select UNIDADE.UNIDADE from UNIDADE ORDER BY UNIDADE.UNIDADE"
  PreencheCombo cboUnidade, strSql, False
  '
  tabDetalhes_Click 0
  '
  If Status = tpStatus_Incluir Then
    mskQtdEstoque.Enabled = True
    Label5.Enabled = True
  Else
    mskQtdEstoque.Enabled = False
    Label5.Enabled = False
  End If
  
  If Status = tpStatus_Incluir Then
    '
  ElseIf Status = tpStatus_Alterar Then
    'Pega Dados do Banco de dados
    Set clsEst = New busSisLoc.clsEstoque
    Set objRs = clsEst.ListarEstoque(lngESTOQUEID)
    '
    If Not objRs.EOF Then
      mskCodigo.Text = objRs.Fields("CODIGO").Value & ""
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
      INCLUIR_VALOR_NO_COMBO objRs.Fields("DESCR_UNIDADE").Value, cboUnidade
      INCLUIR_VALOR_NO_MASK mskQtdEstoque, objRs.Fields("QUANTIDADE").Value, TpMaskLongo
      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VALOR").Value, TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskValorIndenizacao, objRs.Fields("VALORINDENIZACAO").Value, TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskPeso, objRs.Fields("PESO").Value, TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskAltura, objRs.Fields("ALTURA").Value, TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskLargura, objRs.Fields("LARGURA").Value, TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskQtdMinEst, objRs.Fields("QTDMINIMA").Value, TpMaskLongo
      '
      lngQtdAnterior = IIf(Not IsNumeric(objRs.Fields("QUANTIDADE").Value), 0, objRs.Fields("QUANTIDADE").Value)
    End If
    'cboUnidade.Enabled = False 'Não permite alteração deste campo
    objRs.Close
    Set objRs = Nothing
    Set clsEst = Nothing
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
  If Not bFechar Then Cancel = True
End Sub




Private Sub mskLargura_GotFocus()
  Selecionar_Conteudo mskLargura
End Sub

Private Sub mskLargura_LostFocus()
  Pintar_Controle mskLargura, tpCorContr_Normal
End Sub

Private Sub mskPeso_GotFocus()
  Selecionar_Conteudo mskPeso
End Sub

Private Sub mskPeso_LostFocus()
  Pintar_Controle mskPeso, tpCorContr_Normal
End Sub

Private Sub mskQtdEstoque_GotFocus()
  Selecionar_Conteudo mskQtdEstoque
End Sub

Private Sub mskQtdEstoque_LostFocus()
  Pintar_Controle mskQtdEstoque, tpCorContr_Normal
End Sub

Private Sub mskQtdMinEst_GotFocus()
  Selecionar_Conteudo mskQtdMinEst
End Sub

Private Sub mskQtdMinEst_LostFocus()
  Pintar_Controle mskQtdMinEst, tpCorContr_Normal
End Sub

Private Sub mskValor_GotFocus()
  Selecionar_Conteudo mskValor
End Sub

Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub

Private Sub mskValorIndenizacao_GotFocus()
  Selecionar_Conteudo mskValorIndenizacao
End Sub

Private Sub mskValorIndenizacao_LostFocus()
  Pintar_Controle mskValorIndenizacao, tpCorContr_Normal
End Sub

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  Dim strMsgErro    As String
  Dim strCobranca   As String
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    'dados principais da venda
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "SisLoc.frmUserEstoqueInc.tabDetalhes"
  AmpN
End Sub




Private Sub mskCodigo_GotFocus()
  Selecionar_Conteudo mskCodigo
End Sub

Private Sub txtDescricao_GotFocus()
  Selecionar_Conteudo txtDescricao
End Sub

Private Sub txtDescricao_LostFocus()
  Pintar_Controle txtDescricao, tpCorContr_Normal
End Sub


Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Estoque
  LimparCampoTexto mskCodigo
  LimparCampoTexto txtDescricao
  LimparCampoCombo cboUnidade
  LimparCampoMask mskValor
  LimparCampoMask mskQtdMinEst
  LimparCampoMask mskQtdEstoque
  LimparCampoMask mskValorIndenizacao
  LimparCampoMask mskPeso
  LimparCampoMask mskAltura
  LimparCampoMask mskLargura
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserEmpresaInc.LimparCampos]", _
            Err.Description
End Sub


