VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPedidoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Pedido"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5610
      Left            =   8520
      ScaleHeight     =   5610
      ScaleWidth      =   1860
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4695
         Left            =   90
         ScaleHeight     =   4635
         ScaleWidth      =   1605
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   720
         Width           =   1665
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&X"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3600
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2730
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5295
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do Pedido"
      TabPicture(0)   =   "userPedidoInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraProf"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Itens do pedido"
      TabPicture(1)   =   "userPedidoInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdPedido"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "picTrava(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.PictureBox picTrava 
         BorderStyle     =   0  'None
         Height          =   645
         Index           =   1
         Left            =   -74910
         ScaleHeight     =   645
         ScaleWidth      =   8025
         TabIndex        =   24
         Top             =   4530
         Width           =   8025
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Left            =   6300
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   "txtValor"
            Top             =   300
            Width           =   1695
         End
         Begin VB.TextBox txtPeso 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Left            =   4590
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Text            =   "txtPeso"
            Top             =   300
            Width           =   1695
         End
         Begin VB.TextBox txtQuantidade 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Text            =   "txtQuantidade"
            Top             =   300
            Width           =   1695
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Valor"
            Height          =   255
            Left            =   6330
            TabIndex        =   27
            Top             =   30
            Width           =   1665
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Peso"
            Height          =   255
            Left            =   4620
            TabIndex        =   26
            Top             =   30
            Width           =   1665
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Quantidade"
            Height          =   255
            Left            =   2910
            TabIndex        =   25
            Top             =   30
            Width           =   1665
         End
      End
      Begin VB.Frame fraProf 
         Height          =   3972
         Left            =   120
         TabIndex        =   15
         Top             =   330
         Width           =   7935
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   3612
            Index           =   0
            Left            =   120
            ScaleHeight     =   3615
            ScaleWidth      =   7695
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   240
            Width           =   7695
            Begin VB.ComboBox cboFabrica 
               Height          =   315
               Left            =   1230
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   1740
               Width           =   4515
            End
            Begin VB.ComboBox cboAnodizadora 
               Height          =   315
               Left            =   1230
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   1380
               Width           =   4515
            End
            Begin VB.ComboBox cboFornecedor 
               Height          =   315
               Left            =   1230
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   720
               Width           =   4515
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   255
               Left            =   30
               ScaleHeight     =   255
               ScaleWidth      =   3855
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   390
               Width           =   3855
               Begin MSMask.MaskEdBox mskData 
                  Height          =   255
                  Index           =   0
                  Left            =   1200
                  TabIndex        =   1
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  _Version        =   393216
                  BackColor       =   14737632
                  AutoTab         =   -1  'True
                  MaxLength       =   16
                  Mask            =   "##/##/#### ##:##"
                  PromptChar      =   "_"
               End
               Begin VB.Label Label2 
                  Caption         =   "Data"
                  Height          =   255
                  Index           =   1
                  Left            =   0
                  TabIndex        =   18
                  Top             =   0
                  Width           =   615
               End
            End
            Begin VB.TextBox txtNumeroOS 
               BackColor       =   &H00E0E0E0&
               Height          =   288
               Left            =   1230
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtNumeroOS"
               Top             =   30
               Width           =   1815
            End
            Begin MSMask.MaskEdBox mskValor 
               Height          =   255
               Left            =   1230
               TabIndex        =   3
               Top             =   1080
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Fábrica"
               Height          =   195
               Index           =   2
               Left            =   30
               TabIndex        =   23
               Top             =   1770
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Anodizadora"
               Height          =   195
               Index           =   1
               Left            =   30
               TabIndex        =   22
               Top             =   1410
               Width           =   1215
            End
            Begin VB.Label Label4 
               Caption         =   "Valor Alumínio"
               Height          =   255
               Left            =   30
               TabIndex        =   21
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label Label5 
               Caption         =   "Fornecedor"
               Height          =   195
               Index           =   0
               Left            =   30
               TabIndex        =   20
               Top             =   750
               Width           =   1215
            End
            Begin VB.Label Label2 
               Caption         =   "Ano-Número OS"
               Height          =   255
               Index           =   0
               Left            =   30
               TabIndex        =   19
               Top             =   30
               Width           =   1155
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdPedido 
         Height          =   4005
         Left            =   -74910
         OleObjectBlob   =   "userPedidoInc.frx":0038
         TabIndex        =   6
         Top             =   420
         Width           =   7995
      End
   End
End
Attribute VB_Name = "frmPedidoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                 As tpStatus
Public lngPEDIDOID            As Long
Public blnRetorno             As Boolean
Public blnFechar              As Boolean
Private blnPrimeiraVez        As Boolean

Dim PED_COLUNASMATRIZ         As Long
Dim PED_LINHASMATRIZ          As Long
Private PED_Matriz()          As String


Private Sub cboAnodizadora_LostFocus()
  Pintar_Controle cboAnodizadora, tpCorContr_Normal
End Sub

Private Sub cboFabrica_LostFocus()
  Pintar_Controle cboFabrica, tpCorContr_Normal
End Sub

Private Sub cboFornecedor_LostFocus()
  On Error GoTo trata
  Dim objLoja         As busSisMetal.clsLoja
  Dim objRs           As ADODB.Recordset
  If Me.ActiveControl.Name = "cmdCancelar" Then Exit Sub

  Pintar_Controle cboFornecedor, tpCorContr_Normal
  If Len(cboFornecedor.Text) = 0 Then
    Exit Sub
  End If
  Set objLoja = New busSisMetal.clsLoja
  '
  Set objRs = objLoja.SelecionarFornecedorPeloNome(cboFornecedor.Text)
  If Not objRs.EOF Then
    INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VALOR_KG").Value, TpMaskMoeda
  End If
  '
  objRs.Close
  Set objRs = Nothing
  Set objLoja = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdAlterar_Click()
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 1 'Itens do pedido
    If Len(Trim(grdPedido.Columns("PKID").Value & "")) = 0 Then
      MsgBox "Selecione um item do pedido!", vbExclamation, TITULOSISTEMA
      Exit Sub
    End If
    frmItemPedidoInc.Status = tpStatus_Alterar
    frmItemPedidoInc.lngPKID = grdPedido.Columns("PKID").Value
    frmItemPedidoInc.lngPEDIDOID = lngPEDIDOID
    frmItemPedidoInc.strFornecedor = cboFornecedor.Text
    frmItemPedidoInc.strNumero = txtNumeroOS.Text
    frmItemPedidoInc.strData = mskData(0).Text
    frmItemPedidoInc.Show vbModal

    If frmItemPedidoInc.blnRetorno Then
      PED_MontaMatriz
      grdPedido.Bookmark = Null
      grdPedido.ReBind
      grdPedido.ApproxCount = PED_LINHASMATRIZ
    End If
    SetarFoco grdPedido
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub cmdExcluir_Click()
  Dim objItemPedido As busSisMetal.clsItemPedido
  Dim objGer        As busSisMetal.clsGeral
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  '
  On Error GoTo trata
  If Len(Trim(grdPedido.Columns("PKID").Value & "")) = 0 Then
    MsgBox "Selecione um item do pedido para exclusão.", vbExclamation, TITULOSISTEMA
    Exit Sub
  End If
  '
  Set objGer = New busSisMetal.clsGeral
  'ITEM_PEDIDO
  strSql = "Select * from BAIXA_PEDIDO_OS WHERE ITEM_PEDIDOID = " & grdPedido.Columns("PKID").Value
  Set objRs = objGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    objRs.Close
    Set objRs = Nothing
    Set objGer = Nothing
    TratarErroPrevisto "Item do pedido não pode ser excluido pois já possui baixas na OS.", "frmPedidoLis.cmdExcluir_Click"
    SetarFoco grdPedido
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set objGer = Nothing
  '
  '
  If MsgBox("Confirma exclusão do item do pedido " & grdPedido.Columns("Linha-perfil").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then Exit Sub
  'OK
  Set objItemPedido = New busSisMetal.clsItemPedido

  objItemPedido.ExcluirItemPedido CLng(grdPedido.Columns("PKID").Value)
  '
  PED_MontaMatriz
  grdPedido.Bookmark = Null
  grdPedido.ReBind

  Set objItemPedido = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdIncluir_Click()
  On Error GoTo trata
  Dim objForm As Form
  '
  Select Case tabDetalhes.Tab
  Case 1 'Itens do pedido
    frmItemPedidoInc.Status = tpStatus_Incluir
    frmItemPedidoInc.lngPKID = 0
    frmItemPedidoInc.lngPEDIDOID = lngPEDIDOID
    frmItemPedidoInc.strFornecedor = cboFornecedor.Text
    frmItemPedidoInc.strNumero = txtNumeroOS.Text
    frmItemPedidoInc.strData = mskData(0).Text
    frmItemPedidoInc.Show vbModal

    If frmItemPedidoInc.blnRetorno Then
      PED_MontaMatriz
      grdPedido.Bookmark = Null
      grdPedido.ReBind
      grdPedido.ApproxCount = PED_LINHASMATRIZ
    End If
    SetarFoco grdPedido
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim objPedido               As busSisMetal.clsPedido
  Dim objRs                   As ADODB.Recordset
  Dim objGer                  As busSisMetal.clsGeral
  Dim lngFORNECEDORID         As Long
  Dim lngANODIZADORAID        As Long
  Dim lngFABRICAID            As Long
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração do pedido
    If Not ValidaCampos Then Exit Sub
'''    'Valida se pedido já cadastrada
'''    Set objGer = New busSisMetal.clsGeral
'''    strSql = "Select * From LINHA WHERE (NOME = " & Formata_Dados(txtNome.Text, tpDados_Texto, tpNulo_Aceita) & _
'''      " OR CODIGO = " & Formata_Dados(txtCodigo.Text, tpDados_Texto, tpNulo_Aceita) & ") " & _
'''      " AND PKID <> " & Formata_Dados(lngLINHAID, tpDados_Longo, tpNulo_NaoAceita)
'''    Set objRs = objGer.ExecutarSQL(strSql)
'''    If Not objRs.EOF Then
'''      objRs.Close
'''      Set objRs = Nothing
'''      Set objGer = Nothing
'''      TratarErroPrevisto "Nome ou Código da linha já cadastrada", "cmdOK_Click"
'''      Pintar_Controle txtNome, tpCorContr_Erro
'''      Pintar_Controle txtCodigo, tpCorContr_Erro
'''      SetarFoco txtNome
'''      Exit Sub
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
'''    '
    Set objGer = New busSisMetal.clsGeral
    'FORNECEDOR
    lngFORNECEDORID = 0
    strSql = "SELECT LOJA.PKID FROM LOJA " & _
      " INNER JOIN FORNECEDOR ON FORNECEDOR.LOJAID = LOJA.PKID " & _
      " WHERE LOJA.NOME = " & Formata_Dados(cboFornecedor.Text, tpDados_Texto)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngFORNECEDORID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    '
    'ANODIZADORA
    lngANODIZADORAID = 0
    strSql = "SELECT LOJA.PKID FROM LOJA " & _
      " INNER JOIN ANODIZADORA ON ANODIZADORA.LOJAID = LOJA.PKID " & _
      " WHERE LOJA.NOME = " & Formata_Dados(cboAnodizadora.Text, tpDados_Texto)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngANODIZADORAID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    '
    'FABRICA
    lngFABRICAID = 0
    strSql = "SELECT LOJA.PKID FROM LOJA " & _
      " INNER JOIN FABRICA ON FABRICA.LOJAID = LOJA.PKID " & _
      " WHERE LOJA.NOME = " & Formata_Dados(cboFabrica.Text, tpDados_Texto)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngFABRICAID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objGer = Nothing
    '
    Set objPedido = New busSisMetal.clsPedido
    'Altera ou incluiu pedido
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objPedido.AlterarPedido lngPEDIDOID, _
                              lngFORNECEDORID, _
                              lngANODIZADORAID, _
                              lngFABRICAID, _
                              IIf(Len(mskValor.ClipText) = 0, "", mskValor.Text)
      'Set objPedido = Nothing
      '
      blnRetorno = True
      Status = tpStatus_Alterar
      'Reload na tela
      Form_Load
      'Acerta tabs
      'tabDetalhes.TabVisible(2) = True
      tabDetalhes.Tab = 1
      cmdIncluir_Click
      
    ElseIf Status = tpStatus_Incluir Then
      'Código para inclusão
      '
      objPedido.InserirPedido lngPEDIDOID, _
                              lngFORNECEDORID, _
                              lngANODIZADORAID, _
                              lngFABRICAID, _
                              IIf(Len(mskValor.ClipText) = 0, "", mskValor.Text)
      'Set objPedido = Nothing
      '
      blnRetorno = True
      Status = tpStatus_Alterar
      'Reload na tela
      Form_Load
      'Acerta tabs
      'tabDetalhes.TabVisible(2) = True
      tabDetalhes.Tab = 1
      cmdIncluir_Click
    End If
    Set objPedido = Nothing
    'blnFechar = True
    'Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  Dim objPedido     As busSisMetal.clsPedido
  '
  blnFechar = False
  blnRetorno = False
  AmpS
  Me.Height = 5985
  Me.Width = 10470
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, cmdExcluir, , cmdIncluir, cmdAlterar
  '
  LimparCampos
  tabDetalhes_Click 0
  '
  'Fornecedor
  strSql = "SELECT LOJA.NOME FROM LOJA " & _
      " INNER JOIN FORNECEDOR ON LOJA.PKID = FORNECEDOR.LOJAID " & _
      " WHERE LOJA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
      "ORDER BY LOJA.NOME"
  PreencheCombo cboFornecedor, strSql, False, True
  'Anodizadora
  strSql = "SELECT LOJA.NOME FROM LOJA " & _
      " INNER JOIN ANODIZADORA ON LOJA.PKID = ANODIZADORA.LOJAID " & _
      " WHERE LOJA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
      "ORDER BY LOJA.NOME"
  PreencheCombo cboAnodizadora, strSql, False, True
  'Fabrica
  strSql = "SELECT LOJA.NOME FROM LOJA " & _
      " INNER JOIN FABRICA ON LOJA.PKID = FABRICA.LOJAID " & _
      " WHERE LOJA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
      "ORDER BY LOJA.NOME"
  PreencheCombo cboFabrica, strSql, False, True
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    '
    Me.Caption = "Cadastro de Pedido"
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = False
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objPedido = New busSisMetal.clsPedido
    Set objRs = objPedido.ListarPedido(lngPEDIDOID)
    '
    If Not objRs.EOF Then
      'Campos fixos
      Me.Caption = "Cadastro de Pedido [" & objRs.Fields("OS_ANO").Value & "-" & Format(objRs.Fields("OS_NUMERO").Value & "", "0000") & "]"
      txtNumeroOS = objRs.Fields("OS_ANO").Value & "-" & Format(objRs.Fields("OS_NUMERO").Value & "", "0000")
      INCLUIR_VALOR_NO_MASK mskData(0), objRs.Fields("DATA").Value, TpMaskData
      'Campos inserts
      INCLUIR_VALOR_NO_COMBO objRs.Fields("NOME_FORNECEDOR").Value & "", cboFornecedor
      INCLUIR_VALOR_NO_COMBO objRs.Fields("NOME_ANODIZADORA").Value & "", cboAnodizadora
      INCLUIR_VALOR_NO_COMBO objRs.Fields("NOME_FABRICA").Value & "", cboFabrica
      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VALOR_ALUMINIO").Value, TpMaskMoeda
    End If
    Set objPedido = Nothing
    '
    If Status = tpStatus_Alterar Then
      tabDetalhes.TabEnabled(0) = True
      tabDetalhes.TabEnabled(1) = True
    ElseIf Status = tpStatus_Consultar Then
      tabDetalhes.TabEnabled(0) = False
      tabDetalhes.TabEnabled(1) = True
      tabDetalhes.Tab = 1
      tabDetalhes_Click 1
    End If
    '
    
  End If
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
End Sub

Private Sub cmdCancelar_Click()
  blnFechar = True
  blnRetorno = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub


Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Pedido
  LimparCampoTexto txtNumeroOS
  LimparCampoMask mskData(0)
  LimparCampoCombo cboFornecedor
  LimparCampoMask mskValor
  LimparCampoCombo cboAnodizadora
  LimparCampoCombo cboFabrica
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmPedidoInc.LimparCampos]", _
            Err.Description
End Sub

Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  '
  If Not Valida_String(cboFornecedor, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o fornecedor" & vbCrLf
  End If
  If Not Valida_Moeda(mskValor, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Valor do Alumínio inválido" & vbCrLf
  End If
  If Not Valida_String(cboAnodizadora, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar a anodizadora" & vbCrLf
  End If
  If Not Valida_String(cboFabrica, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar a fábrica" & vbCrLf
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmPedidoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[frmPedidoInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Setar foco
    SetarFoco cboFornecedor
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmPedidoInc.Form_Activate]"
End Sub

Public Sub PED_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGeral  As busSisMetal.clsGeral
  '
  Dim curQuantidade   As Currency
  Dim curPeso         As Currency
  Dim curValor        As Currency
  '
  On Error GoTo trata
  
  Set objGeral = New busSisMetal.clsGeral
  '
  strSql = "SELECT ITEM_PEDIDO.PKID, " & _
          "TIPO_LINHA.NOME + ' - ' + LINHA.CODIGO, ITEM_PEDIDO.QUANTIDADE, ITEM_PEDIDO.PESO, ISNULL(VALOR_ALUMINIO, 0) * ISNULL(ITEM_PEDIDO.PESO, 0) AS VALOR, " & _
          "ITEM_PEDIDO.COMPRIMENTO_VARA " & _
          "FROM ITEM_PEDIDO " & _
          " INNER JOIN PEDIDO ON PEDIDO.PKID = ITEM_PEDIDO.PEDIDOID " & _
          " LEFT JOIN LINHA ON LINHA.PKID = ITEM_PEDIDO.LINHAID " & _
          " LEFT JOIN TIPO_LINHA ON TIPO_LINHA.PKID = LINHA.TIPO_LINHAID " & _
          " LEFT JOIN LOJA AS FORNECEDOR ON FORNECEDOR.PKID = PEDIDO.FORNECEDORID " & _
          "WHERE ITEM_PEDIDO.PEDIDOID = " & Formata_Dados(lngPEDIDOID, tpDados_Longo) & _
          " ORDER BY TIPO_LINHA.NOME, LINHA.CODIGO"
  '
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    PED_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim PED_Matriz(0 To PED_COLUNASMATRIZ - 1, 0 To PED_LINHASMATRIZ - 1)
  Else
    ReDim PED_Matriz(0 To PED_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  curQuantidade = 0
  curPeso = 0
  curValor = 0
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To PED_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To PED_COLUNASMATRIZ - 1  'varre as colunas
          PED_Matriz(intJ, intI) = objRs(intJ) & ""
          Select Case intJ
          Case 2: curQuantidade = curQuantidade + IIf(IsNull(objRs(intJ)), 0, objRs(intJ))
          Case 3: curPeso = curPeso + IIf(IsNull(objRs(intJ)), 0, objRs(intJ))
          Case 4: curValor = curValor + IIf(IsNull(objRs(intJ)), 0, objRs(intJ))
          End Select
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  '
  txtQuantidade.Text = Format(curQuantidade, "###,##0")
  txtPeso.Text = Format(curPeso, "###,##0.000")
  txtValor.Text = Format(curValor, "###,##0.00")
  '
  Set objGeral = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then Cancel = True
End Sub

Private Sub grdPedido_UnboundReadDataEx( _
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
               Offset + intI, PED_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, PED_COLUNASMATRIZ, PED_LINHASMATRIZ, PED_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, PED_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmPedidoInc.grdGeral_UnboundReadDataEx]"
End Sub

Private Sub mskValor_GotFocus()
  Seleciona_Conteudo_Controle mskValor
End Sub
Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    grdPedido.Enabled = False
    picTrava(0).Enabled = True
    picTrava(1).Enabled = False
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    SetarFoco cboFornecedor
  Case 1
    'Itens pedido
    grdPedido.Enabled = True
    picTrava(0).Enabled = False
    picTrava(1).Enabled = True
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = True
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = True
    'Montar RecordSet
    PED_COLUNASMATRIZ = grdPedido.Columns.Count
    PED_LINHASMATRIZ = 0
    PED_MontaMatriz
    grdPedido.Bookmark = Null
    grdPedido.ReBind
    grdPedido.ApproxCount = PED_LINHASMATRIZ
    '
    SetarFoco grdPedido
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "SisMetal.frmPedidoInc.tabDetalhes"
  AmpN
End Sub

