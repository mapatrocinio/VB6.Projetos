VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPedidoItemVendaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Pedidos de loja"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   6975
      Left            =   120
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Controle de pedidos"
      TabPicture(0)   =   "userPedidoItemVendaInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdItemPedido"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraPedido"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraItemPedido"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame fraItemPedido 
         Caption         =   "Itens do Pedido"
         Height          =   1185
         Left            =   90
         TabIndex        =   20
         Top             =   1890
         Width           =   9195
         Begin VB.TextBox txtProduto 
            Height          =   285
            Left            =   1290
            MaxLength       =   50
            TabIndex        =   8
            Text            =   "txtProduto"
            Top             =   180
            Width           =   5865
         End
         Begin VB.TextBox txtNomProdutoFim 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   3660
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "txtNomProdutoFim"
            Top             =   540
            Width           =   3495
         End
         Begin VB.TextBox txtCodProdutoFim 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   1290
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   9
            TabStop         =   0   'False
            Text            =   "txtCodProdutoFim"
            Top             =   540
            Width           =   2355
         End
         Begin MSMask.MaskEdBox mskQuantidade 
            Height          =   255
            Left            =   1290
            TabIndex        =   11
            Top             =   870
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Produto"
            Height          =   435
            Index           =   4
            Left            =   180
            TabIndex        =   28
            Top             =   210
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Quantidade"
            Height          =   225
            Left            =   180
            TabIndex        =   27
            Top             =   840
            Width           =   1005
         End
      End
      Begin VB.Frame fraPedido 
         Caption         =   "Pedido"
         Height          =   1575
         Left            =   90
         TabIndex        =   19
         Top             =   330
         Width           =   9195
         Begin VB.OptionButton optDesconto 
            Caption         =   "&Percentual"
            Height          =   195
            Index           =   0
            Left            =   3120
            TabIndex        =   6
            Top             =   1260
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.OptionButton optDesconto 
            Caption         =   "&Valor"
            Height          =   195
            Index           =   1
            Left            =   4530
            TabIndex        =   7
            Top             =   1260
            Width           =   675
         End
         Begin VB.TextBox txtVendedor 
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Left            =   7140
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   2
            TabStop         =   0   'False
            Text            =   "txtVendedor"
            Top             =   180
            Width           =   1695
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   3450
            ScaleHeight     =   255
            ScaleWidth      =   2535
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   210
            Width           =   2535
            Begin MSMask.MaskEdBox mskData 
               Height          =   255
               Index           =   0
               Left            =   750
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
               TabIndex        =   24
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.TextBox txtSequencial 
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Left            =   1290
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   0
            TabStop         =   0   'False
            Text            =   "txtSequencial"
            Top             =   180
            Width           =   1815
         End
         Begin VB.TextBox txtNomeClieFornFim 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   2430
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   4
            TabStop         =   0   'False
            Text            =   "txtNomeClieFornFim"
            Top             =   510
            Width           =   6405
         End
         Begin VB.TextBox txtCodClieFornFim 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   1290
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   3
            TabStop         =   0   'False
            Text            =   "txtCodClieFornFim"
            Top             =   510
            Width           =   1125
         End
         Begin MSMask.MaskEdBox mskDesconto 
            Height          =   255
            Left            =   1290
            TabIndex        =   5
            Top             =   1170
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label lblSelecionarClieForn 
            Caption         =   "Ctrl+S Para selecionar um novo"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1290
            TabIndex        =   29
            Top             =   810
            Width           =   3465
         End
         Begin VB.Label Label2 
            Caption         =   "Vendedor"
            Height          =   255
            Index           =   2
            Left            =   6030
            TabIndex        =   26
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Sequencial"
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   25
            Top             =   180
            Width           =   1155
         End
         Begin VB.Label Label3 
            Caption         =   "Desconto"
            Height          =   195
            Left            =   90
            TabIndex        =   22
            Top             =   1230
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente"
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   21
            Top             =   510
            Width           =   1095
         End
      End
      Begin TrueDBGrid60.TDBGrid grdItemPedido 
         Height          =   3225
         Left            =   90
         OleObjectBlob   =   "userPedidoItemVendaInc.frx":001C
         TabIndex        =   12
         Top             =   3120
         Width           =   9195
      End
   End
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7245
      Left            =   9660
      ScaleHeight     =   7245
      ScaleWidth      =   1860
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2985
         Left            =   120
         ScaleHeight     =   2925
         ScaleWidth      =   1545
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   4110
         Width           =   1605
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   150
            Width           =   1305
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   885
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1020
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   885
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   1890
            Width           =   1335
         End
      End
      Begin Crystal.CrystalReport Report1 
         Left            =   270
         Top             =   3240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
   End
End
Attribute VB_Name = "frmPedidoItemVendaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public StatusItem               As tpStatus
Public lngPEDIDOVENDAID         As Long
Public lngITEMPEDIDOVENDAID     As Long
Public TipoVenda                As tpTipoVenda

Public intCdastro               As Integer
'intOrigem = 0 cadastro de pedido
'intOrigem = 1 cadastro de item do pedido

Dim blnFechar                   As Boolean
Public blnRetorno               As Boolean
Public blnPrimeiraVez           As Boolean
'
Dim ITEMPED_COLUNASMATRIZ        As Long
Dim ITEMPED_LINHASMATRIZ         As Long
Private ITEMPED_Matriz()         As String


Public Sub ITEMPED_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGeral  As busSisMetal.clsGeral
  '
  On Error GoTo trata

  Set objGeral = New busSisMetal.clsGeral
  '
  strSql = "SELECT ITEM_PEDIDOVENDA.PKID, INSUMO.CODIGO, PRODUTO.NOME, " & _
            " ITEM_PEDIDOVENDA.QUANTIDADE, ITEM_PEDIDOVENDA.VALOR, " & _
            " ITEM_PEDIDOVENDA.VALOR_INSTALACAO, ITEM_PEDIDOVENDA.VALOR_FRETE " & _
            " FROM ITEM_PEDIDOVENDA " & _
            " INNER JOIN PRODUTO ON PRODUTO.INSUMOID = ITEM_PEDIDOVENDA.PRODUTOID " & _
            " INNER JOIN INSUMO ON INSUMO.PKID = PRODUTO.INSUMOID "
  strSql = strSql & " WHERE ITEM_PEDIDOVENDA.PEDIDOVENDAID = " & Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo)

  strSql = strSql & " ORDER BY ITEM_PEDIDOVENDA.PKID DESC"
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    ITEMPED_LINHASMATRIZ = objRs.RecordCount
  Else
    ITEMPED_LINHASMATRIZ = 0
  End If
  If Not objRs.EOF Then
    ReDim ITEMPED_Matriz(0 To ITEMPED_COLUNASMATRIZ - 1, 0 To ITEMPED_LINHASMATRIZ - 1)
  Else
    ReDim ITEMPED_Matriz(0 To ITEMPED_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To ITEMPED_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To ITEMPED_COLUNASMATRIZ - 1  'varre as colunas
          ITEMPED_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set objGeral = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdAlterar_Click()
  On Error GoTo trata
  
  Select Case intCdastro
  Case 1 'Itens do Pedido
    If Len(Trim(grdItemPedido.Columns("PKID").Value & "")) = 0 Then
      MsgBox "Selecione um item do pedido!", vbExclamation, TITULOSISTEMA
      SetarFoco grdItemPedido
      Exit Sub
    End If
    'Limpar campos
    'ITEM PEDIDO
    txtProduto.Text = grdItemPedido.Columns("Produto").Value & ""
    txtCodProdutoFim.Text = grdItemPedido.Columns("CODIGO").Value & ""
    txtNomProdutoFim.Text = grdItemPedido.Columns("Produto").Value & ""
    INCLUIR_VALOR_NO_MASK mskQuantidade, grdItemPedido.Columns("Qtd.").Value & "", TpMaskLongo
    '
    StatusItem = tpStatus_Alterar
    lngITEMPEDIDOVENDAID = grdItemPedido.Columns("PKID").Value & ""
    cmdAlterar.Enabled = False
    '
    SetarFoco txtProduto
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim objPedidoVenda          As busSisMetal.clsPedidoVenda
  Dim objItemPedidoVenda      As busSisMetal.clsItemPedidoVenda
  Dim objRs                   As ADODB.Recordset
  Dim objGeral                As busSisMetal.clsGeral
  Dim strStatus               As String
  Dim strTipoVenda            As String
  Dim strTipoDesconto         As String
  Dim lngFICHACLIENTEID       As Long
  Dim lngTIPOVENDAID          As Long
  Dim lngEMPRESAID            As Long
  '
  Dim lngINSUMOID             As Long
  '
  Select Case intCdastro
  Case 0 'Gravar Pedido
    If Not ValidaCampos Then Exit Sub
    'OK procede com o cadastro
    'CADASTRO DE PEDIDO
    '-------------------------
    '
    Set objGeral = New busSisMetal.clsGeral
    '
    strTipoDesconto = ""
    If optDesconto(0).Value Then
      strTipoDesconto = "P"
    ElseIf optDesconto(1).Value Then
      strTipoDesconto = "V"
    End If
    Select Case TipoVenda
    Case tpTipoVenda.tpTipoVenda_Balc
      strStatus = "B"
      strTipoVenda = "BALCÃO"
      lngFICHACLIENTEID = 0
      lngEMPRESAID = 0
    Case tpTipoVenda.tpTipoVenda_Clie
      strStatus = "L"
      strTipoVenda = "CLIENTE"
      lngFICHACLIENTEID = txtCodClieFornFim.Text
      lngEMPRESAID = 0
      '
    Case tpTipoVenda.tpTipoVenda_Emp
      strStatus = "E"
      strTipoVenda = "EMPRESA"
      lngFICHACLIENTEID = 0
      lngEMPRESAID = txtCodClieFornFim.Text
    Case Else
      strStatus = ""
      strTipoVenda = ""
      lngFICHACLIENTEID = 0
      lngEMPRESAID = 0
    End Select
    '
    lngTIPOVENDAID = 0
    strSql = "SELECT TIPOVENDA.PKID FROM TIPOVENDA " & _
      " WHERE TIPOVENDA.DESCRICAO = " & Formata_Dados(strTipoVenda, tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngTIPOVENDAID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    '
    Set objPedidoVenda = New busSisMetal.clsPedidoVenda
    'Altera ou incluiu pedido
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objPedidoVenda.AlterarPedidoVenda lngPEDIDOVENDAID, _
                                        lngFICHACLIENTEID, _
                                        IIf(Len(mskDesconto.ClipText) = 0, "", mskDesconto.Text), _
                                        strTipoDesconto, _
                                        lngEMPRESAID
                              
      '
      blnRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Código para inclusão
      '
      objPedidoVenda.InserirPedidoVenda lngPEDIDOVENDAID, _
                                        strStatus, _
                                        giFunIdUsuLib, _
                                        lngFICHACLIENTEID, _
                                        lngTIPOVENDAID, _
                                        IIf(Len(mskDesconto.ClipText) = 0, "", mskDesconto.Text), _
                                        strTipoDesconto, _
                                        lngEMPRESAID
      '
      blnRetorno = True
    End If
    Set objPedidoVenda = Nothing
    '
    Status = tpStatus_Alterar
    intCdastro = 1
    StatusItem = tpStatus_Incluir
    cmdAlterar.Enabled = True
    lngITEMPEDIDOVENDAID = 0
    Form_Load
    Form_Activate
  Case 1 'Gravar Item Pedido
    If Not ValidaCamposItem Then Exit Sub
    '
    Set objGeral = New busSisMetal.clsGeral
    '
    lngINSUMOID = 0
    strSql = "SELECT INSUMO.PKID FROM INSUMO " & _
      " INNER JOIN PRODUTO ON INSUMO.PKID = PRODUTO.INSUMOID " & _
      " WHERE INSUMO.CODIGO = " & Formata_Dados(txtCodProdutoFim.Text, tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngINSUMOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    '
    'Valida se produto já cadastrado
    Set objGeral = New busSisMetal.clsGeral
    strSql = "SELECT * FROM ITEM_PEDIDOVENDA " & _
      " WHERE PEDIDOVENDAID = " & Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo) & _
      " AND PRODUTOID = " & Formata_Dados(lngINSUMOID, tpDados_Longo) & _
      " AND PKID <> " & Formata_Dados(lngITEMPEDIDOVENDAID, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      TratarErroPrevisto "Produto já cadastrado para este pedido", "cmdOK_Click"
      Pintar_Controle txtProduto, tpCorContr_Erro
      SetarFoco txtProduto
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objGeral = Nothing
    '
    Set objItemPedidoVenda = New busSisMetal.clsItemPedidoVenda
    'Altera ou incluiu item do pedido
    If StatusItem = tpStatus_Alterar Then
      'Código para alteração
      '
      objItemPedidoVenda.AlterarItemPedidoVenda lngITEMPEDIDOVENDAID, _
                                                lngINSUMOID, _
                                                mskQuantidade.Text, _
                                                giFunIdUsuLib
                              
      '
      blnRetorno = True
    ElseIf StatusItem = tpStatus_Incluir Then
      'Código para inclusão
      '
      objItemPedidoVenda.InserirItemPedidoVenda lngPEDIDOVENDAID, _
                                                lngINSUMOID, _
                                                mskQuantidade.Text, _
                                                giFunIdUsuLib
      '
      blnRetorno = True
    End If
    'NOVO - Este paço calcula os totais lançados no Pedido
    Set objPedidoVenda = New busSisMetal.clsPedidoVenda
    objPedidoVenda.CalculaTotaisVenda lngPEDIDOVENDAID
    
    Set objPedidoVenda = Nothing
    'NOVO - FIM
    Set objItemPedidoVenda = Nothing
    'Limpar campos
    'ITEM PEDIDO
    LimparCampoTexto txtProduto
    LimparCampoTexto txtCodProdutoFim
    LimparCampoTexto txtNomProdutoFim
    LimparCampoMask mskQuantidade
    '
    'Montar RecordSet
    ITEMPED_COLUNASMATRIZ = grdItemPedido.Columns.Count
    ITEMPED_LINHASMATRIZ = 0
    ITEMPED_MontaMatriz
    grdItemPedido.Bookmark = Null
    grdItemPedido.ReBind
    grdItemPedido.ApproxCount = ITEMPED_LINHASMATRIZ
    ''Entra em novo evento de alteração do item
    StatusItem = tpStatus_Incluir
    lngITEMPEDIDOVENDAID = 0
    cmdAlterar.Enabled = True
    '
    SetarFoco txtProduto
  End Select
  
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  '
End Sub

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'PEDIDO
  LimparCampoTexto txtSequencial
  LimparCampoMask mskData(0)
  LimparCampoTexto txtVendedor
  LimparCampoTexto txtCodClieFornFim
  LimparCampoTexto txtNomeClieFornFim
  LimparCampoOption optDesconto
  LimparCampoMask mskDesconto

  'ITEM PEDIDO

  LimparCampoTexto txtProduto
  LimparCampoTexto txtCodProdutoFim
  LimparCampoTexto txtNomProdutoFim
  LimparCampoMask mskQuantidade
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmPedidoItemVendaInc.LimparCampos]", _
            Err.Description
End Sub

Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  '
  If txtCodClieFornFim.Text = "" Then
    Select Case TipoVenda
    'Case tpTipoVenda.tpTipoVenda_Balc
    '  strMsg = strMsg & "Venda de balcão não pode ser alterada" & vbCrLf
    Case tpTipoVenda.tpTipoVenda_Clie
      strMsg = strMsg & "Selecionar o cliente" & vbCrLf
      SetarFoco mskDesconto
    Case tpTipoVenda.tpTipoVenda_Emp
      strMsg = strMsg & "Selecionar a empresa" & vbCrLf
      SetarFoco mskDesconto
    End Select
  End If
  If Not Valida_Moeda(mskDesconto, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Valor do desconto inválido" & vbCrLf
  End If
  If mskDesconto.ClipText & "" <> "" Then
    If Not Valida_Option(optDesconto, blnSetarFocoControle) Then
      strMsg = strMsg & "Selecionar o tipo de desconto" & vbCrLf
    End If
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmPedidoItemVendaInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[frmPedidoItemVendaInc.ValidaCampos]", _
            Err.Description
End Function

Private Function ValidaCamposItem() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  '
  If txtCodProdutoFim.Text = "" Then
    strMsg = strMsg & "Selecionar o produto" & vbCrLf
    SetarFoco txtProduto
  End If
  If Not Valida_Moeda(mskQuantidade, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Quantidade inválida" & vbCrLf
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmPedidoItemVendaInc.ValidaCamposItem]"
    ValidaCamposItem = False
  Else
    ValidaCamposItem = True
  End If
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[frmPedidoItemVendaInc.ValidaCamposItem]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    Select Case TipoVenda
    Case tpTipoVenda.tpTipoVenda_Balc
      SetarFoco txtProduto
    Case tpTipoVenda.tpTipoVenda_Clie
      If intCdastro = 0 Then
        SetarFoco mskDesconto
      Else
        SetarFoco txtProduto
      End If
    Case tpTipoVenda.tpTipoVenda_Emp
      If intCdastro = 0 Then
        SetarFoco mskDesconto
      Else
        SetarFoco txtProduto
      End If
    End Select
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmPedidoItemVendaInc.Form_Activate]"
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs             As ADODB.Recordset
  Dim strSql            As String
  Dim objPedidoVenda    As busSisMetal.clsPedidoVenda
  Dim lngTIPOVENDAID    As Long
  Dim objGeral          As busSisMetal.clsGeral
  '
  blnFechar = False 'Não Pode Fechar pelo X
  blnRetorno = False
  blnPrimeiraVez = True
  '
  AmpS
  Me.Height = 7620
  Me.Width = 11610
  CenterForm Me
  '
  LimparCampos
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdFechar, , , , cmdAlterar
  Select Case TipoVenda
  Case tpTipoVenda.tpTipoVenda_Balc
    Me.Caption = Me.Caption & " - BALCÃO"
    Label1(3).Caption = "Balcão"
    '
    fraPedido.Enabled = False
    fraItemPedido.Enabled = True
    cmdOk.Enabled = True
    cmdAlterar.Enabled = True
    grdItemPedido.Enabled = True
    '
  Case tpTipoVenda.tpTipoVenda_Clie
    Me.Caption = Me.Caption & " - CLIENTE"
    Label1(3).Caption = "Cliente"
    '
    If intCdastro = 0 Then
      fraPedido.Enabled = True
      fraItemPedido.Enabled = False
      cmdOk.Enabled = True
      cmdAlterar.Enabled = False
      grdItemPedido.Enabled = True
    Else
      fraPedido.Enabled = False
      fraItemPedido.Enabled = True
      cmdOk.Enabled = True
      cmdAlterar.Enabled = True
      grdItemPedido.Enabled = True
    End If
  Case tpTipoVenda.tpTipoVenda_Emp
    Me.Caption = Me.Caption & " - EMPRESA"
    Label1(3).Caption = "Empresa"
    '
    If intCdastro = 0 Then
      fraPedido.Enabled = True
      fraItemPedido.Enabled = False
      cmdOk.Enabled = True
      cmdAlterar.Enabled = False
      grdItemPedido.Enabled = True
    Else
      fraPedido.Enabled = False
      fraItemPedido.Enabled = True
      cmdOk.Enabled = True
      cmdAlterar.Enabled = True
      grdItemPedido.Enabled = True
    End If
  End Select
  '
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    '
    txtVendedor = gsNomeUsuLib
    Select Case TipoVenda
    Case tpTipoVenda.tpTipoVenda_Balc
      'Se Pedido Venda Balcão incluir automaticamente e passar apra alteração
      '-------------------------
      'CAPTURAR TIPOVENDAID
      Set objGeral = New busSisMetal.clsGeral
      'TIPOVENDAID
      lngTIPOVENDAID = 0
      strSql = "SELECT TIPOVENDA.PKID FROM TIPOVENDA " & _
        " WHERE TIPOVENDA.DESCRICAO = " & Formata_Dados("BALCÃO", tpDados_Texto)
      Set objRs = objGeral.ExecutarSQL(strSql)
      If Not objRs.EOF Then
        lngTIPOVENDAID = objRs.Fields("PKID").Value
      End If
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      '
      Set objPedidoVenda = New busSisMetal.clsPedidoVenda
      objPedidoVenda.InserirPedidoVenda lngPEDIDOVENDAID, _
                                        "B", _
                                        giFunIdUsuLib, _
                                        0, _
                                        lngTIPOVENDAID, _
                                        "", _
                                        "", _
                                        0
      Status = tpStatus_Alterar
      Form_Load
      '
      Set objPedidoVenda = Nothing
    Case tpTipoVenda.tpTipoVenda_Clie
      Form_KeyPress 19
    Case tpTipoVenda.tpTipoVenda_Emp
      Form_KeyPress 19
    End Select
    
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objPedidoVenda = New busSisMetal.clsPedidoVenda
    Set objRs = objPedidoVenda.ListarPedidoVenda(lngPEDIDOVENDAID)
    '
    If Not objRs.EOF Then
      'Campos fixos
      txtSequencial.Text = Format(objRs.Fields("PED_NUMERO").Value, "0000")
      INCLUIR_VALOR_NO_MASK mskData(0), objRs.Fields("DATA").Value, TpMaskData
      txtVendedor.Text = objRs.Fields("NOME_VENDEDOR").Value
      '
      If objRs.Fields("DESC_TIPOVENDA").Value & "" = "CLIENTE" Then
        txtCodClieFornFim.Text = objRs.Fields("PKID_FICHACLIENTE").Value & ""
        txtNomeClieFornFim.Text = objRs.Fields("NOME_FICHACLIENTE").Value & ""
      ElseIf objRs.Fields("DESC_TIPOVENDA").Value & "" = "EMPRESA" Then
        txtCodClieFornFim.Text = objRs.Fields("PKID_EMRPESA").Value & ""
        txtNomeClieFornFim.Text = objRs.Fields("NOME_EMRPESA").Value & ""
      End If
      'Campos inserts
      INCLUIR_VALOR_NO_MASK mskDesconto, objRs.Fields("VALOR_DESCONTO").Value, TpMaskMoeda
      'Tipo de desconto
      If objRs.Fields("TIPO_DESCONTO").Value & "" = "P" Then
        optDesconto(0).Value = True
      ElseIf objRs.Fields("TIPO_DESCONTO").Value & "" = "V" Then
        optDesconto(1).Value = True
      End If
    End If
    Set objPedidoVenda = Nothing
    '
    'Montar RecordSet
    ITEMPED_COLUNASMATRIZ = grdItemPedido.Columns.Count
    ITEMPED_LINHASMATRIZ = 0
    ITEMPED_MontaMatriz
    grdItemPedido.Bookmark = Null
    grdItemPedido.ReBind
    grdItemPedido.ApproxCount = ITEMPED_LINHASMATRIZ
    '
  End If
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub
Private Sub cmdFechar_Click()
  On Error GoTo trata
  Dim objItemPedidoVenda    As busSisMetal.clsItemPedidoVenda
  Dim objPedidoVenda        As busSisMetal.clsPedidoVenda
  Dim objRs                 As ADODB.Recordset
  Dim strSql                As String
  If Status = tpStatus_Alterar Then
    Set objItemPedidoVenda = New busSisMetal.clsItemPedidoVenda
    Set objRs = objItemPedidoVenda.ListarItemPedidoVenda(lngPEDIDOVENDAID)
    If Not objRs.EOF Then
      'Lançou itens
      '------------
      'IMPRIMIR
      '------------
      'NOVO - IMPRIME PEDIDO EM TELA
      If lngPEDIDOVENDAID = 0 Then
      End If
      IMP_COMP_PEDIDO lngPEDIDOVENDAID, gsNomeEmpresa

'''      frmGerencialPed.Report1.Connect = ConnectRpt
'''      frmGerencialPed.Report1.ReportFileName = gsReportPath & "Pedido.rpt"
'''      '
'''      'If optSai1.Value Then
'''        frmGerencialPed.Report1.Destination = 0 'Video
'''      'ElseIf optSai2.Value Then
'''      '  Report1.Destination = 1   'Impressora
'''      'End If
'''      frmGerencialPed.Report1.CopiesToPrinter = 1
'''      frmGerencialPed.Report1.WindowState = crptMaximized
'''      '
'''      frmGerencialPed.Report1.Formulas(0) = "PEDIDOID = " & Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo)
'''      '
'''      frmGerencialPed.Report1.Action = 1
      '
      blnFechar = True
    Else
      If MsgBox("Atenção: Não foram lançados itens neste pedido. Caso confirme, o pedido será cancelado. Deseja continuar ?", vbYesNo, TITULOSISTEMA) = vbNo Then
        blnFechar = False
        objRs.Close
        Set objRs = Nothing
        Set objItemPedidoVenda = Nothing
        SetarFoco txtProduto
        Exit Sub
      Else
        'cancelar pedido e sair
        '-----------------------
        Set objPedidoVenda = New busSisMetal.clsPedidoVenda
        objPedidoVenda.ExcluirPedidoVenda (lngPEDIDOVENDAID)
        Set objPedidoVenda = Nothing
        blnFechar = True
      End If
    End If
    objRs.Close
    Set objRs = Nothing
    Set objItemPedidoVenda = Nothing
  Else
    'incluir
    blnFechar = True
  End If
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  blnFechar = True
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then Cancel = True
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
  Dim objFichaCliente   As SisMetal.frmFichaClienteInc
  Dim objLoja           As SisMetal.frmLojaInc
  On Error GoTo trata
  '
  Select Case KeyAscii
  Case 19
    'CTRL+S Selecionar cliente/fornecedor
    Select Case TipoVenda
    Case tpTipoVenda.tpTipoVenda_Balc
    Case tpTipoVenda.tpTipoVenda_Clie
      If intCdastro = 0 Then
        Set objFichaCliente = New SisMetal.frmFichaClienteInc
        objFichaCliente.Status = tpStatus_Incluir
        objFichaCliente.intOrigem = 1
        objFichaCliente.Show vbModal
        Set objFichaCliente = Nothing
      End If
    Case tpTipoVenda.tpTipoVenda_Emp
      If intCdastro = 0 Then
        Set objLoja = New SisMetal.frmLojaInc
        objLoja.intTipoLoja = tpLoja.tpLoja_Empresa
        objLoja.Status = tpStatus_Incluir
        objLoja.lngPKID = 0
        objLoja.intOrigem = 1
        objLoja.Show vbModal
        Set objLoja = Nothing
      
      End If
    End Select
      
  End Select
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmPedidoItemVendaInc.Form_KeyPress]"
End Sub


Private Sub grdItemPedido_UnboundReadDataEx( _
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
               Offset + intI, ITEMPED_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, ITEMPED_COLUNASMATRIZ, ITEMPED_LINHASMATRIZ, ITEMPED_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ITEMPED_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmPedidoItemVendaInc.grdGeral_UnboundReadDataEx]"
End Sub

Private Sub mskDesconto_GotFocus()
  Seleciona_Conteudo_Controle mskDesconto
End Sub
Private Sub mskDesconto_LostFocus()
  Pintar_Controle mskDesconto, tpCorContr_Normal
End Sub

Private Sub mskQuantidade_GotFocus()
  Seleciona_Conteudo_Controle mskQuantidade
End Sub
Private Sub mskQuantidade_LostFocus()
  Pintar_Controle mskQuantidade, tpCorContr_Normal
End Sub

Private Sub txtProduto_GotFocus()
  Seleciona_Conteudo_Controle txtProduto
End Sub
Private Sub txtProduto_LostFocus()
  On Error GoTo trata
  Dim objProdutoCons  As SisMetal.frmProdutoCons
  Dim objInsumo       As busSisMetal.clsInsumo
  Dim objRs           As ADODB.Recordset
  If Me.ActiveControl.Name = "cmdFechar" Then Exit Sub

  Pintar_Controle txtProduto, tpCorContr_Normal
  If Len(txtProduto.Text) = 0 Then
    If Len(txtCodProdutoFim.Text) <> 0 And Len(txtNomProdutoFim.Text) <> 0 Then
      Exit Sub
    Else
      Exit Sub
    End If
  End If
  Set objInsumo = New busSisMetal.clsInsumo
  '
  Set objRs = objInsumo.CapturaProduto(txtProduto.Text)
  If objRs.EOF Then
    LimparCampoTexto txtCodProdutoFim
    LimparCampoTexto txtNomProdutoFim
    TratarErroPrevisto "Descrição/Código do produto não cadastrado"
    Pintar_Controle txtProduto, tpCorContr_Erro
    SetarFoco txtProduto
    Exit Sub
  Else
    If objRs.RecordCount = 1 Then
      txtCodProdutoFim.Text = objRs.Fields("CODIGO").Value & ""
      txtNomProdutoFim.Text = objRs.Fields("NOME").Value & ""
    Else
      'Novo : apresentar tela para seleção da linha
      Set objProdutoCons = New SisMetal.frmProdutoCons
      objProdutoCons.QuemChamou = 2
      objProdutoCons.strCodigoProduto = txtProduto.Text
      objProdutoCons.Show vbModal
    End If
    '
    objRs.Close
    Set objRs = Nothing
    Set objInsumo = Nothing
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

