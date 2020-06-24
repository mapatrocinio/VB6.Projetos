VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPedidoItemPedidoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Pedidos"
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
      TabIndex        =   14
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
      TabPicture(0)   =   "userPedidoItemPedidoInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "grdItemPedido"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraFiltro"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraPedido"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.Frame fraPedido 
         Caption         =   "Pedido"
         Height          =   1545
         Left            =   90
         TabIndex        =   19
         Top             =   1200
         Width           =   9195
         Begin VB.CommandButton cmdTodos 
            Caption         =   "&Entrega Fábrica"
            Height          =   255
            Index           =   1
            Left            =   3150
            TabIndex        =   30
            Top             =   1230
            Width           =   1815
         End
         Begin VB.CommandButton cmdTodos 
            Caption         =   "&Entrega Anodizadora"
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   29
            Top             =   1230
            Width           =   1815
         End
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   975
            Index           =   0
            Left            =   90
            ScaleHeight     =   975
            ScaleWidth      =   8895
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   180
            Width           =   8895
            Begin VB.TextBox txtNumeroOS 
               BackColor       =   &H00E0E0E0&
               Height          =   288
               Left            =   1230
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   3
               TabStop         =   0   'False
               Text            =   "txtNumeroOS"
               Top             =   30
               Width           =   1815
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   255
               Left            =   4320
               ScaleHeight     =   255
               ScaleWidth      =   3855
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   30
               Width           =   3855
               Begin MSMask.MaskEdBox mskData 
                  Height          =   255
                  Index           =   0
                  Left            =   1200
                  TabIndex        =   4
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
                  Left            =   450
                  TabIndex        =   22
                  Top             =   0
                  Width           =   615
               End
            End
            Begin VB.ComboBox cboFornecedor 
               Height          =   315
               Left            =   1230
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   330
               Width           =   3435
            End
            Begin VB.ComboBox cboAnodizadora 
               Height          =   315
               Left            =   1230
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   660
               Width           =   3435
            End
            Begin VB.ComboBox cboFabrica 
               Height          =   315
               Left            =   5520
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   660
               Width           =   3405
            End
            Begin MSMask.MaskEdBox mskValor 
               Height          =   255
               Left            =   5520
               TabIndex        =   6
               Top             =   330
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label2 
               Caption         =   "Ano-Número OS"
               Height          =   255
               Index           =   0
               Left            =   30
               TabIndex        =   27
               Top             =   30
               Width           =   1155
            End
            Begin VB.Label Label5 
               Caption         =   "Fornecedor"
               Height          =   195
               Index           =   0
               Left            =   30
               TabIndex        =   26
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label4 
               Caption         =   "Valor Alumínio"
               Height          =   405
               Left            =   4770
               TabIndex        =   25
               Top             =   300
               Width           =   615
            End
            Begin VB.Label Label5 
               Caption         =   "Anodizadora"
               Height          =   195
               Index           =   1
               Left            =   30
               TabIndex        =   24
               Top             =   690
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Fábrica"
               Height          =   195
               Index           =   2
               Left            =   4770
               TabIndex        =   23
               Top             =   690
               Width           =   615
            End
         End
      End
      Begin VB.Frame fraFiltro 
         Caption         =   "Filtro"
         Height          =   885
         Left            =   90
         TabIndex        =   18
         Top             =   330
         Width           =   9195
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Left            =   1290
            MaxLength       =   50
            TabIndex        =   0
            Text            =   "txtCodigo"
            Top             =   180
            Width           =   5865
         End
         Begin VB.TextBox txtLinhaFim 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   3660
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   2
            TabStop         =   0   'False
            Text            =   "txtLinhaFim"
            Top             =   540
            Width           =   3495
         End
         Begin VB.TextBox txtCodigoFim 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   1290
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   1
            TabStop         =   0   'False
            Text            =   "txtCodigoFim"
            Top             =   540
            Width           =   2355
         End
         Begin VB.Label Label1 
            Caption         =   "Nome da Linha/Código Perfil"
            Height          =   615
            Index           =   3
            Left            =   90
            TabIndex        =   28
            Top             =   210
            Width           =   1095
         End
      End
      Begin TrueDBGrid60.TDBGrid grdItemPedido 
         Height          =   2910
         Left            =   90
         OleObjectBlob   =   "userPedidoItemPedidoInc.frx":001C
         TabIndex        =   9
         Top             =   2760
         Width           =   9210
      End
      Begin VB.Label Label1 
         Caption         =   $"userPedidoItemPedidoInc.frx":72EF
         ForeColor       =   &H000000FF&
         Height          =   435
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   6360
         Width           =   8715
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "* Tecle ENTER para mudar de coluna, ao final o sistema validará os dados na base salvando as informações"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   6090
         Width           =   8715
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "* Clique no botão SALVAR para registrar todas as definições de anodização para os perfis ou tecle ESC para sair"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   5820
         Width           =   8715
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7245
      Left            =   9660
      ScaleHeight     =   7245
      ScaleWidth      =   1860
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2115
         Left            =   120
         ScaleHeight     =   2055
         ScaleWidth      =   1545
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4980
         Width           =   1605
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Height          =   885
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   150
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   885
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1020
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmPedidoItemPedidoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public lngPEDIDOID              As Long
Public lngLINHAID               As Long

Public lngOSFINALID             As Long
Public lngANODIZACAOITEMID      As Long
Public lngITEMOSFINALID         As Long
Public lngOSID                  As Long
Public lngCORID                 As Long

Public strOSNumero              As String
Public strCor                   As String
Dim blnAlterouPeso              As Boolean

Dim blnFechar                   As Boolean
Public blnRetorno               As Boolean
Public blnPrimeiraVez           As Boolean
'
Dim ITEMPED_COLUNASMATRIZ        As Long
Dim ITEMPED_LINHASMATRIZ         As Long
Private ITEMPED_Matriz()         As String


Public Sub ITEMPED_MontaMatriz(lngLINHASELID As Long)
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGeral  As busSisMetal.clsGeral
  '
  On Error GoTo trata

  Set objGeral = New busSisMetal.clsGeral
  '
  strSql = "SELECT " & _
            IIf(Status = tpStatus_Incluir, "0,", " ITEM_PEDIDO.PKID, ") & _
            " ESTOQUE.LINHAID, " & _
            " ESTOQUE.LINHAID, " & _
            Formata_Dados(IIf(Status = tpStatus_Incluir, "0", "0"), tpDados_Texto) & ", " & _
            " ESTOQUE.NOME, ESTOQUE.CODIGO, " & _
            " ESTOQUE.PESO_MINIMO, " & _
            "  (ESTOQUE.PESO_ESTOQUE + ESTOQUE.PEDIDO_PESO_RESTA + ESTOQUE.OS_PESO_RESTA + ESTOQUE.ANOD_PESO_RESTA) AS PESO_REAL, "
  
  If Status = tpStatus_Incluir Then
    strSql = strSql & " (ESTOQUE.PESO_MINIMO) - (ESTOQUE.PESO_ESTOQUE + ESTOQUE.PEDIDO_PESO_RESTA + ESTOQUE.OS_PESO_RESTA + ESTOQUE.ANOD_PESO_RESTA) AS PESO_APEDIR, '' AS PESO_ANOD, '' AS PESO_FAB "
  Else
    strSql = strSql & " ITEM_PEDIDO.PESO_INI, ITEM_PEDIDO.PESO, ITEM_PEDIDO.PESO_FAB "
  End If
  
  strSql = strSql & " From VW_CONS_ESTOQUE_PERFIL AS ESTOQUE "
  If Status = tpStatus_Incluir Then
    strSql = strSql & " WHERE (ESTOQUE.PESO_ESTOQUE + ESTOQUE.PEDIDO_PESO_RESTA + ESTOQUE.OS_PESO_RESTA + ESTOQUE.ANOD_PESO_RESTA) < ESTOQUE.PESO_MINIMO "
    If lngLINHASELID <> 0 Then
      strSql = strSql & " AND ESTOQUE.LINHAID = " & Formata_Dados(lngLINHASELID, tpDados_Longo)
    End If
  Else
    strSql = strSql & " INNER JOIN ITEM_PEDIDO ON ITEM_PEDIDO.LINHAID = ESTOQUE.LINHAID "
    strSql = strSql & " WHERE ITEM_PEDIDO.PEDIDOID = " & Formata_Dados(lngPEDIDOID, tpDados_Longo)
  End If
  strSql = strSql & " ORDER BY ESTOQUE.NOME, ESTOQUE.CODIGO"
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
          If intJ = ITEMPED_COLUNASMATRIZ - 1 Then
            ITEMPED_Matriz(intJ, intI) = intI & ""
          Else
            ITEMPED_Matriz(intJ, intI) = objRs(intJ) & ""
          End If
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

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim objPedido               As busSisMetal.clsPedido
  Dim objRs                   As ADODB.Recordset
  Dim objGer                  As busSisMetal.clsGeral
  Dim lngFORNECEDORID         As Long
  Dim lngANODIZADORAID        As Long
  Dim lngFABRICAID            As Long
  Dim objItemPedido           As busSisMetal.clsItemPedido
  Dim intI      As Integer
  '
  Select Case tabDetalhes.Tab
  Case 0 'Gravar Anodização
    If Not ValidaCampos Then Exit Sub
  
    If ValidaCamposAnodOrigemAll Then
      SetarFoco grdItemPedido
      If Status = tpStatus_Incluir Or gsNivel <> gsCompra Then
        grdItemPedido.Col = 8
      Else
        grdItemPedido.Col = 9
      End If
      grdItemPedido.Row = 0
      Exit Sub
    End If
    'OK procede com o cadastro
    'CADASTRO DE PEDIDO
    '-------------------------
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
      '
      blnRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Código para inclusão
      '
      objPedido.InserirPedido lngPEDIDOID, _
                              lngFORNECEDORID, _
                              lngANODIZADORAID, _
                              lngFABRICAID, _
                              IIf(Len(mskValor.ClipText) = 0, "", mskValor.Text)
      '
      blnRetorno = True
    End If
    Set objPedido = Nothing
    '
    Set objItemPedido = New busSisMetal.clsItemPedido
    For intI = 0 To ITEMPED_LINHASMATRIZ - 1
      grdItemPedido.Bookmark = CLng(intI)
      'If grdItemPedido.Columns("Branco").Text & "" <> "" Or _
        grdItemPedido.Columns("Brilho").Text & "" <> "" Or _
        grdItemPedido.Columns("Bronze").Text & "" <> "" Or _
        grdItemPedido.Columns("Natural").Text & "" <> "" Then
      If grdItemPedido.Columns("*").Text & "" = "-1" Then
        'Propósito: Cadastrar pedido
        '
        objItemPedido.InserirItemPedidoItem grdItemPedido.Columns("ITEM_PEDIDOID").Text & "", _
                                            lngPEDIDOID, _
                                            grdItemPedido.Columns("LINHAID").Text & "", _
                                            IIf(grdItemPedido.Columns("Peso").Text & "" = "", "0", grdItemPedido.Columns("Peso").Text & ""), _
                                            grdItemPedido.Columns("Anod.").Text, _
                                            grdItemPedido.Columns("Fábrica").Text
                                            
        blnRetorno = True
      End If
    Next
    Set objItemPedido = Nothing
    '
    blnFechar = True
    Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  '
  grdItemPedido.Bookmark = Null
  grdItemPedido.ReBind
  SetarFoco grdItemPedido
  If grdItemPedido.Row <> -1 Then
    If Status = tpStatus_Incluir Or gsNivel <> gsCompra Then
      grdItemPedido.Col = 8
    Else
      grdItemPedido.Col = 9
    End If
    grdItemPedido.Row = 0
  End If
End Sub

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Filtro
  LimparCampoTexto txtCodigo
  LimparCampoTexto txtCodigoFim
  LimparCampoTexto txtLinhaFim
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
            "[frmPedidoItemPedidoInc.LimparCampos]", _
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
  If Not Valida_String(cboAnodizadora, IIf(Status = tpStatus_Incluir Or gsNivel <> gsCompra, TpnaoObrigatorio, TpObrigatorio), blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar a anodizadora" & vbCrLf
  End If
  If Not Valida_String(cboFabrica, IIf(Status = tpStatus_Incluir Or gsNivel <> gsCompra, TpnaoObrigatorio, TpObrigatorio), blnSetarFocoControle) Then
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
            "[frmPedidoItemPedidoInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub cmdTodos_Click(Index As Integer)
  On Error GoTo trata
  Dim strMsg As String
  If Index = 0 Then
    'Entrega anodizadora
    strMsg = "Confirma definição do peso para entrega na anodizadora para todos os ítens?" & vbCrLf & vbCrLf & _
        "Caso haja algum valor já lançado no grid irá ser perdido." & vbCrLf & _
        "Após confirmação você pode entrar em um ítem específico e alterá-lo." & vbCrLf & _
        "Mesmo após esta ação será necessário clicar no botão confirmar para efetivar as alterações no pedido."
  Else
    'Entrega fábrica
    strMsg = "Confirma definição do peso para entrega na fábrica para todos os ítens?" & vbCrLf & vbCrLf & _
        "Caso haja algum valor já lançado no grid irá ser perdido." & vbCrLf & _
        "Após confirmação você pode entrar em um ítem específico e alterá-lo." & vbCrLf & _
        "Mesmo após esta ação será necessário clicar no botão confirmar para efetivar as alterações no pedido."
  End If
  If MsgBox(strMsg, vbYesNo, TITULOSISTEMA) = vbNo Then
    SetarFoco grdItemPedido
    Exit Sub
  End If
  'Ok
  DefinirQuantidade Index
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmPedidoItemPedidoInc.cmdTodos_Click]"
End Sub

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Montar RecordSet
    ITEMPED_COLUNASMATRIZ = grdItemPedido.Columns.Count
    ITEMPED_LINHASMATRIZ = 0
    ITEMPED_MontaMatriz (lngLINHAID)
    grdItemPedido.Bookmark = Null
    grdItemPedido.ReBind
    grdItemPedido.ApproxCount = ITEMPED_LINHASMATRIZ
    '
    If Status = tpStatus_Incluir Then
      SetarFoco txtCodigo
    ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
      SetarFoco cboFornecedor
    End If
    'SetarFoco grdItemPedido
    'grdItemPedido.Col = 8
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmPedidoItemPedidoInc.Form_Activate]"
End Sub


Private Sub grdItemPedido_BeforeUpdate(Cancel As Integer)
  On Error GoTo trata
  'Atualiza Matriz
  If blnAlterouPeso = True Then
    ITEMPED_Matriz(3, grdItemPedido.Columns("ROWNUM").Value) = "-1"
  Else
    ITEMPED_Matriz(3, grdItemPedido.Columns("ROWNUM").Value) = grdItemPedido.Columns(3).Text
  End If
  If Status = tpStatus_Incluir Or gsNivel <> gsCompra Then
    ITEMPED_Matriz(8, grdItemPedido.Columns("ROWNUM").Value) = grdItemPedido.Columns(8).Text
  Else
    ITEMPED_Matriz(9, grdItemPedido.Columns("ROWNUM").Value) = grdItemPedido.Columns(9).Text
    ITEMPED_Matriz(10, grdItemPedido.Columns("ROWNUM").Value) = grdItemPedido.Columns(10).Text
  End If
  blnAlterouPeso = False
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmPedidoItemPedidoInc.grdItemPedido_BeforeRowColChange]"
End Sub

Private Sub grdItemPedido_ColEdit(ByVal ColIndex As Integer)
  On Error GoTo trata
  '
  If grdItemPedido.Col = 8 Or grdItemPedido.Col = 9 Or grdItemPedido.Col = 10 Then
    blnAlterouPeso = True
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmPedidoItemPedidoInc.grdItemPedido_ColEdit]"
End Sub

Private Sub grdItemPedido_GotFocus()
  On Error Resume Next
  If Status = tpStatus_Incluir Or gsNivel <> gsCompra Then
    grdItemPedido.Col = 8
  Else
    grdItemPedido.Col = 9
  End If
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
  TratarErro Err.Number, Err.Description, "[frmAnodizadoraInc.grdItemPedido_UnboundReadDataEx]"
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  Dim objPedido     As busSisMetal.clsPedido
  '
  blnFechar = False 'Não Pode Fechar pelo X
  blnRetorno = False
  blnPrimeiraVez = True
  blnAlterouPeso = False
  lngLINHAID = 0
  '
  AmpS
  Me.Height = 7620
  Me.Width = 11610
  CenterForm Me
  '
  LimparCampos
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdFechar
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
  '
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    '
    fraFiltro.Enabled = True
    fraPedido.Enabled = True
    cmdOk.Enabled = True
    grdItemPedido.Enabled = True
    'No evento de inclusão deve ser habilitado a coluna peso
    grdItemPedido.Columns(8).Locked = False
    grdItemPedido.Columns(9).Visible = False
    grdItemPedido.Columns(10).Visible = False
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objPedido = New busSisMetal.clsPedido
    Set objRs = objPedido.ListarPedido(lngPEDIDOID)
    '
    If Not objRs.EOF Then
      'Campos fixos
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
    fraFiltro.Enabled = False
    If Status = tpStatus_Alterar Then
      fraPedido.Enabled = True
      cmdOk.Enabled = True
      grdItemPedido.Enabled = True
    Else
      fraPedido.Enabled = False
      cmdOk.Enabled = False
      grdItemPedido.Enabled = False
    End If

    'No evento de alteração deve ser habilitado as colunas anod e fabrica
    If gsNivel <> gsCompra Then
      grdItemPedido.Columns(8).Locked = False
      grdItemPedido.Columns(9).Visible = False
      grdItemPedido.Columns(10).Visible = False
      cmdTodos(0).Visible = False
      cmdTodos(1).Visible = False
    Else
      grdItemPedido.Columns(8).Locked = True
      grdItemPedido.Columns(9).Visible = True
      grdItemPedido.Columns(10).Visible = True
      cmdTodos(0).Visible = True
      cmdTodos(1).Visible = True
    End If
  End If
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Function ValidaCamposAnodOrigemLinha(intLinha As Integer) As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  Dim strSql                As String
  Dim objRs                 As ADODB.Recordset
  Dim objGeral              As busSisMetal.clsGeral
  Dim objItemPedido         As busSisMetal.clsItemPedido
  '
  Dim lngQtdIni             As Long
  Dim lngQtdAnod            As Long
  Dim lngQtdFab             As Long

  '
  blnSetarFocoControle = True
  '
  strMsg = ""
  'Validção dos ítens do pedido
  If Not Valida_Moeda(grdItemPedido.Columns("Peso"), TpnaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
    strMsg = strMsg & "Peso inválido na linha " & intLinha + 1 & vbCrLf
  End If
  If Not Valida_Moeda(grdItemPedido.Columns("Anod."), TpnaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
    strMsg = strMsg & "Peso para anodizadora inválido na linha " & intLinha + 1 & vbCrLf
  End If
  If Not Valida_Moeda(grdItemPedido.Columns("Fábrica"), TpnaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
    strMsg = strMsg & "Peso inválido para fábrica na linha " & intLinha + 1 & vbCrLf
  End If
  '
  If Len(strMsg) = 0 Then
    'NOVO - Validações de cálculo de peças (quantidade)
    Set objItemPedido = New busSisMetal.clsItemPedido
    lngQtdIni = objItemPedido.CalculoQuantidadePedido(grdItemPedido.Columns("LINHAID").Value, _
                                                      grdItemPedido.Columns("Peso").Value)
    If lngQtdIni = 0 Then
      strMsg = strMsg & "A quantidade calculada para o perfil deve ser maior que zero na linha " & intLinha + 1 & vbCrLf
    End If
    If grdItemPedido.Columns("Anod.").Value & "" <> "" Then
      'Lançou peso para anodização
      lngQtdAnod = objItemPedido.CalculoQuantidadePedido(grdItemPedido.Columns("LINHAID").Value, _
                                                        grdItemPedido.Columns("Anod.").Value)
      If lngQtdAnod = 0 Then
        strMsg = strMsg & "A quantidade calculada para anodização para o perfil deve ser maior que zero na linha " & intLinha + 1 & vbCrLf
      End If
    End If
    If grdItemPedido.Columns("Fábrica").Value & "" <> "" Then
      'Lançou peso para anodização
      lngQtdFab = objItemPedido.CalculoQuantidadePedido(grdItemPedido.Columns("LINHAID").Value, _
                                                        grdItemPedido.Columns("Fábrica").Value)
      If lngQtdFab = 0 Then
        strMsg = strMsg & "A quantidade calculada para fábrica para o perfil deve ser maior que zero na linha " & intLinha + 1 & vbCrLf
      End If
    End If

    
'''    lngTotal = 0
'''    lngTotalANOD = 0
'''    '
'''    lngTotal = CLng(grdItemPedido.Columns("Qtd. Total")) - CLng(grdItemPedido.Columns("Qtd. Baixa"))
'''    lngTotalANOD = CLng(IIf(Not IsNumeric(grdItemPedido.Columns("Quantidade")), 0, grdItemPedido.Columns("Quantidade")))
    '
    Set objItemPedido = Nothing
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmPedidoItemPedidoInc.ValidaCamposAnodOrigemLinha]"
    ValidaCamposAnodOrigemLinha = False
  Else
    ValidaCamposAnodOrigemLinha = True
  End If
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserLeituraFechaInc.ValidaCamposAnodOrigemLinha]"
  ValidaCamposAnodOrigemLinha = False
End Function

Private Sub DefinirQuantidade(intIndex As Integer)
  On Error GoTo trata
  Dim intRows As Integer
  For intRows = 0 To ITEMPED_LINHASMATRIZ - 1
    grdItemPedido.Bookmark = CLng(intRows)
    '
    If grdItemPedido.Columns("Peso").Text & "" <> "" Then
      If intIndex = 0 Then
        'Tudo anodizadora
        grdItemPedido.Columns("Fábrica").Text = ""
        grdItemPedido.Columns("Anod.").Text = grdItemPedido.Columns("Peso").Text & ""
      Else
        grdItemPedido.Columns("Anod.").Text = ""
        grdItemPedido.Columns("Fábrica").Text = grdItemPedido.Columns("Peso").Text & ""
      End If
      grdItemPedido.Columns("*").Text = "-1"
      'Atualiza matriz
      ITEMPED_Matriz(3, grdItemPedido.Columns("ROWNUM").Value) = "-1"
      ITEMPED_Matriz(9, grdItemPedido.Columns("ROWNUM").Value) = grdItemPedido.Columns(9).Text
      ITEMPED_Matriz(10, grdItemPedido.Columns("ROWNUM").Value) = grdItemPedido.Columns(10).Text
      
    End If
  Next
  '
  grdItemPedido.ReBind
  grdItemPedido.SetFocus
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserLeituraFechaInc.DefinirQuantidade]"
End Sub

Private Function ValidaCamposAnodOrigemAll() As Boolean
  On Error GoTo trata
  Dim blnRetorno            As Boolean
  Dim blnCadastrou1Linha    As Boolean
  Dim blnEncontrouErro      As Boolean
  Dim blnEncontrouErroLinha As Boolean
  Dim intRows               As Integer
  'Validar todas as linhas da matriz
  blnEncontrouErro = False
  blnCadastrou1Linha = False
  blnEncontrouErroLinha = False
  blnRetorno = True
  
  
  For intRows = 0 To ITEMPED_LINHASMATRIZ - 1
    grdItemPedido.Bookmark = CLng(intRows)
    '
    If grdItemPedido.Columns("*").Text & "" = "-1" Then
      'Somente válida se preencheu algo, sneão considera ok
      If grdItemPedido.Columns("Peso").Text & "" <> "" Then
        If Not ValidaCamposAnodOrigemLinha(grdItemPedido.Row) Then
          blnEncontrouErro = True
          blnEncontrouErroLinha = True
        Else
          blnCadastrou1Linha = True
        End If
      Else
        'tudo brnao, considera OK
        blnCadastrou1Linha = True
      End If
    Else
      'blnEncontrouErro = True
    End If
    If blnEncontrouErro = True Then Exit For
  Next
  '
  If blnEncontrouErro = False And blnCadastrou1Linha = True Then
    blnRetorno = False
  End If
  If blnEncontrouErroLinha = False And blnEncontrouErro = False And blnCadastrou1Linha = False Then
    'não ouve erro
    If Status = tpStatus_Incluir Or gsNivel <> gsCompra Then
      TratarErroPrevisto "Selecione no mínimo 1 perfil para cadastro", "[frmPedidoItemPedidoInc.ValidaCamposAnodOrigemAll]"
    Else
      'No caso da alteração não é obrigatório cadastrar ou alterar ítem
      'pode estar apenas alterando dados do pedido
      blnRetorno = False
    End If
    
  End If
  grdItemPedido.ReBind
  grdItemPedido.SetFocus
  ValidaCamposAnodOrigemAll = blnRetorno
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserLeituraFechaInc.ValidaCamposAnodOrigemAll]"
  ValidaCamposAnodOrigemAll = False
End Function
Private Function ValidaCamposAnodOrigem(intLinha As Integer, intColuna As Integer) As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  '
  Select Case intColuna
  Case 5
    'Validção da quantidade branco
    If Not Valida_Moeda(grdItemPedido.Columns(5), TpnaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
      strMsg = strMsg & "Quantidade de perfil branco inválida" & vbCrLf
    End If
  Case 6
    'Validção da quantidade brilho
    If Not Valida_Moeda(grdItemPedido.Columns(6), TpnaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
      strMsg = strMsg & "Quantidade de perfil brilho inválida" & vbCrLf
    End If
  Case 7
    'Validção da quantidade bronze
    If Not Valida_Moeda(grdItemPedido.Columns(7), TpnaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
      strMsg = strMsg & "Quantidade de perfil bronze inválida" & vbCrLf
    End If
  Case 8
    'Validção da quantidade natural
    If Not Valida_Moeda(grdItemPedido.Columns(8), TpnaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
      strMsg = strMsg & "Quantidade de perfil natural inválida" & vbCrLf
    End If
  End Select
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmPedidoItemPedidoInc.ValidaCamposAnodOrigem]"
    ValidaCamposAnodOrigem = False
  Else
    ValidaCamposAnodOrigem = True
  End If
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserGrupoCln.ValidaCamposAnodOrigem]"
  ValidaCamposAnodOrigem = False
End Function
Private Sub cmdFechar_Click()
  blnFechar = True
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then Cancel = True
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  '
  Dim intUltimaColuna   As Integer
  If Me.ActiveControl.Name <> "grdItemPedido" Then
    If KeyAscii = 13 Then
      SendKeys "{tab}"
    End If
  Else
    If Status = tpStatus_Incluir Or gsNivel <> gsCompra Then
      intUltimaColuna = 8
    Else
      intUltimaColuna = 10
    End If
    If KeyAscii = 13 And IsNumeric(grdItemPedido.Columns("ROWNUM").Value & "") = True Then
      If grdItemPedido.Col = intUltimaColuna Then
        If grdItemPedido.Columns("ROWNUM").Value + 1 = ITEMPED_LINHASMATRIZ Then
          cmdOk_Click
        Else
          If Status = tpStatus_Incluir Or gsNivel <> gsCompra Then
            grdItemPedido.Col = intUltimaColuna
          Else
            grdItemPedido.Col = intUltimaColuna - 1
          End If
          
          grdItemPedido.MoveNext
        End If
      Else
        grdItemPedido.Col = grdItemPedido.Col + 1
      End If
    ElseIf (KeyAscii = 8) Or (KeyAscii = 44) Then
    ElseIf (KeyAscii < 48 Or KeyAscii > 57) Then
      KeyAscii = 0
    End If
  End If
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmPedidoItemPedidoInc.Form_KeyPress]"
End Sub


Private Sub mskValor_GotFocus()
  Seleciona_Conteudo_Controle mskValor
End Sub
Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub

Private Sub txtCodigo_LostFocus()
  On Error GoTo trata
  Dim objLinhaCons    As Form
  Dim objLinhaPerfil  As busSisMetal.clsLinhaPerfil
  Dim objRs           As ADODB.Recordset
  Dim blnSelecionou   As Boolean
  If Me.ActiveControl.Name = "cmdFechar" Then Exit Sub
  If Me.ActiveControl.Name = "cmdOk" Then Exit Sub

  Pintar_Controle txtCodigo, tpCorContr_Normal
  If Len(txtCodigo.Text) = 0 Then
    If Len(txtCodigoFim.Text) <> 0 And Len(txtLinhaFim.Text) <> 0 Then
      lngLINHAID = 0
      '
      ITEMPED_MontaMatriz (lngLINHAID)
      grdItemPedido.Bookmark = Null
      grdItemPedido.ReBind
      grdItemPedido.ApproxCount = ITEMPED_LINHASMATRIZ
      '
      Exit Sub
    Else
      'TratarErroPrevisto "Entre com o código ou descrição da linha."
      'Pintar_Controle txtCodigo, tpCorContr_Erro
      'SetarFoco txtCodigo
      'Exit Sub
      lngLINHAID = 0
      '
      ITEMPED_MontaMatriz (lngLINHAID)
      grdItemPedido.Bookmark = Null
      grdItemPedido.ReBind
      grdItemPedido.ApproxCount = ITEMPED_LINHASMATRIZ
      '
      Exit Sub
      
    End If
  End If
  blnSelecionou = False
  Set objLinhaPerfil = New busSisMetal.clsLinhaPerfil
  '
  Set objRs = objLinhaPerfil.CapturaItemLinha(txtCodigo.Text)
  If objRs.EOF Then
    LimparCampoTexto txtCodigoFim
    LimparCampoTexto txtLinhaFim
    TratarErroPrevisto "Descrição/Código da linha não cadastrado"
    Pintar_Controle txtCodigo, tpCorContr_Erro
    SetarFoco txtCodigo
    Exit Sub
  Else
    If objRs.RecordCount = 1 Then
      txtCodigoFim.Text = objRs.Fields("CODIGO").Value & ""
      txtLinhaFim.Text = objRs.Fields("NOME").Value & ""
      lngLINHAID = objRs.Fields("PKID").Value & ""
      blnSelecionou = True
    Else
      'Novo : apresentar tela para seleção da linha
      Set objLinhaCons = New frmLinhaCons
      objLinhaCons.intIcOrigemLn = 4
      objLinhaCons.strCodigoDescricao = txtCodigo.Text
      objLinhaCons.Show vbModal
      blnSelecionou = True
      Set objLinhaCons = Nothing
    End If
    '
    objRs.Close
    Set objRs = Nothing
    Set objLinhaPerfil = Nothing
  End If
  ''''recarregar GRID
  If blnSelecionou = True Then
    '
    ITEMPED_MontaMatriz (lngLINHAID)
    grdItemPedido.Bookmark = Null
    grdItemPedido.ReBind
    grdItemPedido.ApproxCount = ITEMPED_LINHASMATRIZ
  End If
  
  
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  KeyAscii = TRANSFORMA_MAIUSCULA(KeyAscii)
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtCodigo_GotFocus()
  Seleciona_Conteudo_Controle txtCodigo
End Sub

