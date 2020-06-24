VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmUserEstIntermediarioMov 
   Caption         =   "Visualização de movimentação de itens do estoque intermediário"
   ClientHeight    =   6045
   ClientLeft      =   2595
   ClientTop       =   3120
   ClientWidth     =   11415
   Icon            =   "userEstIntermediarioMov.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6045
   ScaleWidth      =   11415
   Begin VB.CommandButton cmdFiltrar 
      Height          =   735
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cboEstInter 
      Height          =   315
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   150
      Width           =   3135
   End
   Begin VB.PictureBox picBotoes 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   11415
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5145
      Width           =   11415
      Begin VB.PictureBox picAlinDir 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   912
         Left            =   0
         ScaleHeight     =   915
         ScaleWidth      =   11295
         TabIndex        =   1
         Top             =   0
         Width           =   11295
         Begin VB.CommandButton cmdImprimir 
            Height          =   735
            Left            =   8760
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdFechar 
            Height          =   735
            Left            =   9960
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   120
            Width           =   1215
         End
      End
   End
   Begin TrueDBGrid60.TDBGrid grdGeral 
      Height          =   3855
      Left            =   0
      OleObjectBlob   =   "userEstIntermediarioMov.frx":000C
      TabIndex        =   2
      Top             =   1080
      Width           =   11415
   End
   Begin VB.Label Label2 
      Caption         =   "Estoque intermediário"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   150
      Width           =   1815
   End
End
Attribute VB_Name = "frmUserEstIntermediarioMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String
Dim blnPrimeiraVez        As Boolean

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdFiltrar_Click()
  Dim strWhere As String
  '
  If cboEstInter.Text = "" Then Exit Sub
  'strWhere = "WHERE"
  If cboEstInter.Text <> "<TODOS>" And cboEstInter.Text <> "" Then
    strWhere = strWhere & " DESCGRUPOEST = " & Formata_Dados(cboEstInter.Text, tpDados_Texto, tpNulo_NaoAceita, 255) & " "
  End If

  'strWhere = IIf(strWhere = "WHERE", "", strWhere)

  '
  COLUNASMATRIZ = 14
  LINHASMATRIZ = 0

  MontaMatriz strWhere
  grdGeral.Bookmark = Null
  grdGeral.ReBind
  grdGeral.ApproxCount = LINHASMATRIZ
  grdGeral.SetFocus
  
End Sub

Private Sub cmdImprimir_Click()
  On Error GoTo TratErro
  AmpS
  '
  'Cabeçalho do report
  grdGeral.PrintInfo.PageHeader = "Movimentação de itens do estoque intermediário - emissão: " & Format(Now, "DD/MM/YYYY hh:mm")
  grdGeral.PrintInfo.PageHeader = grdGeral.PrintInfo.PageHeader & vbCrLf & cboEstInter.Text
  grdGeral.PrintInfo.RepeatColumnHeaders = True
  '
  grdGeral.PrintInfo.SettingsMarginBottom = 400
  grdGeral.PrintInfo.SettingsMarginLeft = 1000
  grdGeral.PrintInfo.SettingsMarginRight = 1000
  grdGeral.PrintInfo.SettingsMarginTop = 600
  grdGeral.PrintInfo.PreviewMaximize = True
  grdGeral.PrintInfo.SettingsOrientation = 2
  grdGeral.PrintInfo.PrintPreview
  '
  AmpN
  Exit Sub
  
TratErro:
  AmpN
  MsgBox "O seguinte Erro Ocorreu: " & Err.Description, vbOKOnly, TITULOSISTEMA

End Sub

Public Sub MontaMatriz(Optional strWhere As String)
  Dim strSql              As String
  Dim objRs               As ADODB.Recordset
  Dim objRsRequisicao     As ADODB.Recordset
  Dim objRsRetEst         As ADODB.Recordset
  Dim objRsBaixaEst       As ADODB.Recordset
  Dim objRsRetDepEst      As ADODB.Recordset
  Dim objRsTransfEntrada  As ADODB.Recordset
  Dim objRsTransfSaida    As ADODB.Recordset
  Dim objRsVenda          As ADODB.Recordset
  Dim objRsPedido         As ADODB.Recordset

  Dim intI                As Integer
  Dim intJ                As Integer
  Dim clsGer              As busSisMotel.clsGeral
  '
  Dim lngQtdAnterior      As Long
  Dim lngQtdRequisicao    As Long
  Dim lngQtdRetReq        As Long
  Dim lngQtdBaixaEst      As Long
  Dim lngQtdRetDepEst     As Long
  Dim lngQtdTransfEntrada As Long
  Dim lngQtdTransfSaida   As Long
  Dim lngQtdVenda         As Long
  Dim lngQtdPedido        As Long
  Dim lngQtdTotal         As Long
  
  Dim strDtInicial        As String
  Dim strDtFinal          As String

  AmpS
  On Error GoTo trata
  '
  strDtInicial = Format(gdMovimentacao, "DD/MM/YYYY hh:mm")
  strDtFinal = Format(Now, "DD/MM/YYYY hh:mm")
  
  Set clsGer = New busSisMotel.clsGeral
  '
  strSql = "Select GRUPOESTOQUE.DESCRICAO AS DESCGRUPOEST, ESTOQUEINTERMEDIARIO.CODIGO, ESTOQUEINTERMEDIARIO.DESCRICAO, IIF(ISNULL(TAB_GRUPOESTESTINTER.QTDANTERIOR), 0, TAB_GRUPOESTESTINTER.QTDANTERIOR) AS QTDANT, IIF(ISNULL(TAB_GRUPOESTESTINTER.QTDESTOQUE), 0, TAB_GRUPOESTESTINTER.QTDESTOQUE) AS QTDEST, ESTOQUEINTERMEDIARIO.UNIDADE " & _
   "From (((GRUPOESTOQUE LEFT JOIN TAB_GRUPOESTESTINTER ON GRUPOESTOQUE.PKID = TAB_GRUPOESTESTINTER.GRUPOESTOQUEID) " & _
   "LEFT JOIN ESTOQUEINTERMEDIARIO ON ESTOQUEINTERMEDIARIO.PKID = TAB_GRUPOESTESTINTER.ESTOQUEINTERMEDIARIOID) " & _
   "LEFT JOIN ESTOQUE ON ESTOQUE.PKID = ESTOQUEINTERMEDIARIO.ESTOQUEID) " & _
   "LEFT JOIN FAMILIAPRODUTOS ON FAMILIAPRODUTOS.PKID = ESTOQUE.FAMILIAPRODUTOSID " & _
   " Order By ESTOQUEINTERMEDIARIO.Descricao,  GRUPOESTOQUE.DESCRICAO"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  '
  objRs.Filter = strWhere
  '
  strSql = "Select GRUPOESTOQUE.DESCRICAO AS DESCGRUPOEST, ESTOQUEINTERMEDIARIO.CODIGO, TAB_REQUISICAOMATERIAL.QUANTIDADEATENDIDA " & _
   "From ((((ESTOQUE INNER JOIN TAB_REQUISICAOMATERIAL ON ESTOQUE.PKID = TAB_REQUISICAOMATERIAL.ESTOQUEID) " & _
   "INNER JOIN REQUISICAOMATERIAL ON REQUISICAOMATERIAL.PKID = TAB_REQUISICAOMATERIAL.REQUISICAOMATERIALID) " & _
   "INNER JOIN ESTOQUEINTERMEDIARIO ON ESTOQUE.PKID = ESTOQUEINTERMEDIARIO.ESTOQUEID) " & _
   "INNER JOIN TAB_GRUPOESTESTINTER ON ESTOQUEINTERMEDIARIO.PKID = TAB_GRUPOESTESTINTER.ESTOQUEINTERMEDIARIOID AND TAB_REQUISICAOMATERIAL.GRUPOESTOQUEID = TAB_GRUPOESTESTINTER.GRUPOESTOQUEID) " & _
   "INNER JOIN GRUPOESTOQUE ON GRUPOESTOQUE.PKID = TAB_GRUPOESTESTINTER.GRUPOESTOQUEID " & _
   "WHERE REQUISICAOMATERIAL.DATA >= " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
   " AND REQUISICAOMATERIAL.DATA <= " & Formata_Dados(strDtFinal, tpDados_DataHora, tpNulo_NaoAceita) & _
   " Order By ESTOQUEINTERMEDIARIO.Descricao,  GRUPOESTOQUE.DESCRICAO"
  '
  Set objRsRequisicao = clsGer.ExecutarSQL(strSql)
  '
  objRsRequisicao.Filter = strWhere
  '
  strSql = "Select GRUPOESTOQUE.DESCRICAO AS DESCGRUPOEST, ESTOQUEINTERMEDIARIO.CODIGO, TAB_RETORNOREQUISICAO.QUANTIDADE " & _
   "From ((((ESTOQUE INNER JOIN TAB_RETORNOREQUISICAO ON ESTOQUE.PKID = TAB_RETORNOREQUISICAO.ESTOQUEID) " & _
   "INNER JOIN RETORNOREQUISICAO ON RETORNOREQUISICAO.PKID = TAB_RETORNOREQUISICAO.RETORNOREQUISICAOID) " & _
   "INNER JOIN ESTOQUEINTERMEDIARIO ON ESTOQUE.PKID = ESTOQUEINTERMEDIARIO.ESTOQUEID) " & _
   "INNER JOIN TAB_GRUPOESTESTINTER ON ESTOQUEINTERMEDIARIO.PKID = TAB_GRUPOESTESTINTER.ESTOQUEINTERMEDIARIOID AND TAB_RETORNOREQUISICAO.GRUPOESTOQUEID = TAB_GRUPOESTESTINTER.GRUPOESTOQUEID) " & _
   "INNER JOIN GRUPOESTOQUE ON GRUPOESTOQUE.PKID = TAB_GRUPOESTESTINTER.GRUPOESTOQUEID " & _
   "WHERE RETORNOREQUISICAO.DATA >= " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
   " AND RETORNOREQUISICAO.DATA <= " & Formata_Dados(strDtFinal, tpDados_DataHora, tpNulo_NaoAceita) & _
   " Order By ESTOQUEINTERMEDIARIO.Descricao,  GRUPOESTOQUE.DESCRICAO"
   
  '
  Set objRsRetEst = clsGer.ExecutarSQL(strSql)
  '
  objRsRetEst.Filter = strWhere
  '
  strSql = "Select GRUPOESTOQUE.DESCRICAO AS DESCGRUPOEST, ESTOQUEINTERMEDIARIO.CODIGO, TAB_BAIXAESTOQUEINTERMEDIARIO.QUANTIDADE " & _
           "From ((((ESTOQUE INNER JOIN ESTOQUEINTERMEDIARIO ON ESTOQUE.PKID = ESTOQUEINTERMEDIARIO.ESTOQUEID) " & _
           "INNER JOIN TAB_GRUPOESTESTINTER ON ESTOQUEINTERMEDIARIO.PKID = TAB_GRUPOESTESTINTER.ESTOQUEINTERMEDIARIOID) " & _
           "INNER JOIN TAB_BAIXAESTOQUEINTERMEDIARIO ON TAB_GRUPOESTESTINTER.PKID = TAB_BAIXAESTOQUEINTERMEDIARIO.TAB_GRUPOESTESTINTERID) " & _
           "INNER JOIN BAIXAESTOQUE ON BAIXAESTOQUE.PKID = TAB_BAIXAESTOQUEINTERMEDIARIO.BAIXAESTOQUEID) " & _
           "INNER JOIN GRUPOESTOQUE ON GRUPOESTOQUE.PKID = TAB_GRUPOESTESTINTER.GRUPOESTOQUEID " & _
           " WHERE BAIXAESTOQUE.TIPO = " & Formata_Dados("I", tpDados_Texto, tpNulo_NaoAceita) & _
           " AND BAIXAESTOQUE.DATA >= " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
           " AND BAIXAESTOQUE.DATA <= " & Formata_Dados(strDtFinal, tpDados_DataHora, tpNulo_NaoAceita) & _
           " Order By ESTOQUEINTERMEDIARIO.Descricao,  GRUPOESTOQUE.DESCRICAO"
  '
  Set objRsBaixaEst = clsGer.ExecutarSQL(strSql)
  '
  objRsBaixaEst.Filter = strWhere
  '
  strSql = "Select GRUPOESTOQUE.DESCRICAO AS DESCGRUPOEST, ESTOQUEINTERMEDIARIO.CODIGO, TAB_RETORNOESTOQUEINTERMEDIARIO.QUANTIDADE " & _
           "From ((((ESTOQUE INNER JOIN ESTOQUEINTERMEDIARIO ON ESTOQUE.PKID = ESTOQUEINTERMEDIARIO.ESTOQUEID) " & _
           "INNER JOIN TAB_GRUPOESTESTINTER ON ESTOQUEINTERMEDIARIO.PKID = TAB_GRUPOESTESTINTER.ESTOQUEINTERMEDIARIOID) " & _
           "INNER JOIN TAB_RETORNOESTOQUEINTERMEDIARIO ON TAB_GRUPOESTESTINTER.PKID = TAB_RETORNOESTOQUEINTERMEDIARIO.TAB_GRUPOESTESTINTERID) " & _
           "INNER JOIN RETORNOESTOQUE ON RETORNOESTOQUE.PKID = TAB_RETORNOESTOQUEINTERMEDIARIO.RETORNOESTOQUEID) " & _
           "INNER JOIN GRUPOESTOQUE ON GRUPOESTOQUE.PKID = TAB_GRUPOESTESTINTER.GRUPOESTOQUEID " & _
           " WHERE RETORNOESTOQUE.TIPO = " & Formata_Dados("I", tpDados_Texto, tpNulo_NaoAceita) & _
           " AND RETORNOESTOQUE.DATA >= " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
           " AND RETORNOESTOQUE.DATA <= " & Formata_Dados(strDtFinal, tpDados_DataHora, tpNulo_NaoAceita) & _
           " Order By ESTOQUEINTERMEDIARIO.Descricao,  GRUPOESTOQUE.DESCRICAO"
  '
  Set objRsRetDepEst = clsGer.ExecutarSQL(strSql)
  '
  objRsRetDepEst.Filter = strWhere
  '
  strSql = "Select GRUPOESTOQUE.DESCRICAO AS DESCGRUPOEST, ESTOQUEINTERMEDIARIO.CODIGO, TAB_TRANSFESTINTER.QUANTIDADE " & _
           "From ((((ESTOQUE INNER JOIN ESTOQUEINTERMEDIARIO ON ESTOQUE.PKID = ESTOQUEINTERMEDIARIO.ESTOQUEID) " & _
           "INNER JOIN TAB_GRUPOESTESTINTER ON ESTOQUEINTERMEDIARIO.PKID = TAB_GRUPOESTESTINTER.ESTOQUEINTERMEDIARIOID) " & _
           "INNER JOIN TAB_TRANSFESTINTER ON TAB_GRUPOESTESTINTER.PKID = TAB_TRANSFESTINTER.TAB_GRUPOESTESTINTERENTRADAID) " & _
           "INNER JOIN TRANSFESTINTER ON TRANSFESTINTER.PKID = TAB_TRANSFESTINTER.TRANSFESTINTERID) " & _
           "INNER JOIN GRUPOESTOQUE ON GRUPOESTOQUE.PKID = TAB_GRUPOESTESTINTER.GRUPOESTOQUEID " & _
           " WHERE TRANSFESTINTER.DATA >= " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
           " AND TRANSFESTINTER.DATA <= " & Formata_Dados(strDtFinal, tpDados_DataHora, tpNulo_NaoAceita) & _
           " Order By ESTOQUEINTERMEDIARIO.Descricao,  GRUPOESTOQUE.DESCRICAO"
  '
  Set objRsTransfEntrada = clsGer.ExecutarSQL(strSql)
  '
  objRsTransfEntrada.Filter = strWhere
  '
  strSql = "Select GRUPOESTOQUE.DESCRICAO AS DESCGRUPOEST, ESTOQUEINTERMEDIARIO.CODIGO, TAB_TRANSFESTINTER.QUANTIDADE " & _
           "From ((((ESTOQUE INNER JOIN ESTOQUEINTERMEDIARIO ON ESTOQUE.PKID = ESTOQUEINTERMEDIARIO.ESTOQUEID) " & _
           "INNER JOIN TAB_GRUPOESTESTINTER ON ESTOQUEINTERMEDIARIO.PKID = TAB_GRUPOESTESTINTER.ESTOQUEINTERMEDIARIOID) " & _
           "INNER JOIN TAB_TRANSFESTINTER ON TAB_GRUPOESTESTINTER.PKID = TAB_TRANSFESTINTER.TAB_GRUPOESTESTINTERSAIDAID) " & _
           "INNER JOIN TRANSFESTINTER ON TRANSFESTINTER.PKID = TAB_TRANSFESTINTER.TRANSFESTINTERID) " & _
           "INNER JOIN GRUPOESTOQUE ON GRUPOESTOQUE.PKID = TAB_GRUPOESTESTINTER.GRUPOESTOQUEID " & _
           " WHERE TRANSFESTINTER.DATA >= " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
           " AND TRANSFESTINTER.DATA <= " & Formata_Dados(strDtFinal, tpDados_DataHora, tpNulo_NaoAceita) & _
           " Order By ESTOQUEINTERMEDIARIO.Descricao,  GRUPOESTOQUE.DESCRICAO"
  '
  Set objRsTransfSaida = clsGer.ExecutarSQL(strSql)
  '
  objRsTransfSaida.Filter = strWhere
  '
  strSql = "Select GRUPOESTOQUE.DESCRICAO AS DESCGRUPOEST, ESTOQUEINTERMEDIARIO.CODIGO, (TAB_VENDACARD.QUANTIDADE * TAB_CARDESTINTER.QUANTIDADE) AS QUANTIDADEBAIXA " & _
            "From ((((((VENDA INNER JOIN TAB_VENDACARD ON VENDA.PKID = TAB_VENDACARD.VENDAID) " & _
            "INNER JOIN CARDAPIO ON CARDAPIO.PKID = TAB_VENDACARD.CARDAPIOID) " & _
            "INNER JOIN TAB_CARDESTINTER ON CARDAPIO.PKID = TAB_CARDESTINTER.CARDAPIOID) " & _
            "INNER JOIN ESTOQUEINTERMEDIARIO ON ESTOQUEINTERMEDIARIO.PKID = TAB_CARDESTINTER.ESTOQUEINTERMEDIARIOID) " & _
            "INNER JOIN TAB_GRUPOESTESTINTER ON ESTOQUEINTERMEDIARIO.PKID = TAB_GRUPOESTESTINTER.ESTOQUEINTERMEDIARIOID) " & _
            "INNER JOIN GRUPOESTOQUE ON GRUPOESTOQUE.PKID = TAB_GRUPOESTESTINTER.GRUPOESTOQUEID) " & _
            "Where IIf(IsNull(TAB_CARDESTINTER.Tipo), CARDAPIO.Tipo, TAB_CARDESTINTER.Tipo) = GRUPOESTOQUE.Tipo " & _
            " AND VENDA.DATA >= " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
           " AND VENDA.DATA <= " & Formata_Dados(strDtFinal, tpDados_DataHora, tpNulo_NaoAceita) & _
           " Order By ESTOQUEINTERMEDIARIO.Descricao,  GRUPOESTOQUE.DESCRICAO"
  '
  Set objRsVenda = clsGer.ExecutarSQL(strSql)
  '
  objRsVenda.Filter = strWhere
  '
  strSql = "Select GRUPOESTOQUE.DESCRICAO AS DESCGRUPOEST, ESTOQUEINTERMEDIARIO.CODIGO, (TAB_PEDIDOCARD.QUANTIDADE * TAB_CARDESTINTER.QUANTIDADE) AS QUANTIDADEBAIXA " & _
            "From ((((((((LOCACAO INNER JOIN PEDIDO ON LOCACAO.PKID = PEDIDO.ALOCACAOID) " & _
            "INNER JOIN TAB_PEDIDOCARD ON PEDIDO.PKID = TAB_PEDIDOCARD.PEDIDOID) " & _
            "INNER JOIN CARDAPIO ON CARDAPIO.PKID = TAB_PEDIDOCARD.CARDAPIOID) " & _
            "INNER JOIN TAB_CARDESTINTER ON CARDAPIO.PKID = TAB_CARDESTINTER.CARDAPIOID) " & _
            "INNER JOIN ESTOQUEINTERMEDIARIO ON ESTOQUEINTERMEDIARIO.PKID = TAB_CARDESTINTER.ESTOQUEINTERMEDIARIOID) " & _
            "INNER JOIN TAB_GRUPOESTESTINTER ON ESTOQUEINTERMEDIARIO.PKID = TAB_GRUPOESTESTINTER.ESTOQUEINTERMEDIARIOID) " & _
            "INNER JOIN GRUPOESTOQUE ON GRUPOESTOQUE.PKID = TAB_GRUPOESTESTINTER.GRUPOESTOQUEID) " & _
            "INNER JOIN TAB_GRUPOESTAPTO ON GRUPOESTOQUE.PKID = TAB_GRUPOESTAPTO.GRUPOESTOQUEID) " & _
            "Where IIf(IsNull(TAB_CARDESTINTER.Tipo), CARDAPIO.Tipo, TAB_CARDESTINTER.Tipo) = GRUPOESTOQUE.Tipo " & _
            " AND TAB_GRUPOESTAPTO.APARTAMENTOID = LOCACAO.APARTAMENTOID " & _
            " AND PEDIDO.DTPEDIDO >= " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
            " AND PEDIDO.DTPEDIDO <= " & Formata_Dados(strDtFinal, tpDados_DataHora, tpNulo_NaoAceita) & _
            " Order By ESTOQUEINTERMEDIARIO.Descricao,  GRUPOESTOQUE.DESCRICAO"
  '
  Set objRsPedido = clsGer.ExecutarSQL(strSql)
  '
  objRsPedido.Filter = strWhere
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
        For intJ = 0 To 3 'varre as colunas
          Matriz(intJ, intI) = IIf(objRs(intJ) = 0, "", objRs(intJ) & "")
          If intJ = 3 Then lngQtdAnterior = IIf(Not IsNumeric(objRs(intJ)), 0, objRs(intJ))
        Next
        'Qtd Requisição
        lngQtdRequisicao = 0
        If Not objRsRequisicao.EOF Then   'se já houver algum item em requisição
          Do While objRsRequisicao.Fields("CODIGO").Value & "" = objRs.Fields("CODIGO").Value & "" And _
                   objRsRequisicao.Fields("DESCGRUPOEST").Value & "" = objRs.Fields("DESCGRUPOEST").Value & ""
            lngQtdRequisicao = lngQtdRequisicao + objRsRequisicao.Fields("QUANTIDADEATENDIDA").Value
            objRsRequisicao.MoveNext
            If objRsRequisicao.EOF Then Exit Do
          Loop
        End If
        Matriz(4, intI) = IIf(lngQtdRequisicao = 0, "", lngQtdRequisicao)
        'Qtd Retorno Requisição
        lngQtdRetReq = 0
        If Not objRsRetEst.EOF Then   'se já houver algum item em retorno requisição
          Do While objRsRetEst.Fields("CODIGO").Value & "" = objRs.Fields("CODIGO").Value & "" And _
                   objRsRetEst.Fields("DESCGRUPOEST").Value & "" = objRs.Fields("DESCGRUPOEST").Value & ""
            lngQtdRetReq = lngQtdRetReq + objRsRetEst.Fields("QUANTIDADE").Value
            objRsRetEst.MoveNext
            If objRsRetEst.EOF Then Exit Do
          Loop
        End If
        Matriz(5, intI) = IIf(lngQtdRetReq = 0, "", lngQtdRetReq)
        'Qtd Baixa Estoque
        lngQtdBaixaEst = 0
        If Not objRsBaixaEst.EOF Then   'se já houver algum item em baixa estoque
          Do While objRsBaixaEst.Fields("CODIGO").Value & "" = objRs.Fields("CODIGO").Value & "" And _
                   objRsBaixaEst.Fields("DESCGRUPOEST").Value & "" = objRs.Fields("DESCGRUPOEST").Value & ""
            lngQtdBaixaEst = lngQtdBaixaEst + objRsBaixaEst.Fields("QUANTIDADE").Value
            objRsBaixaEst.MoveNext
            If objRsBaixaEst.EOF Then Exit Do
          Loop
        End If
        Matriz(6, intI) = IIf(lngQtdBaixaEst = 0, "", lngQtdBaixaEst)
        'Qtd Retorno depósito p/ Estoque
        lngQtdRetDepEst = 0
        If Not objRsRetDepEst.EOF Then   'se já houver algum item em retorno depósito p/ estoque
          Do While objRsRetDepEst.Fields("CODIGO").Value & "" = objRs.Fields("CODIGO").Value & "" And _
                   objRsRetDepEst.Fields("DESCGRUPOEST").Value & "" = objRs.Fields("DESCGRUPOEST").Value & ""
            lngQtdRetDepEst = lngQtdRetDepEst + objRsRetDepEst.Fields("QUANTIDADE").Value
            objRsRetDepEst.MoveNext
            If objRsRetDepEst.EOF Then Exit Do
          Loop
        End If
        Matriz(7, intI) = IIf(lngQtdRetDepEst = 0, "", lngQtdRetDepEst)
        'Qtd transferencia entre est inter - entrada
        lngQtdTransfEntrada = 0
        If Not objRsTransfEntrada.EOF Then   'se já houver algum item em transf entre estoques
          Do While objRsTransfEntrada.Fields("CODIGO").Value & "" = objRs.Fields("CODIGO").Value & "" And _
                   objRsTransfEntrada.Fields("DESCGRUPOEST").Value & "" = objRs.Fields("DESCGRUPOEST").Value & ""
            lngQtdTransfEntrada = lngQtdTransfEntrada + objRsTransfEntrada.Fields("QUANTIDADE").Value
            objRsTransfEntrada.MoveNext
            If objRsTransfEntrada.EOF Then Exit Do
          Loop
        End If
        Matriz(8, intI) = IIf(lngQtdTransfEntrada = 0, "", lngQtdTransfEntrada)
        'Qtd transferencia entre est inter - saída
        lngQtdTransfSaida = 0
        If Not objRsTransfSaida.EOF Then   'se já houver algum item em transf entre estoques
          Do While objRsTransfSaida.Fields("CODIGO").Value & "" = objRs.Fields("CODIGO").Value & "" And _
                   objRsTransfSaida.Fields("DESCGRUPOEST").Value & "" = objRs.Fields("DESCGRUPOEST").Value & ""
            lngQtdTransfSaida = lngQtdTransfSaida + objRsTransfSaida.Fields("QUANTIDADE").Value
            objRsTransfSaida.MoveNext
            If objRsTransfSaida.EOF Then Exit Do
          Loop
        End If
        Matriz(9, intI) = IIf(lngQtdTransfSaida = 0, "", lngQtdTransfSaida)
        'Qtd VENDA
        lngQtdVenda = 0
        If Not objRsVenda.EOF Then   'se já houver alguma venda
          Do While objRsVenda.Fields("CODIGO").Value & "" = objRs.Fields("CODIGO").Value & "" And _
                   objRsVenda.Fields("DESCGRUPOEST").Value & "" = objRs.Fields("DESCGRUPOEST").Value & ""
            lngQtdVenda = lngQtdVenda + objRsVenda.Fields("QUANTIDADEBAIXA").Value
            objRsVenda.MoveNext
            If objRsVenda.EOF Then Exit Do
          Loop
        End If
        Matriz(10, intI) = IIf(lngQtdVenda = 0, "", lngQtdVenda)
        'Qtd PEDIDO
        lngQtdPedido = 0
        If Not objRsPedido.EOF Then   'se já houver alguma venda
          Do While objRsPedido.Fields("CODIGO").Value & "" = objRs.Fields("CODIGO").Value & "" And _
                   objRsPedido.Fields("DESCGRUPOEST").Value & "" = objRs.Fields("DESCGRUPOEST").Value & ""
            lngQtdPedido = lngQtdPedido + objRsPedido.Fields("QUANTIDADEBAIXA").Value
            objRsPedido.MoveNext
            If objRsPedido.EOF Then Exit Do
          Loop
        End If
        Matriz(11, intI) = IIf(lngQtdPedido = 0, "", lngQtdPedido)
        'Qtd total movimentada
        lngQtdTotal = lngQtdAnterior + lngQtdRequisicao - lngQtdRetReq - lngQtdBaixaEst + lngQtdRetDepEst + lngQtdTransfEntrada - lngQtdTransfSaida - lngQtdVenda - lngQtdPedido
        Matriz(12, intI) = IIf(lngQtdTotal = 0, "", lngQtdTotal)
        
        'Qtd em estoque
        Matriz(13, intI) = IIf(objRs.Fields("QTDEST").Value = 0, "", objRs.Fields("QTDEST").Value)
        
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set clsGer = Nothing
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Activate()
  If blnPrimeiraVez Then
    DoEvents
    '
    COLUNASMATRIZ = 14
    LINHASMATRIZ = 0
  
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = LINHASMATRIZ
    blnPrimeiraVez = False
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ReDim Matriz(0 To 0, 0 To 0)
  LINHASMATRIZ = 0
  COLUNASMATRIZ = 0
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
  TratarErro Err.Number, Err.Description, "[frmUserPedidoLis.grdGeral_UnboundReadDataEx]"
End Sub


Private Sub Form_Load()
On Error GoTo trata
  AmpS
  Me.Height = 6450
  Me.Width = 11535
  blnPrimeiraVez = True
  
  CenterForm Me
  PreencheCombo cboEstInter, "SELECT DESCRICAO FROM  GRUPOESTOQUE ORDER BY DESCRICAO"
  
  Me.Caption = Me.Caption
  LerFiguras Me, tpBmp_Vazio, , , cmdFechar, , , , , cmdFiltrar, cmdImprimir
  '
  AmpN
  Exit Sub
trata:
  AmpN
  TratarErro Err.Number, Err.Description, Err.Source
End Sub
