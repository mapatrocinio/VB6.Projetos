VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPedidoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function CalculaTotaisVenda(lngPEDIDOVENDAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objPedidoVenda  As datSisMetal.clsPedidoVenda
  Dim objGeral        As datSisMetal.clsGeral
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim vrTotoal        As Currency
  Dim vrComissao      As Currency
  Dim vrDesconto      As Currency
  Dim PercVenda       As Currency
  'Capturar valores
  Set objGeral = New datSisMetal.clsGeral
  vrTotoal = 0
  vrComissao = 0
  vrDesconto = 0
  PercVenda = 0
  strSql = "SELECT ISNULL(ITEM.VR_TOT_PEDIDO , 0) AS VR_TOT_PEDIDO, " & _
      " CASE ISNULL(PEDIDOVENDA.TIPO_DESCONTO , 'I') " & _
      " WHEN 'P' THEN " & _
      " ISNULL(ITEM.VR_TOT_PEDIDO , 0) * ISNULL(PEDIDOVENDA.VALOR_DESCONTO , 0) / 100 " & _
      " WHEN 'V' THEN " & _
      " ISNULL(ITEM.VR_TOT_PEDIDO , 0) " & _
      " Else 0 End " & _
      " AS VRCALCDESCONTO, " & _
      " (ISNULL(ITEM.VR_TOT_PEDIDO , 0)) As VRCALCTOTAL, " & _
      " FUNCIONARIO.PERC_VENDA, " & _
      " PEDIDOVENDA.* " & _
      " From PEDIDOVENDA " & _
      " LEFT JOIN (SELECT PEDIDOVENDAID, " & _
      " ISNULL(SUM(VALOR),0) + ISNULL(SUM(VALOR_INSTALACAO),0) + ISNULL(SUM(VALOR_FRETE),0) AS VR_TOT_PEDIDO " & _
      " From ITEM_PEDIDOVENDA " & _
      " GROUP BY PEDIDOVENDAID) AS ITEM " & _
      " ON ITEM.PEDIDOVENDAID = PEDIDOVENDA.PKID " & _
      " LEFT JOIN PESSOA ON PEDIDOVENDA.VENDEDORID = PESSOA.PKID " & _
      " LEFT JOIN FUNCIONARIO ON PESSOA.PKID = FUNCIONARIO.PESSOAID " & _
      " WHERE PEDIDOVENDA.PKID = " & _
      Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo)

  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    If Not IsNull(objRs.Fields("VRCALCTOTAL").Value) Then
      vrTotoal = objRs.Fields("VRCALCTOTAL").Value
    End If
    If Not IsNull(objRs.Fields("VRCALCDESCONTO").Value) Then
      vrDesconto = objRs.Fields("VRCALCDESCONTO").Value
    End If
    If Not IsNull(objRs.Fields("PERC_VENDA").Value) Then
      PercVenda = objRs.Fields("PERC_VENDA").Value
    End If
  End If
  'Calculo da comissão
  vrComissao = (vrTotoal - vrDesconto) * PercVenda / 100
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  '------------------------------
  'Capturar valor total da venda
  Set objPedidoVenda = New datSisMetal.clsPedidoVenda
  'CalculaTotaisVenda
  Set CalculaTotaisVenda = objPedidoVenda.CalculaTotaisVenda(lngPEDIDOVENDAID, _
                                                             Format(vrTotoal, "###,##0.000"), _
                                                             Format(vrComissao, "###,##0.000"), _
                                                             Format(vrDesconto, "###,##0.000"))
  '
  Set objPedidoVenda = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarPedidoVenda(lngPEDIDOVENDAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objPedidoVenda As datSisMetal.clsPedidoVenda
  Set objPedidoVenda = New datSisMetal.clsPedidoVenda
  
  'Listar PedidoVenda
  Set SelecionarPedidoVenda = objPedidoVenda.SelecionarPedidoVenda(lngPEDIDOVENDAID)
  '
  Set objPedidoVenda = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub BaixaEstoqueProduto(ByVal lngPEDIDOVENDAID As Long)
  On Error GoTo trata
  '
  Dim objPedidoVenda          As datSisMetal.clsPedidoVenda
  '
  '
  Set objPedidoVenda = New datSisMetal.clsPedidoVenda
  '
  'Excluir PedidoVenda
  objPedidoVenda.BaixaEstoqueProduto lngPEDIDOVENDAID
  '
  Set objPedidoVenda = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AtivaInativaPedidoVenda(ByVal lngPEDIDOVENDAID As Long, _
                                   ByVal strStatus As String)
  On Error GoTo trata
  '
  Dim objPedidoVenda          As datSisMetal.clsPedidoVenda
'''  Dim objItemPedidoVenda      As datSisMetal.clsItemPedidoVenda
  Dim objRs                   As ADODB.Recordset
  Dim strStatusAtivacao       As String
  '
  If strStatus = "C" Or strStatus = "R" Then
    'Pedido venda cancelado, ativar
    'Verifica qual status ativar
    strStatusAtivacao = ""
    Set objRs = ListarPedidoVenda(lngPEDIDOVENDAID)
    If Not objRs.EOF Then
      Select Case objRs.Fields("DESC_TIPOVENDA").Value & ""
      Case "BALCÃO": strStatusAtivacao = "B"
      Case "CLIENTE": strStatusAtivacao = "L"
      Case "EMPRESA": strStatusAtivacao = "E"
      End Select
    End If
    objRs.Close
    Set objRs = Nothing
    '
    If strStatus = "R" Then strStatusAtivacao = "R"
    '
    Set objPedidoVenda = New datSisMetal.clsPedidoVenda
    '
    'Excluir PedidoVenda
    objPedidoVenda.AlterarStatusPedidoVenda lngPEDIDOVENDAID, _
                                            strStatusAtivacao
    '
    '
    Set objPedidoVenda = Nothing
  Else
    'Pedido venda ativo, cancelar
    ExcluirPedidoVenda lngPEDIDOVENDAID
  End If
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub ExcluirPedidoVenda(ByVal lngPEDIDOVENDAID As Long)
  On Error GoTo trata
  '
  Dim objPedidoVenda          As datSisMetal.clsPedidoVenda
  Dim objItemPedidoVenda      As datSisMetal.clsItemPedidoVenda
  Dim objRs                   As ADODB.Recordset
  Dim strStatus               As String
  '
  'Verifica se há itens lançados
  Set objItemPedidoVenda = New datSisMetal.clsItemPedidoVenda
  Set objRs = objItemPedidoVenda.ListarItemPedidoVenda(lngPEDIDOVENDAID)
  If objRs.EOF Then
    strStatus = "X"
  Else
    strStatus = "C"
  End If
  objRs.Close
  Set objRs = Nothing
  Set objItemPedidoVenda = Nothing
  '
  Set objPedidoVenda = New datSisMetal.clsPedidoVenda
  '
  'Excluir PedidoVenda
  objPedidoVenda.AlterarStatusPedidoVenda lngPEDIDOVENDAID, _
                                          strStatus
  '
  '
  Set objPedidoVenda = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirPedidoVenda(lngPEDIDOVENDAID As Long, _
                              strStatus As String, _
                              lngVENDEDORID As Long, _
                              lngFichaClienteId As Long, _
                              lngTIPOVENDAID As Long, _
                              strValorDesconto As String, _
                              strTipoDesconto As String, _
                              lngEMPRESAID As Long)
  On Error GoTo trata
  Dim objPedidoVenda  As datSisMetal.clsPedidoVenda
  Dim strSequncial    As String
  Dim strData         As String
  Dim strDataSeq      As String
  '
  Set objPedidoVenda = New datSisMetal.clsPedidoVenda
  '
  strData = Format(Now, "DD/MM/YYYY hh:mm")
  strDataSeq = Format(Now, "DD/MM/YYYY")
  'Capturar Sequencial do PedidoVenda
  objPedidoVenda.CapturarSequencialPedidoVenda strDataSeq, _
                                              strSequncial
  'Inserir PedidoVenda
  objPedidoVenda.InserirPedidoVenda lngPEDIDOVENDAID, _
                                    strData, _
                                    strStatus, _
                                    strSequncial, _
                                    lngVENDEDORID, _
                                    lngFichaClienteId, _
                                    lngTIPOVENDAID, _
                                    strValorDesconto, _
                                    strTipoDesconto, _
                                    lngEMPRESAID
  '
  Set objPedidoVenda = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarPedidoVenda(lngPEDIDOVENDAID As Long, _
                              lngFichaClienteId As Long, _
                              strValorDesconto As String, _
                              strTipoDesconto As String, _
                              lngEMPRESAID As Long)
  On Error GoTo trata
  Dim objPedidoVenda          As datSisMetal.clsPedidoVenda
  '
  Set objPedidoVenda = New datSisMetal.clsPedidoVenda
  '
  'Alterar PedidoVenda
  objPedidoVenda.AlterarPedidoVenda lngPEDIDOVENDAID, _
                                    lngFichaClienteId, _
                                    strValorDesconto, _
                                    strTipoDesconto, _
                                    lngEMPRESAID
  '
  Set objPedidoVenda = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub AlterarPedidoVendaCaixa(lngPEDIDOVENDAID As Long, _
                              lngCAIXAID As Long)
  On Error GoTo trata
  Dim objPedidoVenda          As datSisMetal.clsPedidoVenda
  '
  Set objPedidoVenda = New datSisMetal.clsPedidoVenda
  '
  'AlterarPedidoVendaCaixa
  objPedidoVenda.AlterarPedidoVendaCaixa lngPEDIDOVENDAID, _
                                    lngCAIXAID
  '
  Set objPedidoVenda = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarPedidoVenda(lngPEDIDOVENDAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objPedidoVenda As datSisMetal.clsPedidoVenda
  Set objPedidoVenda = New datSisMetal.clsPedidoVenda

  'Listar PedidoVenda
  Set ListarPedidoVenda = objPedidoVenda.ListarPedidoVenda(lngPEDIDOVENDAID)
  '
  Set objPedidoVenda = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub FecharVenda(lngPEDIDOVENDAID As Long, _
                       lngCAIXAID As Long, _
                       blnVendaFechada As Boolean)
  On Error GoTo trata
  Dim objPedidoVenda  As datSisMetal.clsPedidoVenda
  Dim objGeral        As datSisMetal.clsGeral
  Dim strStatus       As String
  Dim vrValorJaPago   As Currency
  Dim vrTotLoc        As Currency
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  '
  Set objPedidoVenda = New datSisMetal.clsPedidoVenda
  '
  '-------------------------------
  '------------------------------
  'verifica se pedido está totalmente pago, se sim dá recebimento,
  'senão crava o status original do pedido
  '------------------------------
  'Capturar valor já pago
  Set objGeral = New datSisMetal.clsGeral
  vrValorJaPago = 0
  strSql = "SELECT SUM(ISNULL(VALOR,0) - ISNULL(VRTROCO,0)) AS VALORJAPAGO " & _
    "FROM CONTACORRENTEPED "
  strSql = strSql & " WHERE PEDIDOVENDAID = " & Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo)

  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    If Not IsNull(objRs.Fields("VALORJAPAGO").Value) Then
      vrValorJaPago = objRs.Fields("VALORJAPAGO").Value
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  '------------------------------
  'Capturar valor total da venda
  '
  Set objPedidoVenda = New datSisMetal.clsPedidoVenda
  vrTotLoc = 0
  Set objRs = objPedidoVenda.SelecionarPedidoVenda(lngPEDIDOVENDAID)
  If Not objRs.EOF Then
    vrTotLoc = IIf(Not IsNumeric(objRs.Fields("VRCALCTOTAL").Value), 0, objRs.Fields("VRCALCTOTAL").Value)
  End If
  '
  objRs.Close
  Set objRs = Nothing
  Set objPedidoVenda = Nothing
  '--------------------------
  'verifica se venda esta fechada
  blnVendaFechada = False
  
  If vrValorJaPago = vrTotLoc Then
    'Está ok, se for recebimento de pedido, fechar venda
    '-------------------------------
    strStatus = "R"
    
    blnVendaFechada = True
  Else
    strStatus = "C"
  End If
  'Alterar Status para Comprador
  AtivaInativaPedidoVenda lngPEDIDOVENDAID, _
                          strStatus
  '
  If strStatus = "R" Then
    'Venda fechada neste caso dá baixa nos ítens do pedido
    BaixaEstoqueProduto lngPEDIDOVENDAID
    'Gravar caixa que registrou o pedido
    AlterarPedidoVendaCaixa lngPEDIDOVENDAID, _
                          lngCAIXAID
  End If
  
  Set objPedidoVenda = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'''Public Sub AlterarStatusFornecedor(lngPEDIDOVENDAID As Long)
'''  On Error GoTo trata
'''  Dim objPedidoVenda As datSisMetal.clsPedidoVenda
'''  Dim strData As String
'''  Dim strStatus As String
'''  Set objPedidoVenda = New datSisMetal.clsPedidoVenda
'''  '
'''  strStatus = "F"
'''  strData = Format(Now, "DD/MM/YYYY hh:mm")
'''  'Alterar Status para Fornecedor
'''  objPedidoVenda.AlterarStatusFornecedor lngPEDIDOVENDAID, _
'''                                    strStatus, _
'''                                    strData
'''  '
'''  Set objPedidoVenda = Nothing
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Sub
'''
'''Public Function ValidaPedidoVendaFechado(lngPEDIDOVENDAID As Long) As Boolean
'''  On Error GoTo trata
'''  Dim objPedidoVenda As datSisMetal.clsPedidoVenda
'''  Dim blnRetorno As Boolean
'''  Dim objRs As ADODB.Recordset
'''  Set objPedidoVenda = New datSisMetal.clsPedidoVenda
'''  blnRetorno = True
'''  'Listar PedidoVenda
'''  Set objRs = objPedidoVenda.ListarPedidoVendaAberto(lngPEDIDOVENDAID)
'''  If Not objRs.EOF Then
'''    blnRetorno = False
'''  End If
'''  '
'''  objRs.Close
'''  Set objRs = Nothing
'''  Set objPedidoVenda = Nothing
'''  ValidaPedidoVendaFechado = blnRetorno
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Function
'''
