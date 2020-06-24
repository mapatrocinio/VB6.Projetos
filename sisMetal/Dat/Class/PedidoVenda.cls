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

Public Sub BaixaEstoqueProduto(lngPEDIDOVENDAID As Long)
  On Error GoTo trata
  
  '
  Dim strSql As String
  '
  strSql = "EXEC SP_BAIXA_ESTOQUE_PRODUTO " & Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedidoVenda.BaixaEstoqueProduto]", _
            Err.Description
End Sub



Public Function CalculaTotaisVenda(lngPEDIDOVENDAID As Long, _
                                   strTotoal As String, _
                                   strComissao As String, _
                                   strDesconto As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE PEDIDOVENDA SET " & _
    " VALOR_CALC_TOTAL = " & Formata_Dados(strTotoal, tpDados_Moeda) & _
    ", VALOR_CALC_COMISS = " & Formata_Dados(strComissao, tpDados_Moeda) & _
    ", VALOR_CALC_DESCONTO = " & Formata_Dados(strDesconto, tpDados_Moeda) & _
    " WHERE PEDIDOVENDA.PKID = " & Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo)
  '
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedidoVenda.CalculaTotaisVenda]", _
            Err.Description
End Function

Public Function SelecionarPedidoVenda(lngPEDIDOVENDAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  'strSql = "SELECT ISNULL(ITEM.VR_TOT_PEDIDO , 0) AS VR_TOT_PEDIDO, " & _
      " CASE ISNULL(PEDIDOVENDA.TIPO_DESCONTO , 'I') " & _
      " WHEN 'P' THEN " & _
      " ISNULL(ITEM.VR_TOT_PEDIDO , 0) * ISNULL(PEDIDOVENDA.VALOR_DESCONTO , 0) / 100 " & _
      " WHEN 'V' THEN " & _
      " ISNULL(ITEM.VR_TOT_PEDIDO , 0) " & _
      " Else 0 End " & _
      " AS VRCALCDESCONTO, " & _
      " (ISNULL(ITEM.VR_TOT_PEDIDO , 0)) - " & _
      " CASE ISNULL(PEDIDOVENDA.TIPO_DESCONTO , 'I') " & _
      " WHEN 'P' THEN " & _
      " ISNULL(ITEM.VR_TOT_PEDIDO , 0) * ISNULL(PEDIDOVENDA.VALOR_DESCONTO , 0) / 100 " & _
      "  WHEN 'V' THEN " & _
      " ISNULL(ITEM.VR_TOT_PEDIDO , 0) " & _
      " Else 0 End " & _
      " AS VRCALCTOTAL, " & _
      " PEDIDOVENDA.* " & _
      " From PEDIDOVENDA " & _
      " LEFT JOIN (SELECT PEDIDOVENDAID, " & _
      " ISNULL(SUM(VALOR),0) + ISNULL(SUM(VALOR_INSTALACAO),0) + ISNULL(SUM(VALOR_FRETE),0) AS VR_TOT_PEDIDO " & _
      " From ITEM_PEDIDOVENDA " & _
      " GROUP BY PEDIDOVENDAID) AS ITEM " & _
      " ON ITEM.PEDIDOVENDAID = PEDIDOVENDA.PKID " & _
      " WHERE PEDIDOVENDA.PKID = " & _
      Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo)
  strSql = "SELECT PEDIDOVENDA.*, VRCALCTOTAL = ISNULL(VALOR_CALC_TOTAL,0) - ISNULL(VALOR_CALC_DESCONTO,0) " & _
      " From PEDIDOVENDA " & _
      " WHERE PEDIDOVENDA.PKID = " & _
      Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo)

  Set SelecionarPedidoVenda = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedidoVenda.SelecionarPedidoVenda]", _
            Err.Description
End Function

Public Sub AlterarStatusPedidoVenda(lngPEDIDOVENDAID As Long, _
                                    strStatus As String)
  On Error GoTo trata
  
  '
  Dim strSql As String

  strSql = "UPDATE PEDIDOVENDA " & _
    " SET STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & _
    " WHERE PKID = " & _
      Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedidoVenda.AlterarStatusPedidoVenda]", _
            Err.Description
End Sub

Public Function ListarPedidoVenda(lngPEDIDOVENDAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT PEDIDOVENDA.*, TIPOVENDA.DESCRICAO AS DESC_TIPOVENDA, " & _
    " FICHACLIENTE.PKID AS PKID_FICHACLIENTE, FICHACLIENTE.NOME AS NOME_FICHACLIENTE, LOJA.NOME AS NOME_EMRPESA, LOJA.PKID AS PKID_EMRPESA, " & _
    " PESSOA.NOME AS NOME_VENDEDOR " & _
    " FROM PEDIDOVENDA " & _
    " LEFT JOIN TIPOVENDA ON TIPOVENDA.PKID = PEDIDOVENDA.TIPOVENDAID " & _
    " LEFT JOIN FICHACLIENTE ON FICHACLIENTE.PKID = PEDIDOVENDA.FICHACLIENTEID " & _
    " LEFT JOIN LOJA ON LOJA.PKID = PEDIDOVENDA.EMPRESAID " & _
    " LEFT JOIN PESSOA ON PESSOA.PKID = PEDIDOVENDA.VENDEDORID " & _
    " WHERE PEDIDOVENDA.PKID = " & _
    Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarPedidoVenda = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedidoVenda.ListarPedidoVenda]", _
            Err.Description
End Function

Public Function ListarPedidoVendaTotal(lngPEDIDOVENDAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT PEDIDOVENDA.PKID, " & _
    " FROM PEDIDOVENDA " & _
    " LEFT JOIN TIPOVENDA ON TIPOVENDA.PKID = PEDIDOVENDA.TIPOVENDAID " & _
    " LEFT JOIN FICHACLIENTE ON FICHACLIENTE.PKID = PEDIDOVENDA.FICHACLIENTEID " & _
    " LEFT JOIN LOJA ON LOJA.PKID = PEDIDOVENDA.EMPRESAID " & _
    " LEFT JOIN PESSOA ON PESSOA.PKID = PEDIDOVENDA.VENDEDORID " & _
    " WHERE PEDIDOVENDA.PKID = " & _
    Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarPedidoVendaTotal = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedidoVenda.ListarPedidoVendaTotal]", _
            Err.Description
End Function

'''Public Function ListarPedidoVendaTotalAberto(lngPEDIDOVENDAID As Long) As ADODB.Recordset
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "SELECT PEDIDOVENDA.* " & _
'''    " FROM PEDIDOVENDA " & _
'''    " INNER JOIN ITEM_PEDIDOVENDA ON PEDIDOVENDA.PKID = ITEM_PEDIDOVENDA.PEDIDOVENDAID " & _
'''    " WHERE PEDIDOVENDA.PKID = " & _
'''    Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo, tpNulo_NaoAceita) & _
'''    " AND ISNULL(PESO_INI,0) <> (ISNULL(PESO,0) + ISNULL(PESO_FAB,0)) "
'''  Set ListarPedidoVendaTotalAberto = RunSPReturnRS(strSql)
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsPedidoVenda.ListarPedidoVendaTotalAberto]", _
'''            Err.Description
'''End Function
'''

Public Sub InserirPedidoVenda(lngPEDIDOVENDAID As Long, _
                              strData As String, _
                              strStatus As String, _
                              strNumero As String, _
                              lngVENDEDORID As Long, _
                              lngFichaClienteId As Long, _
                              lngTIPOVENDAID As Long, _
                              strValorDesconto As String, _
                              strTipoDesconto As String, _
                              lngEMPRESAID As Long)
  On Error GoTo trata
  '
  Dim objRs As ADODB.Recordset
  Dim strSql As String
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO PEDIDOVENDA (DATA, STATUS, PED_NUMERO, VENDEDORID, FICHACLIENTEID, TIPOVENDAID, VALOR_DESCONTO, TIPO_DESCONTO, EMPRESAID) VALUES (" & _
    Formata_Dados(strData, tpDados_DataHora) & ", " & _
    Formata_Dados(strStatus, tpDados_Texto) & ", " & _
    Formata_Dados(strNumero, tpDados_Longo) & ", " & _
    Formata_Dados(IIf(lngVENDEDORID = 0, "", lngVENDEDORID), tpDados_Longo) & ", " & _
    Formata_Dados(IIf(lngFichaClienteId = 0, "", lngFichaClienteId), tpDados_Longo) & ", " & _
    Formata_Dados(lngTIPOVENDAID, tpDados_Longo) & ", " & _
    Formata_Dados(strValorDesconto, tpDados_Moeda) & ", " & _
    Formata_Dados(strTipoDesconto, tpDados_Texto) & ", " & _
    Formata_Dados(IIf(lngEMPRESAID = 0, "", lngEMPRESAID), tpDados_Longo) & ");"
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngPEDIDOVENDAID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedidoVenda.InserirPedidoVenda]", _
            Err.Description
End Sub

Public Sub AlterarPedidoVendaCaixa(lngPEDIDOVENDAID As Long, _
                                   lngCAIXAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE PEDIDOVENDA SET " & _
    " CAIXAID = " & Formata_Dados(lngCAIXAID, tpDados_Longo) & _
    " WHERE PEDIDOVENDA.PKID = " & Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo)
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedidoVenda.AlterarPedidoVendaCaixa]", _
            Err.Description
End Sub

Public Sub AlterarPedidoVenda(lngPEDIDOVENDAID As Long, _
                              lngFichaClienteId As Long, _
                              strValorDesconto As String, _
                              strTipoDesconto As String, _
                              lngEMPRESAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE PEDIDOVENDA SET " & _
    " FICHACLIENTEID = " & Formata_Dados(IIf(lngFichaClienteId = 0, "", lngFichaClienteId), tpDados_Longo) & _
    ", EMPRESAID = " & Formata_Dados(IIf(lngEMPRESAID = 0, "", lngEMPRESAID), tpDados_Longo) & _
    ", VALOR_DESCONTO = " & Formata_Dados(strValorDesconto, tpDados_Moeda) & _
    ", TIPO_DESCONTO = " & Formata_Dados(strTipoDesconto, tpDados_Texto) & _
    " WHERE PEDIDOVENDA.PKID = " & Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo)
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedidoVenda.AlterarPedidoVenda]", _
            Err.Description
End Sub

Public Function CapturarSequencialPedidoVenda(strData As String, _
                                              strSequencial As String) As Long
  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "DECLARE @DATA AS SMALLDATETIME; "
  strSql = strSql & "SET @DATA = convert(datetime, '" & strData & "',103); "
  strSql = strSql & "EXEC SP_SEQUENCIAL_PEDIDO_VENDA @DATA;"
  '
  Set objRs = RunSPReturnRS(strSql)
  'Não pode haver erro
  strSequencial = objRs.Fields("SEQUENCIAL").Value
  objRs.Close
  Set objRs = Nothing
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedidoVenda.CapturarSequencialPedidoVenda]", _
            Err.Description
End Function

'''Public Sub FecharVenda(lngPEDIDOVENDAID As Long, _
'''                       strStatus As String)
'''
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "UPDATE PEDIDOVENDA SET "
'''  strSql = strSql & " STATUS = " & Formata_Dados(strStatus, tpDados_Texto)
'''  '
'''  strSql = strSql & " WHERE PKID = " & Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo)
'''  RunSP strSql
'''  '
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsPedidoVenda.FecharVenda]", _
'''            Err.Description
'''End Sub
'''
'''Public Sub AlterarStatusFornecedor(lngPEDIDOVENDAID As Long, _
'''                                   strStatus As String, _
'''                                   strData As String)
'''
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "UPDATE PEDIDOVENDA SET "
'''  strSql = strSql & " STATUS = " & Formata_Dados(strStatus, tpDados_Texto)
'''  strSql = strSql & ", DATA_ENVIO_FORN = " & Formata_Dados(strData, tpDados_DataHora)
'''  strSql = strSql & " WHERE PKID = " & Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo)
'''  RunSP strSql
'''  '
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsPedidoVenda.AlterarStatusFornecedor]", _
'''            Err.Description
'''End Sub
'''
'''
'''
