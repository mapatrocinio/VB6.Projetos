VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItemPedidoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirItemPedidoVenda(lngPEDIDOVENDAID As Long, _
                                 lngPRODUTOID As Long, _
                                 strQuantidade As String, _
                                 strValor As String, _
                                 strValorInstalacao As String, _
                                 strValorFrete As String, _
                                 strValorComissVend As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO ITEM_PEDIDOVENDA " & _
      "(PEDIDOVENDAID, PRODUTOID, QUANTIDADE, VALOR, VALOR_INSTALACAO, VALOR_FRETE, VALOR_COMISSAO) VALUES (" & _
      Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo) & _
      "," & Formata_Dados(lngPRODUTOID, tpDados_Longo) & _
      "," & Formata_Dados(strQuantidade, tpDados_Moeda) & _
      "," & Formata_Dados(strValor, tpDados_Moeda) & _
      "," & Formata_Dados(strValorInstalacao, tpDados_Moeda) & _
      "," & Formata_Dados(strValorFrete, tpDados_Moeda) & _
      "," & Formata_Dados(strValorComissVend, tpDados_Moeda) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemOS.InserirItemPedidoVenda]", _
            Err.Description
End Sub


Public Function ExcluirItemPedidoVenda(lngItemPedidoId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM ITEM_PEDIDOVENDA " & _
    "WHERE PKID = " & Formata_Dados(lngItemPedidoId, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemPedido.ExcluirItemPedidoVenda]", _
            Err.Description
End Function

Public Function SelecionarItemPedidoVendaPeloPkid(lngItemPedidoId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ITEM_PEDIDOVENDA.*, PRODUTO.NOME AS NOME_PRODUTO, INSUMO.CODIGO AS CODIGO_PRODUTO " & _
    "FROM ITEM_PEDIDOVENDA " & _
    " INNER JOIN PRODUTO ON PRODUTO.INSUMOID = ITEM_PEDIDOVENDA.PRODUTOID " & _
    " INNER JOIN INSUMO ON INSUMO.PKID = PRODUTO.INSUMOID " & _
    "WHERE ITEM_PEDIDOVENDA.PKID = " & _
    Formata_Dados(lngItemPedidoId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarItemPedidoVendaPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemPedido.SelecionarItemPedidoVendaPeloPkid]", _
            Err.Description
End Function

Public Function ListarItemPedidoVenda(lngPEDIDOVENDAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ITEM_PEDIDOVENDA.*, PRODUTO.NOME AS NOME_PRODUTO, INSUMO.CODIGO AS CODIGO_PRODUTO " & _
    "FROM ITEM_PEDIDOVENDA " & _
    " INNER JOIN PRODUTO ON PRODUTO.INSUMOID = ITEM_PEDIDOVENDA.PRODUTOID " & _
    " INNER JOIN INSUMO ON INSUMO.PKID = PRODUTO.INSUMOID " & _
    "WHERE ITEM_PEDIDOVENDA.PEDIDOVENDAID = " & _
    Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo)
  Set ListarItemPedidoVenda = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemPedidoVenda.ListarItemPedidoVenda]", _
            Err.Description
End Function

Public Sub AlterarItemPedidoVenda(lngPKID As Long, _
                                  strQuantidade As String, _
                                  strValor As String, _
                                  strValorInstalacao As String, _
                                  strValorFrete As String, _
                                  strValorComissVend As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE ITEM_PEDIDOVENDA SET " & _
    " QUANTIDADE = " & Formata_Dados(strQuantidade, tpDados_Moeda) & _
    ", VALOR = " & Formata_Dados(strValor, tpDados_Moeda) & _
    ", VALOR_INSTALACAO = " & Formata_Dados(strValorInstalacao, tpDados_Moeda) & _
    ", VALOR_FRETE = " & Formata_Dados(strValorFrete, tpDados_Moeda) & _
    ", VALOR_COMISSAO = " & Formata_Dados(strValorComissVend, tpDados_Moeda) & _
    " WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemPedido.AlterarItemPedidoVenda]", _
            Err.Description
End Sub
