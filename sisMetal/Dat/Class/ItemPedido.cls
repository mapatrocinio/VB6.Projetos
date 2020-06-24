VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItemPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirItemPedidoItem(lngITEM_PEDIDOID As Long, _
                                 lngPEDIDOID As Long, _
                                 lngLINHAID As Long, _
                                 strPeso As String, _
                                 strComprimento As String, _
                                 strPesoAnod As String, _
                                 strPesoFab As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "EXEC SP_INSERE_ITEM_PEDIDO " & Formata_Dados(lngITEM_PEDIDOID, tpDados_Longo) & _
      "," & Formata_Dados(lngPEDIDOID, tpDados_Longo) & _
      "," & Formata_Dados(lngLINHAID, tpDados_Longo) & _
      "," & Formata_Dados(strPeso, tpDados_Moeda) & _
      "," & Formata_Dados(strComprimento, tpDados_Moeda) & _
      "," & Formata_Dados(strPesoAnod, tpDados_Moeda) & _
      "," & Formata_Dados(strPesoFab, tpDados_Moeda)
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemOS.InserirItemPedidoItem]", _
            Err.Description
End Sub


Public Function CalculoQuantidadePedido(lngLINHAID As Long, _
                                        strPeso As String) As Long

  On Error GoTo trata
  Dim objGeral            As datSisMetal.clsGeral
  Dim objRs               As ADODB.Recordset
  Dim lngQUANTIDADE       As Long
  Dim strSql              As String
  'Obter peso da linha
  Set objGeral = New datSisMetal.clsGeral
  lngQUANTIDADE = 0
  strSql = "declare @peso as INT; "
  strSql = strSql & "exec @peso = dbo.UFN_CALCULA_QUANTIDADE " & _
    Formata_Dados(lngLINHAID, tpDados_Longo) & _
    "," & Formata_Dados(strPeso, tpDados_Moeda) & _
    "; "
  strSql = strSql & "select @peso AS QUANTIDADE;"
  Set objRs = objGeral.ExecutarSQL(strSql)

  If Not objRs.EOF Then
    If IsNumeric(objRs.Fields("QUANTIDADE").Value) Then
      lngQUANTIDADE = objRs.Fields("QUANTIDADE").Value
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  '
  CalculoQuantidadePedido = lngQUANTIDADE
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemPedido.CalculoQuantidadePedido]", _
            Err.Description
End Function
Public Function CalculoPesoPedido(lngLINHAID As Long, _
                                  strQuantidade As String) As Currency

  On Error GoTo trata
  Dim objGeral            As datSisMetal.clsGeral
  Dim objRs               As ADODB.Recordset
  Dim curPeso             As Currency
  Dim strSql              As String
  'Obter peso da linha
  Set objGeral = New datSisMetal.clsGeral
  curPeso = 0
  strSql = "declare @peso as MONEY; "
  strSql = strSql & "exec @peso = dbo.UFN_CALCULA_PESO " & _
    Formata_Dados(lngLINHAID, tpDados_Longo) & _
    "," & Formata_Dados(strQuantidade, tpDados_Longo) & _
    "; "
  strSql = strSql & "select @peso AS PESO;"
  Set objRs = objGeral.ExecutarSQL(strSql)

  If Not objRs.EOF Then
    If IsNumeric(objRs.Fields("PESO").Value) Then
      curPeso = objRs.Fields("PESO").Value
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  '
  CalculoPesoPedido = curPeso
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemPedido.CalculoPesoPedido]", _
            Err.Description
End Function

Public Function ExcluirItemPedido(lngItemPedidoId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "EXEC SP_EXCLUI_ITEM_PEDIDO " & Formata_Dados(lngItemPedidoId, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemPedido.ExcluirItemPedido]", _
            Err.Description
End Function

Public Function SelecionarItemPedidoPeloPkid(lngItemPedidoId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ITEM_PEDIDO.*, TIPO_LINHA.NOME AS NOME_LINHA, LINHA.CODIGO AS CODIGO_LINHA " & _
    "FROM ITEM_PEDIDO " & _
    " INNER JOIN LINHA ON LINHA.PKID = ITEM_PEDIDO.LINHAID " & _
    " INNER JOIN TIPO_LINHA ON TIPO_LINHA.PKID = LINHA.TIPO_LINHAID " & _
    "WHERE ITEM_PEDIDO.PKID = " & _
    Formata_Dados(lngItemPedidoId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarItemPedidoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemPedido.SelecionarItemPedidoPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarItemPedido(lngPKID As Long, _
                             lngLINHAID As Long, _
                             strQuantidade As String, _
                             strPeso As String, _
                             strComprimentoVara As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE ITEM_PEDIDO SET " & _
    "LINHAID = " & Formata_Dados(lngLINHAID, tpDados_Longo) & _
    ", QUANTIDADE = " & Formata_Dados(strQuantidade, tpDados_Moeda) & _
    ", PESO = " & Formata_Dados(strPeso, tpDados_Moeda) & _
    ", COMPRIMENTO_VARA = " & Formata_Dados(strComprimentoVara, tpDados_Moeda) & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemPedido.AlterarItemPedido]", _
            Err.Description
End Sub

Public Sub InserirItemPedido(lngPEDIDOID As Long, _
                             lngLINHAID As Long, _
                             strQuantidade As String, _
                             strPeso As String, _
                             strComprimentoVara As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO ITEM_PEDIDO (PEDIDOID, LINHAID, QUANTIDADE, PESO, COMPRIMENTO_VARA) VALUES (" & _
    Formata_Dados(lngPEDIDOID, tpDados_Longo) & ", " & _
    Formata_Dados(lngLINHAID, tpDados_Longo) & ", " & _
    Formata_Dados(strQuantidade, tpDados_Moeda) & ", " & _
    Formata_Dados(strPeso, tpDados_Moeda) & ", " & _
    Formata_Dados(strComprimentoVara, tpDados_Moeda) & " " & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemPedido.InserirItemPedido]", _
            Err.Description
End Sub

Public Function SelecionarItemPedido(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ITEM_PEDIDO.* " & _
    "FROM ITEM_PEDIDO " & _
    "WHERE ITEM_PEDIDO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarItemPedido = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemPedido.SelecionarItemPedido]", _
            Err.Description
End Function

