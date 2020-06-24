VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItemAjuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirItemAjuste(lngITEM_AJUSTEID As Long, _
                             lngAJUSTEID As Long, _
                             lngPERFILID As Long, _
                             strQuantidade As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "EXEC SP_INSERE_ITEM_AJUSTE " & Formata_Dados(lngITEM_AJUSTEID, tpDados_Longo) & _
      "," & Formata_Dados(lngAJUSTEID, tpDados_Longo) & _
      "," & Formata_Dados(lngPERFILID, tpDados_Longo) & _
      "," & Formata_Dados(strQuantidade, tpDados_Moeda)
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemOS.InserirItemAjuste]", _
            Err.Description
End Sub
'''
'''
'''Public Function CalculoQuantidadePedido(lngLINHAID As Long, _
'''                                        strPeso As String) As Long
'''
'''  On Error GoTo trata
'''  Dim objGeral            As datSisMetal.clsGeral
'''  Dim objRs               As ADODB.Recordset
'''  Dim lngQuantidade       As Long
'''  Dim strSql              As String
'''  'Obter peso da linha
'''  Set objGeral = New datSisMetal.clsGeral
'''  lngQuantidade = 0
'''  strSql = "declare @peso as INT; "
'''  strSql = strSql & "exec @peso = dbo.UFN_CALCULA_QUANTIDADE " & _
'''    Formata_Dados(lngLINHAID, tpDados_Longo) & _
'''    "," & Formata_Dados(strPeso, tpDados_Moeda) & _
'''    "; "
'''  strSql = strSql & "select @peso AS QUANTIDADE;"
'''  Set objRs = objGeral.ExecutarSQL(strSql)
'''
'''  If Not objRs.EOF Then
'''    If IsNumeric(objRs.Fields("QUANTIDADE").Value) Then
'''      lngQuantidade = objRs.Fields("QUANTIDADE").Value
'''    End If
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
'''  '
'''  CalculoQuantidadePedido = lngQuantidade
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsItemPedido.CalculoQuantidadePedido]", _
'''            Err.Description
'''End Function
'''
'''Public Function ExcluirItemPedido(lngItemPedidoId As Long) As ADODB.Recordset
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "EXEC SP_EXCLUI_ITEM_AJUSTE " & Formata_Dados(lngItemPedidoId, tpDados_Longo)
'''  RunSP strSql
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsItemPedido.ExcluirItemPedido]", _
'''            Err.Description
'''End Function
'''
'''Public Function SelecionarItemPedidoPeloPkid(lngItemPedidoId As Long) As ADODB.Recordset
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "SELECT ITEM_AJUSTE.*, TIPO_LINHA.NOME AS NOME_LINHA, LINHA.CODIGO AS CODIGO_LINHA " & _
'''    "FROM ITEM_AJUSTE " & _
'''    " INNER JOIN LINHA ON LINHA.PKID = ITEM_AJUSTE.LINHAID " & _
'''    " INNER JOIN TIPO_LINHA ON TIPO_LINHA.PKID = LINHA.TIPO_LINHAID " & _
'''    "WHERE ITEM_AJUSTE.PKID = " & _
'''    Formata_Dados(lngItemPedidoId, tpDados_Longo, tpNulo_Aceita)
'''  Set SelecionarItemPedidoPeloPkid = RunSPReturnRS(strSql)
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsItemPedido.SelecionarItemPedidoPeloPkid]", _
'''            Err.Description
'''End Function
'''
'''
'''Public Sub AlterarItemPedido(lngPKID As Long, _
'''                             lngLINHAID As Long, _
'''                             strQuantidade As String, _
'''                             strPeso As String, _
'''                             strComprimentoVara As String)
'''
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "UPDATE ITEM_AJUSTE SET " & _
'''    "LINHAID = " & Formata_Dados(lngLINHAID, tpDados_Longo) & _
'''    ", QUANTIDADE = " & Formata_Dados(strQuantidade, tpDados_Moeda) & _
'''    ", PESO = " & Formata_Dados(strPeso, tpDados_Moeda) & _
'''    ", COMPRIMENTO_VARA = " & Formata_Dados(strComprimentoVara, tpDados_Moeda) & _
'''    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
'''  RunSP strSql
'''  '
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsItemPedido.AlterarItemPedido]", _
'''            Err.Description
'''End Sub
'''
'''Public Sub InserirItemPedido(lngAJUSTEID As Long, _
'''                             lngLINHAID As Long, _
'''                             strQuantidade As String, _
'''                             strPeso As String, _
'''                             strComprimentoVara As String)
'''
'''  On Error GoTo trata
'''  Dim strSql As String
'''  '
'''  strSql = "INSERT INTO ITEM_AJUSTE (AJUSTEID, LINHAID, QUANTIDADE, PESO, COMPRIMENTO_VARA) VALUES (" & _
'''    Formata_Dados(lngAJUSTEID, tpDados_Longo) & ", " & _
'''    Formata_Dados(lngLINHAID, tpDados_Longo) & ", " & _
'''    Formata_Dados(strQuantidade, tpDados_Moeda) & ", " & _
'''    Formata_Dados(strPeso, tpDados_Moeda) & ", " & _
'''    Formata_Dados(strComprimentoVara, tpDados_Moeda) & " " & ") "
'''  '
'''  RunSP strSql
'''  '
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsItemPedido.InserirItemPedido]", _
'''            Err.Description
'''End Sub
'''
'''Public Function SelecionarItemPedido(lngPKID As Long) As ADODB.Recordset
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "SELECT ITEM_AJUSTE.* " & _
'''    "FROM ITEM_AJUSTE " & _
'''    "WHERE ITEM_AJUSTE.PKID = " & _
'''    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
'''  Set SelecionarItemPedido = RunSPReturnRS(strSql)
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsItemPedido.SelecionarItemPedido]", _
'''            Err.Description
'''End Function
'''
