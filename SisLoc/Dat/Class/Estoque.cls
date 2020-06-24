VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ListarEstoquePeloCodigo(strCodigo As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  '
  strSql = "SELECT ESTOQUE.* " & _
    "FROM ESTOQUE " & _
    "WHERE ESTOQUE.CODIGO = " & _
    Formata_Dados(strCodigo, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarEstoquePeloCodigo = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEstoque.ListarEstoquePeloCodigo]", _
            Err.Description
End Function

Public Function CapturaEstoquePeloCodigo(ByVal strCodigo As String) As ADODB.Recordset
  '
  On Error GoTo trata
  Dim strSql  As String
  '
  strSql = "SELECT ESTOQUE.CODIGO, ESTOQUE.DESCRICAO, ESTOQUE.QUANTIDADE, ESTOQUE.* " & _
      "FROM ESTOQUE WHERE CODIGO LIKE " & Formata_Dados(strCodigo & "%", tpDados_Texto, tpNulo_NaoAceita) & _
      " ORDER BY ESTOQUE.DESCRICAO;"
  Set CapturaEstoquePeloCodigo = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEstoque.CapturaEstoquePeloCodigo]", _
            Err.Description
End Function

Public Sub InserirEstoque(lngESTOQUEID As Long, _
                          strCodigo As String, _
                          strDescricao As String, _
                          strUnidadeId As String, _
                          strQuantidade As String, _
                          strValor As String, _
                          strValorIndenizacao As String, _
                          strPeso As String, _
                          strAltura As String, _
                          strLargura As String, _
                          strQuantidadeMin As String)
  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO ESTOQUE (CODIGO, DESCRICAO, UNIDADEID, QUANTIDADE, QTDMINIMA, VALOR, " & _
    "VALORINDENIZACAO, PESO, ALTURA, LARGURA) VALUES (" & _
    Formata_Dados(strCodigo, tpDados_Texto) & ", " & _
    Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    Formata_Dados(strUnidadeId, tpDados_Longo) & ", " & _
    Formata_Dados(strQuantidade, tpDados_Longo) & ", " & _
    Formata_Dados(strQuantidadeMin, tpDados_Longo) & ", " & _
    Formata_Dados(strValor, tpDados_Moeda) & ", " & _
    Formata_Dados(strValorIndenizacao, tpDados_Moeda) & ", " & _
    Formata_Dados(strPeso, tpDados_Moeda) & ", " & _
    Formata_Dados(strAltura, tpDados_Moeda) & ", " & _
    Formata_Dados(strLargura, tpDados_Moeda) & ");"
    
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngESTOQUEID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEstoque.InserirEstoque]", _
            Err.Description
End Sub

Public Sub AlterarQuantidadeEstoque(lngESTOQUEID As Long, _
                                    lngQtdRealEstoque As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'Atualiza QtdEstoque No Estoque
  strSql = "UPDATE ESTOQUE SET QUANTIDADE = " & Formata_Dados(lngQtdRealEstoque, tpDados_Longo) & " " & _
    "WHERE PKID = " & lngESTOQUEID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEstoque.AlterarQuantidadeEstoque]", _
            Err.Description
End Sub

Public Sub AlterarEstoque(lngESTOQUEID As Long, _
                          strCodigo As String, _
                          strDescricao As String, _
                          strUnidadeId As String, _
                          strQuantidade As String, _
                          strValor As String, _
                          strValorIndenizacao As String, _
                          strPeso As String, _
                          strAltura As String, _
                          strLargura As String, _
                          strQuantidadeMin As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE ESTOQUE SET CODIGO = " & Formata_Dados(strCodigo, tpDados_Texto) & ", " & _
    "DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    "UNIDADEID = " & Formata_Dados(strUnidadeId, tpDados_Longo) & ", " & _
    "QUANTIDADE = " & Formata_Dados(strQuantidade, tpDados_Longo) & ", " & _
    "QTDMINIMA = " & Formata_Dados(strQuantidadeMin, tpDados_Longo) & ", " & _
    "VALOR = " & Formata_Dados(strValor, tpDados_Moeda) & ", " & _
    "VALORINDENIZACAO = " & Formata_Dados(strValorIndenizacao, tpDados_Moeda) & ", " & _
    "PESO = " & Formata_Dados(strPeso, tpDados_Moeda) & ", " & _
    "ALTURA = " & Formata_Dados(strAltura, tpDados_Moeda) & ", " & _
    "LARGURA = " & Formata_Dados(strLargura, tpDados_Moeda) & " " & _
    "WHERE PKID = " & lngESTOQUEID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEstoque.AlterarEstoque]", _
            Err.Description
End Sub

Public Function ListarEstoque(lngESTOQUEID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ESTOQUE.*, UNIDADE.UNIDADE AS DESCR_UNIDADE " & _
    "FROM ESTOQUE LEFT JOIN UNIDADE ON UNIDADE.PKID = ESTOQUE.UNIDADEID " & _
    "WHERE ESTOQUE.PKID = " & _
    Formata_Dados(lngESTOQUEID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarEstoque = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEstoque.ListarEstoque]", _
            Err.Description
End Function

Public Sub ExcluirESTOQUE(lngESTOQUEID As Long)
  
  Dim objRs As ADODB.Recordset
  Dim objRsInterno As ADODB.Recordset
  Dim strSql As String
  Dim clsGer As datSisLoc.clsGeral
  
  Set clsGer = New datSisLoc.clsGeral
  On Error GoTo trata
  '
'''  'EXCLUIR TABELA TAB_ENTRADAMATERIAL
'''  strSql = "DELETE FROM TAB_ENTRADAMATERIAL WHERE TAB_ENTRADAMATERIAL.ESTOQUEID = " & _
'''    Formata_Dados(lngESTOQUEID, tpDados_Longo, tpNulo_NaoAceita)
'''  RunSP strSql
'''  'EXCLUIR TABELA TAB_REQUISICAOMATERIAL
'''  strSql = "DELETE FROM TAB_REQUISICAOMATERIAL WHERE TAB_REQUISICAOMATERIAL.ESTOQUEID = " & _
'''    Formata_Dados(lngESTOQUEID, tpDados_Longo, tpNulo_NaoAceita)
'''  RunSP strSql
  'EXCLUIR TABELA ESTOQUE
  strSql = "DELETE FROM ESTOQUE WHERE PKID = " & _
    Formata_Dados(lngESTOQUEID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Set clsGer = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEstoque.ExcluirESTOQUE]", _
            Err.Description
End Sub

