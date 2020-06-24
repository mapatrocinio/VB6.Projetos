VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParceiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub ExlcuirTab_ControlParc(lngCONTROLEACESSOID As Long, _
                                  lngPARCEIROID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  '
  strSql = "DELETE FROM TAB_CONTROL_PARC WHERE PARCEIROID = " & _
    Formata_Dados(lngPARCEIROID, tpDados_Longo) & _
    " AND CONTROLEACESSOID = " & _
    Formata_Dados(lngCONTROLEACESSOID, tpDados_Longo) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsParceiro.ExlcuirTab_ControlParc]", _
            Err.Description
End Sub

Public Sub InserirTab_ControlParc(lngCONTROLEACESSOID As Long, _
                                  lngPARCEIROID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO TAB_CONTROL_PARC (PARCEIROID, CONTROLEACESSOID) VALUES (" & _
    Formata_Dados(lngPARCEIROID, tpDados_Longo) & ", " & _
    Formata_Dados(lngCONTROLEACESSOID, tpDados_Longo) & ");"
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsParceiro.InserirTab_ControlParc]", _
            Err.Description
End Sub

'''Public Function ListarTAB_GRUPOESTESTINTERPelaDescCod(strDescGrupoEstoque As String, _
'''                                                      strCodigo As String) As ADODB.Recordset
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "SELECT TAB_GRUPOESTESTINTER.* " & _
'''    "FROM (TAB_GRUPOESTESTINTER " & _
'''    "INNER JOIN ESTOQUEINTERMEDIARIO ON ESTOQUEINTERMEDIARIO.PKID = TAB_GRUPOESTESTINTER.ESTOQUEINTERMEDIARIOID) " & _
'''    "INNER JOIN GRUPOESTOQUE ON GRUPOESTOQUE.PKID = TAB_GRUPOESTESTINTER.GRUPOESTOQUEID " & _
'''    "WHERE ESTOQUEINTERMEDIARIO.CODIGO = " & Formata_Dados(strCodigo, tpDados_Texto, tpNulo_NaoAceita) & _
'''    " AND GRUPOESTOQUE.DESCRICAO = " & Formata_Dados(strDescGrupoEstoque, tpDados_Texto, tpNulo_NaoAceita)
'''  Set ListarTAB_GRUPOESTESTINTERPelaDescCod = RunSPReturnRS(strSql)
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsEstoqueIntermediario.ListarTAB_GRUPOESTESTINTERPelaDescCod]", _
'''            Err.Description
'''End Function
'''
'''Public Sub AlterarEstoque(lngESTOQUEID As Long, _
'''                          lngQtd As Long)
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''  strSql = "UPDATE ESTOQUE SET QTDESTOQUE = " & _
'''    Formata_Dados(lngQtd, tpDados_Longo, tpNulo_NaoAceita) & _
'''    " WHERE PKID = " & _
'''    Formata_Dados(lngESTOQUEID, tpDados_Longo, tpNulo_NaoAceita) & ";"
'''  '
'''  RunSP strSql
'''  '
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsEstoqueIntermediario.AlterarEstoque]", _
'''            Err.Description
'''End Sub
'''
'''Public Sub AlterarTAB_GRUPOEST(lngTAB_GRUPOEST As Long, _
'''                               lngQtdAtual As Long, _
'''                               strTipo As String)
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''  If strTipo = "C" Or strTipo = "Z" Or strTipo = "F" Then
'''    strSql = "UPDATE TAB_GRUPOESTESTINTER SET QTDESTOQUE = " & _
'''      Formata_Dados(lngQtdAtual, tpDados_Longo, tpNulo_NaoAceita) & _
'''      " WHERE PKID = " & _
'''      Formata_Dados(lngTAB_GRUPOEST, tpDados_Longo, tpNulo_NaoAceita) & ";"
'''  Else
'''    strSql = "UPDATE TAB_GRUPOESTCONSINT SET QTDESTOQUE = " & _
'''      Formata_Dados(lngQtdAtual, tpDados_Longo, tpNulo_NaoAceita) & _
'''      " WHERE PKID = " & _
'''      Formata_Dados(lngTAB_GRUPOEST, tpDados_Longo, tpNulo_NaoAceita) & ";"
'''  End If
'''  RunSP strSql
'''  '
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsEstoqueIntermediario.AlterarTAB_GRUPOEST]", _
'''            Err.Description
'''End Sub
'''
'''Public Sub ExcluirESTOQUEINTERMEDIARIO(strTipo As String, _
'''                                       lngTAB_GRUPOEST As Long, _
'''                                       lngESTOQUEINTERMEDIARIOID As Long, _
'''                                       blnTAB_GRUPO As Boolean)
'''  '------------------
'''  'strTipo Assume:
'''  'C/Z/F  - Cardápio
'''  'I      - Consumo Interno
'''  '------------------
'''  Dim strSql As String
'''  Dim clsGer As datSisMotel.clsGeral
'''
'''  Set clsGer = New datSisMotel.clsGeral
'''  On Error GoTo trata
'''  '
''''''  If strTIPO = "C" Or strTIPO = "Z" Or strTIPO = "F" Then 'CARDÁPIO
'''    'EXCLUIR RELAÇÃO COM ESTOQUEINTERMEDIARIO - TAB_GRUPOESTESTINTER
'''    strSql = "DELETE FROM TAB_GRUPOESTESTINTER WHERE PKID = " & _
'''      Formata_Dados(lngTAB_GRUPOEST, tpDados_Longo, tpNulo_NaoAceita)
'''    RunSP strSql
''''''    If blnTAB_GRUPO = False Then 'Caso não tenha outros TABGRUPOESTESTINTER, Lierado para exclusão
''''''      'EXCLUIR RELAÇÃO COM CARDAPIO - TAB_CARDESTINTER
''''''      strSql = "DELETE FROM TAB_CARDESTINTER WHERE ESTOQUEINTERMEDIARIOID = " & _
''''''        Formata_Dados(lngESTOQUEINTERMEDIARIOID, tpDados_Longo, tpNulo_NaoAceita)
''''''      RunSP strSql
''''''      'EXCLUIR ESTOQUEINTERMEDIARIO
''''''      strSql = "DELETE FROM ESTOQUEINTERMEDIARIO WHERE PKID = " & _
''''''        Formata_Dados(lngESTOQUEINTERMEDIARIOID, tpDados_Longo, tpNulo_NaoAceita)
''''''      RunSP strSql
''''''    End If
''''''  ElseIf strTIPO = "I" Then 'CONSUMO INTERNO
''''''    'EXCLUIR RELAÇÃO COM BAIXAESTOQUE - TAB_BAIXAESTOQUECONSINT
''''''    strSql = "DELETE FROM TAB_BAIXAESTOQUECONSINTERNO WHERE TAB_GRUPOESTCONSINTID = " & _
''''''      Formata_Dados(lngTAB_GRUPOEST, tpDados_Longo, tpNulo_NaoAceita)
''''''    RunSP strSql
''''''    'EXCLUIR RELAÇÃO COM CONSUMOINTERNO - TAB_GRUPOESTCONSINT
''''''    strSql = "DELETE FROM TAB_GRUPOESTCONSINT WHERE PKID = " & _
''''''      Formata_Dados(lngTAB_GRUPOEST, tpDados_Longo, tpNulo_NaoAceita)
''''''    RunSP strSql
''''''    If blnTAB_GRUPO = False Then 'Caso não tenha outros TAB_GRUPOESTCONSINT, Liberado para exclusão
''''''      'EXCLUIR CONSUMOINTERNO
''''''      strSql = "DELETE FROM CONSUMOINTERNO WHERE PKID = " & _
''''''        Formata_Dados(lngESTOQUEINTERMEDIARIOID, tpDados_Longo, tpNulo_NaoAceita)
''''''      RunSP strSql
''''''    End If
''''''  End If
'''  '
'''  Set clsGer = Nothing
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsEstoqueIntermediario.ExcluirESTOQUEINTERMEDIARIO]", _
'''            Err.Description
'''End Sub
'''
'''
'''Public Sub AlterarEstoqueIntermediario(strCodigo As String, _
'''                                       strDescricao As String, _
'''                                       strUnidade As String, _
'''                                       lngESTOQUEID As Long)
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "UPDATE ESTOQUEINTERMEDIARIO SET CODIGO = " & Formata_Dados(strCodigo, tpDados_Texto, tpNulo_NaoAceita) & _
'''    ", DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & _
'''    ", UNIDADE = " & Formata_Dados(strUnidade, tpDados_Texto, tpNulo_NaoAceita) & _
'''    ", ESTOQUEID = " & Formata_Dados(lngESTOQUEID, tpDados_Longo, tpNulo_NaoAceita) & _
'''    " WHERE ESTOQUEID = " & lngESTOQUEID
'''  '
'''  RunSP strSql
'''  '
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsEstoqueIntermediario.AlterarEstoqueIntermediario]", _
'''            Err.Description
'''End Sub
'''
'''Public Sub InserirEstoqueIntermediario(strCodigo As String, _
'''                                       strDescricao As String, _
'''                                       strUnidade As String, _
'''                                       lngESTOQUEID As Long)
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "INSERT INTO ESTOQUEINTERMEDIARIO (CODIGO, DESCRICAO, UNIDADE, ESTOQUEID) VALUES (" & _
'''    Formata_Dados(strCodigo, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
'''    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
'''    Formata_Dados(strUnidade, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
'''    Formata_Dados(lngESTOQUEID, tpDados_Longo, tpNulo_NaoAceita) & ");"
'''  '
'''  RunSP strSql
'''  '
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsEstoqueIntermediario.InserirEstoqueIntermediario]", _
'''            Err.Description
'''End Sub
'''
'''Public Function ListarEstoqueIntermediarioPeloCodigo(strCodigo As String) As ADODB.Recordset
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "SELECT * FROM ESTOQUEINTERMEDIARIO WHERE CODIGO = " & _
'''    Formata_Dados(strCodigo, tpDados_Texto, tpNulo_NaoAceita)
'''  Set ListarEstoqueIntermediarioPeloCodigo = RunSPReturnRS(strSql)
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsEstoqueIntermediario.ListarEstoqueIntermediarioPeloCodigo]", _
'''            Err.Description
'''End Function
'''
'''Public Function CapturarGrupoEstoqueID(ByVal lngLocacaoId As Long, _
'''                                       ByVal strTipo As String) As Long
'''  On Error GoTo trata
'''  '
'''  Dim strSql    As String
'''  Dim objRs     As ADODB.Recordset
'''  Dim lngRet    As Long
'''
'''  If lngLocacaoId = 0 Then ' VENDA
'''    strSql = "SELECT GRUPOESTOQUE.PKID FROM GRUPOESTOQUE WHERE TIPO = " & _
'''      Formata_Dados(strTipo, tpDados_Texto, tpNulo_NaoAceita) & _
'''      " AND PADRAOVENDAS = " & Formata_Dados(True, tpDados_Boolean)
'''  Else ' PEDIDO
'''    strSql = "SELECT GRUPOESTOQUE.PKID FROM ((LOCACAO INNER JOIN APARTAMENTO ON LOCACAO.APARTAMENTOID = APARTAMENTO.PKID) " & _
'''      " INNER JOIN TAB_GRUPOESTAPTO ON APARTAMENTO.PKID = TAB_GRUPOESTAPTO.APARTAMENTOID) " & _
'''      " INNER JOIN GRUPOESTOQUE ON GRUPOESTOQUE.PKID = TAB_GRUPOESTAPTO.GRUPOESTOQUEID " & _
'''      " WHERE LOCACAO.PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_NaoAceita) & _
'''      " AND GRUPOESTOQUE.TIPO = " & Formata_Dados(strTipo, tpDados_Texto, tpNulo_NaoAceita)
'''
'''  End If
'''  Set objRs = RunSPReturnRS(strSql)
'''  If objRs.EOF Then
'''    lngRet = 0
'''  Else
'''    lngRet = objRs.Fields("PKID").Value
'''  End If
'''  '
'''  objRs.Close
'''  Set objRs = Nothing
'''  CapturarGrupoEstoqueID = lngRet
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsEstoqueIntermediario.CapturarGrupoEstoqueID]", _
'''            Err.Description
'''End Function
'''
'''
'''Public Sub ExlcuirTab_GrupoEstApto(lngAPARTAMENTOID As Long, _
'''                                   lngGRUPOESTOQUEID As Long)
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''  '
'''  strSql = "DELETE FROM TAB_GRUPOESTAPTO WHERE GRUPOESTOQUEID = " & _
'''    Formata_Dados(lngGRUPOESTOQUEID, tpDados_Longo, tpNulo_NaoAceita) & _
'''    " AND APARTAMENTOID = " & _
'''    Formata_Dados(lngAPARTAMENTOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
'''
'''  RunSP strSql
'''  '
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsEstoqueIntermediario.ExlcuirTab_GrupoEstApto]", _
'''            Err.Description
'''End Sub
'''
'''
'''Public Sub InserirTab_GrupoEstApto(lngAPARTAMENTOID As Long, _
'''                                   lngGRUPOESTOQUEID As Long)
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "INSERT INTO TAB_GRUPOESTAPTO (APARTAMENTOID, GRUPOESTOQUEID) VALUES (" & _
'''    Formata_Dados(lngAPARTAMENTOID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
'''    Formata_Dados(lngGRUPOESTOQUEID, tpDados_Longo, tpNulo_NaoAceita) & ");"
'''  RunSP strSql
'''  '
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsEstoqueIntermediario.InserirTab_GrupoEstApto]", _
'''            Err.Description
'''End Sub
'''
Public Sub ExlcuirParceiro(lngPARCEIROID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  '
  strSql = "DELETE FROM TAB_CONTROL_PARC WHERE PARCEIROID = " & _
    Formata_Dados(lngPARCEIROID, tpDados_Longo) & ";"

  RunSP strSql
  '
  strSql = "DELETE FROM PARCEIRO WHERE PKID = " & _
    Formata_Dados(lngPARCEIROID, tpDados_Longo) & ";"

  RunSP strSql
  '

  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsParceiro.ExlcuirParceiro]", _
            Err.Description
End Sub


Public Function VerificaParceiro(lngPARCEIROID As Long) As Boolean
  On Error GoTo trata
  '
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim blnRet    As Boolean

  strSql = "SELECT COUNT(*) AS TOTALLANC FROM TAB_CONTROL_PARC WHERE PARCEIROID = " & _
    Formata_Dados(lngPARCEIROID, tpDados_Longo)
  Set objRs = RunSPReturnRS(strSql)
  If objRs.EOF Then
    blnRet = True
  Else
    If objRs.Fields("TOTALLANC").Value = 0 Then
      blnRet = True
    Else
      blnRet = False
    End If
  End If
  '
  objRs.Close
  Set objRs = Nothing

  VerificaParceiro = blnRet
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsParceiro.VerificaParceiro]", _
            Err.Description
End Function

Public Sub AlterarParceiro(strDescricao As String, _
                           lngPARCEIROID As Long)
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE PARCEIRO SET DESCRICAO = " & _
    Formata_Dados(strDescricao, tpDados_Texto) & _
    " WHERE PKID = " & _
    Formata_Dados(lngPARCEIROID, tpDados_Longo) & ";"
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsParceiro.AlterarParceiro]", _
            Err.Description
End Sub

Public Sub InserirParceiro(strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "INSERT INTO PARCEIRO (DESCRICAO) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ");"
    
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsParceiro.InserirParceiro]", _
            Err.Description
End Sub
'''
'''Public Function ValidacamposGrupoEstoque(strDescricao As String, _
'''                                         strTipo As String, _
'''                                         blnPadraoVendas As Boolean, _
'''                                         lngGRUPOESTOQUEID As Long, _
'''                                         strMsgErro As String) As Boolean
'''  On Error GoTo trata
'''  '
'''  Dim strSql    As String
'''  Dim objRs     As ADODB.Recordset
'''  Dim blnRet    As Boolean
'''
'''  strSql = "SELECT * FROM GRUPOESTOQUE WHERE DESCRICAO = " & _
'''    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & _
'''    " AND PKID <> " & Formata_Dados(lngGRUPOESTOQUEID, tpDados_Longo, tpNulo_NaoAceita)
'''
'''  Set objRs = RunSPReturnRS(strSql)
'''  If Not objRs.EOF Then
'''    blnRet = False
'''    strMsgErro = "Estoque intermediário já cadastrado."
'''  Else
'''    blnRet = True
'''    strMsgErro = ""
'''  End If
'''  '
'''  objRs.Close
'''  Set objRs = Nothing
'''  '
'''  If blnRet Then 'Se não houve erro
'''    If blnPadraoVendas Then 'Caso esteja marcada a opção Padrão vendas, verifica se já não há um padrão cadastrado para este tipo de estoque
'''      strSql = "SELECT COUNT(*) AS TOTALGRUPO FROM GRUPOESTOQUE WHERE TIPO = " & _
'''        Formata_Dados(strTipo, tpDados_Texto, tpNulo_NaoAceita) & _
'''        " AND PADRAOVENDAS = " & Formata_Dados(True, tpDados_Boolean) & _
'''        " AND PKID <> " & Formata_Dados(lngGRUPOESTOQUEID, tpDados_Longo, tpNulo_NaoAceita)
'''
'''      Set objRs = RunSPReturnRS(strSql)
'''      If objRs.EOF Then
'''        blnRet = True
'''        strMsgErro = ""
'''      Else
'''        If objRs.Fields("TOTALGRUPO").Value = 0 Then
'''          blnRet = True
'''          strMsgErro = ""
'''        Else
'''          blnRet = False
'''          strMsgErro = "Já existe estoque intermediário cadastrado para este tipo de estoque intermediário com padrão vendas marcado para ele."
'''        End If
'''      End If
'''      '
'''      objRs.Close
'''      Set objRs = Nothing
'''      '
'''    End If
'''  End If
'''  ValidacamposGrupoEstoque = blnRet
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsEstoqueIntermediario.ValidacamposGrupoEstoque]", _
'''            Err.Description
'''End Function
'''
'''
'''Public Sub AtualizaQtdSUBITEM(curQuantidade As Currency, _
'''                              lngTAB_GRUPOESTESTINTERID As Long)
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "UPDATE TAB_GRUPOESTESTINTER SET QTDESTOQUE = " & Formata_Dados(curQuantidade, tpDados_Moeda, tpNulo_NaoAceita) & " " & _
'''          "WHERE PKID = " & Formata_Dados(lngTAB_GRUPOESTESTINTERID, tpDados_Longo, tpNulo_NaoAceita)
'''
'''  RunSP strSql
'''  '
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsEstoqueFrigobar.AtualizaQtdSUBITEM]", _
'''            Err.Description
'''End Sub
'''
Public Function ListarParceiroPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT PARCEIRO.* FROM PARCEIRO WHERE PARCEIRO.DESCRICAO = " & _
    Formata_Dados(strDescricao, tpDados_Texto)
  Set ListarParceiroPelaDesc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsParceiro.ListarParceiroPelaDesc]", _
            Err.Description
End Function

Public Function ListarParceiro(lngPARCEIROID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT PARCEIRO.* FROM PARCEIRO WHERE PKID = " & _
    Formata_Dados(lngPARCEIROID, tpDados_Longo)
  Set ListarParceiro = RunSPReturnRS(strSql)
  If ListarParceiro.EOF Then
    ListarParceiro.Close
    Set ListarParceiro = Nothing
    Err.Raise lngCteErroData, "[" & strClassName & ".clsParceiro.ListarParceiro]", "Não há parceiro cadastrado."
  End If
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsParceiro.ListarParceiro]", _
            Err.Description
End Function



