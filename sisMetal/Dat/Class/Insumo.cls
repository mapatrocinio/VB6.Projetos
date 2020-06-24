VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Public Function CapturaProduto(ByVal strCodigo As String) As ADODB.Recordset
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "SELECT PRODUTO.*, INSUMO.CODIGO FROM PRODUTO " & _
    " INNER JOIN INSUMO ON INSUMO.PKID = PRODUTO.INSUMOID " & _
    " WHERE " & _
    " INSUMO.CODIGO LIKE " & Formata_Dados("%" & strCodigo & "%", tpDados_Texto) & _
    " OR PRODUTO.NOME LIKE " & Formata_Dados("%" & strCodigo & "%", tpDados_Texto)
  Set CapturaProduto = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLinhaPerfil.CapturaProduto]", _
            Err.Description
End Function

Public Function ExcluirInsumo(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  'PERFIL
  strSql = "DELETE FROM PERFIL " & _
    "WHERE INSUMOID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  'ACESSORIO
  strSql = "DELETE FROM ACESSORIO " & _
    "WHERE INSUMOID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  'INSUMO
  strSql = "DELETE FROM INSUMO " & _
    "WHERE PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsInsumo.ExcluirInsumo]", _
            Err.Description
End Function


Public Function SelecionarInsumoPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT INSUMO.*, " & _
    " GRUPO.NOME AS NOME_GRUPO, EMBALAGEM.NOME AS NOME_EMBALAGEM, ACESSORIO.NOME, ACESSORIO.VALOR, ACESSORIO.QTD_MINIMA, ACESSORIO.QTD_ESTOQUE, " & _
    " PERFIL.PESO_MINIMO, PERFIL.PESO_ESTOQUE, PERFIL.LINHAID, PERFIL.CORID, COR.NOME AS NOME_COR, TIPO_LINHA.NOME AS NOME_LINHA, LINHA.CODIGO AS CODIGO_LINHA " & _
    "FROM INSUMO " & _
    " LEFT JOIN ACESSORIO ON INSUMO.PKID = ACESSORIO.INSUMOID " & _
    " LEFT JOIN PERFIL ON INSUMO.PKID = PERFIL.INSUMOID " & _
    " LEFT JOIN COR ON COR.PKID = PERFIL.CORID " & _
    " LEFT JOIN GRUPO ON GRUPO.PKID = ACESSORIO.GRUPOID " & _
    " LEFT JOIN EMBALAGEM ON EMBALAGEM.PKID = ACESSORIO.EMBALAGEMID " & _
    " LEFT JOIN LINHA ON LINHA.PKID = PERFIL.LINHAID " & _
    " LEFT JOIN TIPO_LINHA ON TIPO_LINHA.PKID = LINHA.TIPO_LINHAID " & _
    "WHERE INSUMO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarInsumoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsInsumo.SelecionarInsumoPeloPkid]", _
            Err.Description
End Function

Public Function SelecionarProdutoPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT PRODUTO.*, " & _
    " GRUPO_PRODUTO.NOME AS NOME_GRUPO, EMBALAGEM.NOME AS NOME_EMBALAGEM, LOJA.NOME AS NOME_FORNECEDOR, " & _
    " FAMILIAPRODUTOS.DESCRICAO AS NOME_FAMILIA, IPI.IPI AS NOME_IPI, ICMS.ICMS AS NOME_ICMS " & _
    "FROM INSUMO " & _
    " INNER JOIN PRODUTO ON INSUMO.PKID = PRODUTO.INSUMOID " & _
    " LEFT JOIN EMBALAGEM ON EMBALAGEM.PKID = PRODUTO.EMBALAGEMID " & _
    " LEFT JOIN GRUPO_PRODUTO ON GRUPO_PRODUTO.PKID = PRODUTO.GRUPO_PRODUTOID " & _
    " LEFT JOIN LOJA ON LOJA.PKID = PRODUTO.FORNECEDORID " & _
    " LEFT JOIN FAMILIAPRODUTOS ON FAMILIAPRODUTOS.PKID = PRODUTO.FAMILIAPRODUTOSID " & _
    " LEFT JOIN IPI ON IPI.PKID = PRODUTO.IPIID " & _
    " LEFT JOIN ICMS ON ICMS.PKID = PRODUTO.ICMSID " & _
    "WHERE INSUMO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarProdutoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsInsumo.SelecionarProdutoPeloPkid]", _
            Err.Description
End Function

Public Function SelecionarInsumoPeloCodigo(strCodigo As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT INSUMO.*, ACESSORIO.NOME, ACESSORIO.VALOR, PERFIL.LINHAID, PERFIL.CORID " & _
    "FROM INSUMO " & _
    " LEFT JOIN ACESSORIO ON INSUMO.PKID = ACESSORIO.INSUMOID " & _
    " LEFT JOIN PERFIL ON INSUMO.PKID = PERFIL.INSUMOID " & _
    " WHERE INSUMO.CODIGO = " & _
    Formata_Dados(strCodigo, tpDados_Texto)
  Set SelecionarInsumoPeloCodigo = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsInsumo.SelecionarInsumoPeloCodigo]", _
            Err.Description
End Function
Public Function SelecionarProdutoPeloCodigo(strCodigo As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT PRODUTO.*, INSUMO.CODIGO " & _
    "FROM PRODUTO " & _
    " INNER JOIN INSUMO ON INSUMO.PKID = PRODUTO.INSUMOID " & _
    " WHERE INSUMO.CODIGO LIKE '%" & _
    strCodigo & "%'"
  Set SelecionarProdutoPeloCodigo = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsInsumo.SelecionarProdutoPeloCodigo]", _
            Err.Description
End Function


Public Sub AlterarInsumo(lngPKID As Long, _
                         strCodigo As String, _
                         strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE INSUMO SET " & _
    "CODIGO = " & Formata_Dados(strCodigo, tpDados_Texto) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsInsumo.AlterarInsumo]", _
            Err.Description
End Sub

Public Sub InserirInsumo(ByRef lngPKID As Long, _
                         strCodigo As String, _
                         strStatus As String)

  On Error GoTo trata
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO INSUMO (CODIGO, STATUS) VALUES (" & _
    Formata_Dados(strCodigo, tpDados_Texto) & ", " & _
    Formata_Dados(strStatus, tpDados_Texto) & ");"
  '
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"

  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngPKID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsInsumo.InserirInsumo]", _
            Err.Description
End Sub

Public Sub AlterarAcessorio(lngPKID As Long, _
                            lngGRUPOID As Long, _
                            lngEMBALAGEMID As Long, _
                            strNome As String, _
                            strValor As String, _
                            strQtdMinima As String, _
                            strQtdEstoque As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "UPDATE ACESSORIO " & _
    "SET VALOR = " & Formata_Dados(strValor, tpDados_Moeda) & ", " & _
    " NOME = " & Formata_Dados(strNome, tpDados_Texto) & ", " & _
    " GRUPOID = " & Formata_Dados(lngGRUPOID, tpDados_Longo) & ", " & _
    " EMBALAGEMID = " & Formata_Dados(lngEMBALAGEMID, tpDados_Longo) & ", " & _
    " QTD_MINIMA = " & Formata_Dados(strQtdMinima, tpDados_Longo) & ", " & _
    " QTD_ESTOQUE = " & Formata_Dados(strQtdEstoque, tpDados_Longo) & " " & _
    " WHERE INSUMOID = " & Formata_Dados(lngPKID, tpDados_Longo)
  '
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsInsumo.AlterarAcessorio]", _
            Err.Description
End Sub

Public Sub AlterarProduto(lngPKID As Long, _
                          lngGRUPOPRODUTOID As Long, _
                          lngEMBALAGEMPRODUTOID As Long, _
                          lngFORNECEDORID As Long, _
                          lngFAMILIAID As Long, _
                          lngIPIID As Long, _
                          lngICMSID As Long, _
                          strNomeProduto As String, _
                          strNomeAbrevProduto As String, _
                          strValorProduto As String, _
                          strPeso As String, _
                          strAltura As String, _
                          strLargura As String, _
                          strModRef As String, _
                          strTabela As String, _
                          strFinancVenda As String, _
                          strEstoqueMinimo As String, _
                          strMargemEstoque As String, _
                          strSaldoEstoque As String, _
                          strCustoProduto As String, _
                          strMargemAjuste As String, _
                          strPrecoVenda As String, _
                          strTAM As String, _
                          strPAD As String, _
                          strSOB As String, strComissaoVendedor As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "UPDATE PRODUTO " & _
    " SET NOME = " & Formata_Dados(strNomeProduto, tpDados_Texto) & _
    ", NOMEABREVIADO = " & Formata_Dados(strNomeAbrevProduto, tpDados_Texto) & _
    ", EMBALAGEMID = " & Formata_Dados(IIf(lngEMBALAGEMPRODUTOID = 0, "", lngEMBALAGEMPRODUTOID), tpDados_Longo) & _
    ", GRUPO_PRODUTOID = " & Formata_Dados(IIf(lngGRUPOPRODUTOID = 0, "", lngGRUPOPRODUTOID), tpDados_Longo) & _
    ", FORNECEDORID = " & Formata_Dados(IIf(lngFORNECEDORID = 0, "", lngFORNECEDORID), tpDados_Longo) & _
    ", PRECO = " & Formata_Dados(strValorProduto, tpDados_Moeda) & _
    ", PESO = " & Formata_Dados(strPeso, tpDados_Moeda) & _
    ", ALTESQUADRIA = " & Formata_Dados(strAltura, tpDados_Moeda) & _
    ", LARGESQUADRIA = " & Formata_Dados(strLargura, tpDados_Moeda) & _
    ", MODELOREFERENCIA = " & Formata_Dados(strModRef, tpDados_Texto) & _
    ", TABELA = " & Formata_Dados(strTabela, tpDados_Texto) & _
    ", FAMILIAPRODUTOSID = " & Formata_Dados(IIf(lngFAMILIAID = 0, "", lngFAMILIAID), tpDados_Longo) & _
    ", IPIID = " & Formata_Dados(IIf(lngIPIID = 0, "", lngIPIID), tpDados_Longo) & _
    ", ICMSID = " & Formata_Dados(IIf(lngICMSID = 0, "", lngICMSID), tpDados_Longo)
  strSql = strSql & _
    ", FINANCVENDA = " & Formata_Dados(strFinancVenda, tpDados_Moeda) & _
    ", ESTMINIMO = " & Formata_Dados(strEstoqueMinimo, tpDados_Moeda) & _
    ", MARGEMESTOQUE = " & Formata_Dados(strMargemEstoque, tpDados_Moeda) & _
    ", SALDOESTOQUE = " & Formata_Dados(strSaldoEstoque, tpDados_Moeda) & _
    ", CUSTOPRODUTO = " & Formata_Dados(strCustoProduto, tpDados_Moeda) & _
    ", MARGEMAJUSTE = " & Formata_Dados(strMargemAjuste, tpDados_Moeda) & _
    ", PRECOVENDA = " & Formata_Dados(strPrecoVenda, tpDados_Moeda) & _
    ", TAM = " & Formata_Dados(strTAM, tpDados_Moeda) & _
    ", PAD = " & Formata_Dados(strPAD, tpDados_Longo) & _
    ", SOB = " & Formata_Dados(strSOB, tpDados_Longo) & _
    ", COMISSAO_VENDEDOR = " & Formata_Dados(strComissaoVendedor, tpDados_Boolean) & _
    " WHERE INSUMOID = " & Formata_Dados(lngPKID, tpDados_Longo)
  '
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsInsumo.AlterarProduto]", _
            Err.Description
End Sub

Public Sub AlterarPerfil(lngPKID As Long, _
                         lngLINHAID As Long, _
                         lngCORID As Long, _
                         strPesoMinimo As String, _
                         strPesoEstoque As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "UPDATE PERFIL " & _
    "SET LINHAID = " & Formata_Dados(lngLINHAID, tpDados_Longo) & ", " & _
    "CORID = " & Formata_Dados(lngCORID, tpDados_Longo) & ", " & _
    "PESO_MINIMO = " & Formata_Dados(strPesoMinimo, tpDados_Moeda) & ", " & _
    "PESO_ESTOQUE = " & Formata_Dados(strPesoEstoque, tpDados_Moeda) & " " & _
    " WHERE INSUMOID = " & Formata_Dados(lngPKID, tpDados_Longo)
  '
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsInsumo.AlterarPerfil]", _
            Err.Description
End Sub


Public Sub InserirAcessorio(lngPKID As Long, _
                            lngGRUPOID As Long, _
                            lngEMBALAGEMID As Long, _
                            strNome As String, _
                            strValor As String, _
                            strQtdMinima As String, _
                            strQtdEstoque As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO ACESSORIO (INSUMOID, NOME, GRUPOID, EMBALAGEMID, VALOR, QTD_MINIMA, QTD_ESTOQUE) VALUES (" & _
    Formata_Dados(lngPKID, tpDados_Longo) & ", " & _
    Formata_Dados(strNome, tpDados_Texto) & ", " & _
    Formata_Dados(lngGRUPOID, tpDados_Longo) & ", " & _
    Formata_Dados(lngEMBALAGEMID, tpDados_Longo) & ", " & _
    Formata_Dados(strValor, tpDados_Moeda) & ", " & _
    Formata_Dados(strQtdMinima, tpDados_Longo) & ", " & _
    Formata_Dados(strQtdEstoque, tpDados_Longo) & " " & ");"
  '
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsInsumo.InserirAcessorio]", _
            Err.Description
End Sub

Public Sub InserirProduto(lngPKID As Long, _
                          lngGRUPOPRODUTOID As Long, _
                          lngEMBALAGEMPRODUTOID As Long, _
                          lngFORNECEDORID As Long, _
                          lngFAMILIAID As Long, _
                          lngIPIID As Long, _
                          lngICMSID As Long, _
                          strNomeProduto As String, _
                          strNomeAbrevProduto As String, _
                          strValorProduto As String, _
                          strPeso As String, _
                          strAltura As String, _
                          strLargura As String, _
                          strModRef As String, _
                          strTabela As String, _
                          strFinancVenda As String, _
                          strEstoqueMinimo As String, _
                          strMargemEstoque As String, _
                          strSaldoEstoque As String, _
                          strCustoProduto As String, _
                          strMargemAjuste As String, _
                          strPrecoVenda As String, _
                          strTAM As String, _
                          strPAD As String, _
                          strSOB As String, strComissaoVendedor As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO PRODUTO (INSUMOID, NOME, NOMEABREVIADO, EMBALAGEMID, GRUPO_PRODUTOID, FORNECEDORID, PRECO, PESO, ALTESQUADRIA, " & _
    "LARGESQUADRIA, MODELOREFERENCIA, TABELA, FAMILIAPRODUTOSID, IPIID, ICMSID, FINANCVENDA, " & _
    "ESTMINIMO , MARGEMESTOQUE, SALDOESTOQUE, CUSTOPRODUTO, MARGEMAJUSTE, PRECOVENDA, TAM, PAD, SOB, COMISSAO_VENDEDOR) VALUES ("
  strSql = strSql & _
    " " & Formata_Dados(lngPKID, tpDados_Longo) & _
    ", " & Formata_Dados(strNomeProduto, tpDados_Texto) & _
    ", " & Formata_Dados(strNomeAbrevProduto, tpDados_Texto) & _
    ", " & Formata_Dados(IIf(lngEMBALAGEMPRODUTOID = 0, "", lngEMBALAGEMPRODUTOID), tpDados_Longo) & _
    ", " & Formata_Dados(IIf(lngGRUPOPRODUTOID = 0, "", lngGRUPOPRODUTOID), tpDados_Longo) & _
    ", " & Formata_Dados(IIf(lngFORNECEDORID = 0, "", lngFORNECEDORID), tpDados_Longo) & _
    ", " & Formata_Dados(strValorProduto, tpDados_Moeda) & _
    ", " & Formata_Dados(strPeso, tpDados_Moeda) & _
    ", " & Formata_Dados(strAltura, tpDados_Moeda) & _
    ", " & Formata_Dados(strLargura, tpDados_Moeda) & _
    ", " & Formata_Dados(strModRef, tpDados_Texto) & _
    ", " & Formata_Dados(strTabela, tpDados_Texto) & _
    ", " & Formata_Dados(IIf(lngFAMILIAID = 0, "", lngFAMILIAID), tpDados_Longo) & _
    ", " & Formata_Dados(IIf(lngIPIID = 0, "", lngIPIID), tpDados_Longo) & _
    ", " & Formata_Dados(IIf(lngICMSID = 0, "", lngICMSID), tpDados_Longo)
  strSql = strSql & _
    ", " & Formata_Dados(strFinancVenda, tpDados_Moeda) & _
    ", " & Formata_Dados(strEstoqueMinimo, tpDados_Moeda) & _
    ", " & Formata_Dados(strMargemEstoque, tpDados_Moeda) & _
    ", " & Formata_Dados(strSaldoEstoque, tpDados_Moeda) & _
    ", " & Formata_Dados(strCustoProduto, tpDados_Moeda) & _
    ", " & Formata_Dados(strMargemAjuste, tpDados_Moeda) & _
    ", " & Formata_Dados(strPrecoVenda, tpDados_Moeda) & _
    ", " & Formata_Dados(strTAM, tpDados_Moeda) & _
    ", " & Formata_Dados(strPAD, tpDados_Longo) & _
    ", " & Formata_Dados(strSOB, tpDados_Longo) & _
    ", " & Formata_Dados(strComissaoVendedor, tpDados_Boolean) & _
    ")"
  '
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsInsumo.InserirProduto]", _
            Err.Description
End Sub

Public Sub InserirPerfil(lngPKID As Long, _
                         lngLINHAID As Long, _
                         lngCORID As Long, _
                         strPesoMinimo As String, _
                         strPesoEstoque As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO PERFIL (INSUMOID, LINHAID, CORID, PESO_MINIMO, PESO_ESTOQUE) VALUES (" & _
    Formata_Dados(lngPKID, tpDados_Longo) & ", " & _
    Formata_Dados(lngLINHAID, tpDados_Longo) & ", " & _
    Formata_Dados(lngCORID, tpDados_Longo) & ", " & _
    Formata_Dados(strPesoMinimo, tpDados_Moeda) & ", " & _
    Formata_Dados(strPesoEstoque, tpDados_Moeda) & " " & ");"
  '

  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsInsumo.InserirPerfil]", _
            Err.Description
End Sub


