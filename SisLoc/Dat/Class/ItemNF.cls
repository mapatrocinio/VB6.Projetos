VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItemNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarItemNFContrato(lngCONTRATOID As Long, _
                                         lngOBRAID As Long, _
                                         lngESTOQUEID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  strSql = "SELECT ITEMNF.PKID,ITEMNF.NFID, ISNULL(MIN(vw_cons_baixa.QTD_REAL), 0) AS QTD_REAL " & _
    "FROM ITEMNF INNER JOIN NF ON NF.PKID = ITEMNF.NFID " & _
    " INNER JOIN ESTOQUE ON ESTOQUE.PKID = ITEMNF.ESTOQUEID " & _
    " INNER JOIN vw_cons_baixa ON ITEMNF.PKID = vw_cons_baixa.ITEMNFID " & _
    "WHERE NF.CONTRATOID = " & Formata_Dados(lngCONTRATOID, tpDados_Longo) & _
    " AND NF.OBRAID = " & Formata_Dados(lngOBRAID, tpDados_Longo) & _
    " AND ESTOQUE.PKID = " & Formata_Dados(lngESTOQUEID, tpDados_Longo) & _
    " AND NF.STATUS = " & Formata_Dados("F", tpDados_Texto) & _
    " GROUP BY ITEMNF.PKID, ITEMNF.NFID, NF.DTSAIDA " & _
    " ORDER BY NF.DTSAIDA "
  Set SelecionarItemNFContrato = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsNF.SelecionarItemNFContrato]", _
            Err.Description
End Function

Public Sub AlterarEstoquePeloRetItemNF(ByVal lngESTOQUEID As Long, _
                                       ByVal strQuantidade As String)
  '
  On Error GoTo trata
  '
  Dim strSql  As String
  '
  strSql = "UPDATE ESTOQUE SET QUANTIDADE = ISNULL(QUANTIDADE,0) + " & Formata_Dados(strQuantidade, tpDados_Longo) & " " & _
    "WHERE PKID = " & lngESTOQUEID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.AlterarEstoquePeloRetItemNF]", _
            Err.Description
End Sub

Public Sub AlterarEstoquePeloItemNF(ByVal lngESTOQUEID As Long, _
                                    ByVal strQuantidade As String)
  '
  On Error GoTo trata
  '
  Dim strSql  As String
  '
  strSql = "UPDATE ESTOQUE SET QUANTIDADE = ISNULL(QUANTIDADE,0) - " & Formata_Dados(strQuantidade, tpDados_Longo) & " " & _
    "WHERE PKID = " & lngESTOQUEID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.AlterarEstoquePeloItemNF]", _
            Err.Description
End Sub
            
Public Sub InserirITEMNF(lngNFID As Long, _
                         lngESTOQUEID As Long, _
                         strQuantidade As String, _
                         strValor As String, _
                         strValorEstoque As String, _
                         strLargura As String, _
                         strAltura As String, _
                         strValorLanc As String)
  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO ITEMNF (NFID, ESTOQUEID, QUANTIDADE, VALORUNITARIO, VALORESTOQUE, LARGURA, ALTURA, VALORLANC) VALUES (" & _
    " " & Formata_Dados(lngNFID, tpDados_Longo) & _
    ", " & Formata_Dados(lngESTOQUEID, tpDados_Longo) & _
    ", " & Formata_Dados(strQuantidade, tpDados_Longo) & _
    ", " & Formata_Dados(strValor, tpDados_Moeda) & _
    ", " & Formata_Dados(strValorEstoque, tpDados_Moeda) & _
    ", " & Formata_Dados(strLargura, tpDados_Moeda) & _
    ", " & Formata_Dados(strAltura, tpDados_Moeda) & _
    ", " & Formata_Dados(strValorLanc, tpDados_Moeda) & ");"
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsITEMNF.InserirITEMNF]", _
            Err.Description
End Sub


Public Function ExcluirITEMNF(lngITEMNFID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "DELETE FROM ITEMNF " & _
    "WHERE PKID = " & _
    Formata_Dados(lngITEMNFID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsITEMNF.ExcluirITEMNF]", _
            Err.Description
End Function
'''
'''
'''Public Function SelecionarNFTotalProc(lngNFID As Long) As ADODB.Recordset
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''  strSql = "SELECT isnull(SUM(ITEMNF.VALORUNITARIO), 0) AS TOTALAPAGAR " & _
'''    "FROM ITEMNF " & _
'''    "WHERE ITEMNF.NFID = " & _
'''    Formata_Dados(lngNFID, tpDados_Longo)
'''  Set SelecionarNFTotalProc = RunSPReturnRS(strSql)
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsNF.SelecionarNFTotalProc]", _
'''            Err.Description
'''End Function

'''Public Function CapturaNFTurnoCorrente(strTURNOID As String, _
'''                                       lngCONTRATOID As Long) As ADODB.Recordset
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''  Dim strData As String
'''  strSql = "SELECT DISTINCT NF.PKID, NF.DATA, NF.SEQUENCIAL, EMPRESA.NOME, ESPECIALIDADE.ESPECIALIDADE, PRESTADOR.NOME, SALA.NUMERO AS SALA_NUMERO,  '(' + FUNC.NIVEL + ') ' + FUNCIONARIO.NOME, NF.STATUS, NF.STATUSIMPRESSAO " & _
'''    "FROM NF INNER JOIN EMPRESA ON EMPRESA.PKID = NF.EMPRESAID " & _
'''    " INNER JOIN TURNO ON TURNO.PKID = NF.TURNOID " & _
'''    "   OR TURNO.PKID = NF.TURNOLABID " & _
'''    " INNER JOIN ATENDE ON ATENDE.PKID = NF.ATENDEID " & _
'''      " INNER JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
'''      " INNER JOIN EMPRESA AS PRESTADOR ON PRESTADOR.PKID = ATENDE.EMPRESAID " & _
'''    " INNER JOIN EMPRESA AS FUNCIONARIO ON FUNCIONARIO.PKID = NF.FUNCIONARIOID " & _
'''    " INNER JOIN FUNCIONARIO AS FUNC ON FUNCIONARIO.PKID = FUNC.EMPRESAID " & _
'''    " INNER JOIN ESPECIALIDADE ON ESPECIALIDADE.PKID = NF.ESPECIALIDADEID " & _
'''    "WHERE NF.TURNOID IN " & _
'''      strTURNOID & _
'''      " OR (TURNO.STATUS = " & Formata_Dados(1, tpDados_Longo) & _
'''      " AND TURNO.EMPRESAID = " & Formata_Dados(lngCONTRATOID, tpDados_Longo) & ")" & _
'''      " ORDER BY NF.DATA DESC, NF.SEQUENCIAL DESC "
'''  Set CapturaNFTurnoCorrente = RunSPReturnRS(strSql)
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsNF.CapturaNFTurnoCorrente]", _
'''            Err.Description
'''End Function
'''
'''Public Function CapturaNFTurnoCorrenteTODOS(strTURNOID As String) As ADODB.Recordset
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''  Dim strData As String
'''  strSql = "SELECT DISTINCT NF.PKID, NF.DATA, NF.SEQUENCIAL, EMPRESA.NOME, ESPECIALIDADE.ESPECIALIDADE, PRESTADOR.NOME, SALA.NUMERO AS SALA_NUMERO,  '(' + FUNC.NIVEL + ') ' + FUNCIONARIO.NOME, NF.STATUS, NF.STATUSIMPRESSAO " & _
'''    "FROM NF INNER JOIN EMPRESA ON EMPRESA.PKID = NF.EMPRESAID " & _
'''    " INNER JOIN TURNO ON TURNO.PKID = NF.TURNOID " & _
'''    "   OR TURNO.PKID = NF.TURNOLABID " & _
'''    " INNER JOIN ATENDE ON ATENDE.PKID = NF.ATENDEID " & _
'''      " INNER JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
'''      " INNER JOIN EMPRESA AS PRESTADOR ON PRESTADOR.PKID = ATENDE.EMPRESAID " & _
'''    " INNER JOIN EMPRESA AS FUNCIONARIO ON FUNCIONARIO.PKID = NF.FUNCIONARIOID " & _
'''    " INNER JOIN FUNCIONARIO AS FUNC ON FUNCIONARIO.PKID = FUNC.EMPRESAID " & _
'''    " INNER JOIN ESPECIALIDADE ON ESPECIALIDADE.PKID = NF.ESPECIALIDADEID " & _
'''    "WHERE NF.TURNOID IN " & _
'''      strTURNOID & _
'''      " ORDER BY NF.DATA DESC, NF.SEQUENCIAL DESC "
'''  Set CapturaNFTurnoCorrenteTODOS = RunSPReturnRS(strSql)
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsNF.CapturaNFTurnoCorrenteTODOS]", _
'''            Err.Description
'''End Function
