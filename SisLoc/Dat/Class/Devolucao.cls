VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDevolucao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExlcuirItemDevolucao(lngITEMDEVOLUCAOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM ITEMDEVOLUCAO WHERE PKID = " & _
    Formata_Dados(lngITEMDEVOLUCAOID, tpDados_Longo) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemDevolucao.ExlcuirItemDevolucao]", _
            Err.Description
End Sub

Public Sub InserirItemDevolucao(lngITEMNFID As Long, _
                                lngDEVOLUCAOID As Long, _
                                strQuantidade As String, _
                                strQuantidadeNF As String, _
                                strQuantidadeAvaria As String, _
                                strValorIndenizacaoest As String)
  '
  On Error GoTo trata
  '
  Dim strSql  As String
  '
  strSql = "INSERT INTO ITEMDEVOLUCAO (ITEMNFID, DEVOLUCAOID, QUANTIDADE, QUANTIDADENF, QUANTIDADEAVARIA, VALORINDENIZACAOEST) VALUES (" & _
    " " & Formata_Dados(lngITEMNFID, tpDados_Longo) & _
    ", " & Formata_Dados(lngDEVOLUCAOID, tpDados_Longo) & _
    ", " & Formata_Dados(strQuantidade, tpDados_Longo) & _
    ", " & Formata_Dados(strQuantidadeNF, tpDados_Longo) & _
    ", " & Formata_Dados(strQuantidadeAvaria, tpDados_Longo) & _
    ", " & Formata_Dados(strValorIndenizacaoest, tpDados_Moeda) & _
    ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDevolucao.InserirItemDeolucao]", _
            Err.Description
End Sub

Public Sub AlterarEstoquePelaDevolucao(ByVal lngESTOQUEID As Long, _
                                       ByVal strQuantidade As String, _
                                       ByVal strStatus As String)
  '
  On Error GoTo trata
  '
  Dim strSql  As String
  '
  strSql = "UPDATE ESTOQUE SET QUANTIDADE = ISNULL(QUANTIDADE,0) " & IIf(strStatus = "INC", " + ", " - ") & Formata_Dados(strQuantidade, tpDados_Longo) & " " & _
    "WHERE PKID = " & lngESTOQUEID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDevolucao.AlterarEstoquePelaDevolucao]", _
            Err.Description
End Sub

Public Sub InserirDevolucao(ByRef lngDEVOLUCAOID As Long, _
                            ByRef lngCONTRATOID As Long, _
                            ByRef lngOBRAID As Long, _
                            ByRef strSequencial As String, _
                            strData As String, _
                            strNFCliente As String, _
                            strDataEmissao As String, _
                            strDataDevolucao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO DEVOLUCAO (SEQUENCIAL, DATA, CONTRATOID, OBRAID, NUMERONF, DATAEMISSAO, DATADEVOLUCAO) VALUES (" & _
    Formata_Dados(strSequencial, tpDados_Longo) & _
    ", " & Formata_Dados(strData, tpDados_DataHora) & _
    ", " & Formata_Dados(lngCONTRATOID, tpDados_Longo) & _
    ", " & Formata_Dados(lngOBRAID, tpDados_Longo) & _
    ", " & Formata_Dados(strNFCliente, tpDados_Texto) & _
    ", " & Formata_Dados(strDataEmissao, tpDados_DataHora) & _
    ", " & Formata_Dados(strDataDevolucao, tpDados_DataHora) & _
    "); "
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngDEVOLUCAOID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDevolucao.InserirDevolucao]", _
            Err.Description
End Sub

Public Sub AlterarDevolucao(lngDEVOLUCAOID As Long, _
                            strData As String, _
                            strNFCliente As String, _
                            strDataEmissao As String, _
                            strDataDevolucao As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE DEVOLUCAO SET " & _
    " DATA =  " & Formata_Dados(strData, tpDados_DataHora) & _
    ", NUMERONF =  " & Formata_Dados(strNFCliente, tpDados_Texto) & _
    ", DATAEMISSAO =  " & Formata_Dados(strDataEmissao, tpDados_DataHora) & _
    ", DATADEVOLUCAO =  " & Formata_Dados(strDataDevolucao, tpDados_DataHora) & _
    "WHERE PKID = " & lngDEVOLUCAOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDevolucao.AlterarDevolucao]", _
            Err.Description
End Sub

Public Function ListarDevolucao(lngDEVOLUCAOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT DEVOLUCAO.*, right('000' + convert(varchar(20), SEQUENCIAL), 3) + '-' + right('00' + CONVERT(VARCHAR(2), DAY(DATA)), 2) + '/' + right('00' + CONVERT(VARCHAR(2), MONTH(DATA)), 2) + '/' + right('0000' + CONVERT(VARCHAR(4), YEAR(DATA)), 4) AS DEV " & _
    "FROM DEVOLUCAO " & _
    "WHERE DEVOLUCAO.PKID = " & _
    Formata_Dados(lngDEVOLUCAOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarDevolucao = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDevolucao.ListarDevolucao]", _
            Err.Description
End Function

Public Function ListarDevolucaoPeloSeq(lngCONTRATOID As Long, _
                                       lngOBRAID As Long, _
                                       strSequencial As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT DEVOLUCAO.* " & _
    "FROM DEVOLUCAO " & _
    "WHERE DEVOLUCAO.CONTRATOID = " & _
    Formata_Dados(lngCONTRATOID, tpDados_Longo) & _
    " AND DEVOLUCAO.OBRAID = " & Formata_Dados(lngOBRAID, tpDados_Longo) & _
    " AND SEQUENCIAL = " & Formata_Dados(strSequencial, tpDados_Longo)
  Set ListarDevolucaoPeloSeq = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDevolucao.ListarDevolucaoPeloSeq]", _
            Err.Description
End Function

Public Sub ExlcuirDevolucao(lngDEVOLUCAOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM DEVOLUCAO WHERE PKID = " & _
    Formata_Dados(lngDEVOLUCAOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDevolucao.ExlcuirDevolucao]", _
            Err.Description
End Sub
Public Function CapturaSequencialDevolucao(lngCONTRATOID As Long) As Long
  '
  Dim objRs   As ADODB.Recordset
  Dim lngRet  As Long
  
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "SELECT MAX(SEQUENCIAL) AS COD FROM DEVOLUCAO " & _
        " WHERE CONTRATOID = " & Formata_Dados(lngCONTRATOID, tpDados_Longo)
  
  Set objRs = RunSPReturnRS(strSql)
  If objRs.EOF Then
    lngRet = 1
  ElseIf Not IsNumeric(objRs.Fields("COD").Value) Then
    lngRet = 1
  Else
    lngRet = CLng(objRs.Fields("COD").Value) + 1
  End If
  '
  objRs.Close
  Set objRs = Nothing
  '
  CapturaSequencialDevolucao = lngRet
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDevolucao.CapturaSequencialDevolucao]", _
            Err.Description
End Function

