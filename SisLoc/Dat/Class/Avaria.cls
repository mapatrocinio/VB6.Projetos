VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAvaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExlcuirItemAvaria(lngITEMAVARIAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM ITEMAVARIA WHERE PKID = " & _
    Formata_Dados(lngITEMAVARIAID, tpDados_Longo) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemAvaria.ExlcuirItemAvaria]", _
            Err.Description
End Sub

Public Sub InserirItemAvaria(lngITEMNFID As Long, _
                             lngAVARIAID As Long, _
                             strQuantidade As String, _
                             strValorIndenizacao As String)
  '
  On Error GoTo trata
  '
  Dim strSql  As String
  '
  strSql = "INSERT INTO ITEMAVARIA (ITEMNFID, AVARIAID, QUANTIDADE, VALORINDENIZACAOEST) VALUES (" & _
    " " & Formata_Dados(lngITEMNFID, tpDados_Longo) & _
    ", " & Formata_Dados(lngAVARIAID, tpDados_Longo) & _
    ", " & Formata_Dados(strQuantidade, tpDados_Longo) & _
    ", " & Formata_Dados(strValorIndenizacao, tpDados_Moeda) & _
    ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAvaria.InserirItemDeolucao]", _
            Err.Description
End Sub

Public Sub AlterarEstoquePelaAvaria(ByVal lngESTOQUEID As Long, _
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
            Err.Source & ".[clsAvaria.AlterarEstoquePelaAvaria]", _
            Err.Description
End Sub

Public Sub InserirAvaria(ByRef lngAVARIAID As Long, _
                         ByRef lngNFID As Long, _
                         ByRef strSequencial As String, _
                         strData As String, _
                         strDescricao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO AVARIA (SEQUENCIAL, DATA, NFID, DESCRICAO) VALUES (" & _
    Formata_Dados(strSequencial, tpDados_Longo) & _
    ", " & Formata_Dados(strData, tpDados_DataHora) & _
    ", " & Formata_Dados(lngNFID, tpDados_Longo) & _
    ", " & Formata_Dados(strDescricao, tpDados_Texto) & _
    "); "
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngAVARIAID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAvaria.InserirAvaria]", _
            Err.Description
End Sub

Public Sub AlterarAvaria(lngAVARIAID As Long, _
                         strData As String, _
                         strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE AVARIA SET DATA = " & Formata_Dados(strData, tpDados_DataHora) & ", " & _
    "DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto) & _
    " WHERE PKID = " & lngAVARIAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAvaria.AlterarAvaria]", _
            Err.Description
End Sub

Public Function ListarAvaria(lngAVARIAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT AVARIA.*, right('000' + convert(varchar(20), SEQUENCIAL), 3) + '-' + right('00' + CONVERT(VARCHAR(2), DAY(DATA)), 2) + '/' + right('00' + CONVERT(VARCHAR(2), MONTH(DATA)), 2) + '/' + right('0000' + CONVERT(VARCHAR(4), YEAR(DATA)), 4) AS DEV " & _
    "FROM AVARIA " & _
    "WHERE AVARIA.PKID = " & _
    Formata_Dados(lngAVARIAID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarAvaria = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAvaria.ListarAvaria]", _
            Err.Description
End Function

Public Function ListarAvariaPeloSeq(lngNFID As Long, _
                                    strSequencial As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT AVARIA.* " & _
    "FROM AVARIA " & _
    "WHERE AVARIA.NFID = " & _
    Formata_Dados(lngNFID, tpDados_Longo) & _
    " AND SEQUENCIAL = " & Formata_Dados(strSequencial, tpDados_Longo)
  Set ListarAvariaPeloSeq = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAvaria.ListarAvariaPeloSeq]", _
            Err.Description
End Function

Public Sub ExlcuirAvaria(lngAVARIAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM AVARIA WHERE PKID = " & _
    Formata_Dados(lngAVARIAID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAvaria.ExlcuirAvaria]", _
            Err.Description
End Sub
Public Function CapturaSequencialAvaria(lngNFID As Long) As Long
  '
  Dim objRs   As ADODB.Recordset
  Dim lngRet  As Long
  
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "SELECT MAX(SEQUENCIAL) AS COD FROM AVARIA " & _
        " WHERE NFID = " & Formata_Dados(lngNFID, tpDados_Longo)
  
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
  CapturaSequencialAvaria = lngRet
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAvaria.CapturaSequencialAvaria]", _
            Err.Description
End Function

