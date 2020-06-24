VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaixaArrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub FecharCaixaArrec(lngTURNOFECHAID As Long, _
                            strValorFecha As String, _
                            lngTURNOCAIXAARRECID As Long)
  On Error GoTo trata
  Dim strSql  As String
  Dim strStatus As String
  Dim strData   As String
  Dim objRs   As ADODB.Recordset
  strData = Format(Now, "DD/MM/YYYY hh:mm")
  'TURNO
  strSql = strSql & "UPDATE CAIXAARREC SET " & _
    " TURNOFECHAID = " & Formata_Dados(lngTURNOFECHAID, tpDados_Longo) & _
    ", VALORDEVOL = " & Formata_Dados(strValorFecha, tpDados_Moeda) & _
    " WHERE PKID = " & Formata_Dados(lngTURNOCAIXAARRECID, tpDados_Longo)
  '
  RunSP strSql
  'BOLETO ARRECADADOR
  strSql = "SELECT BOLETOARREC.PKID, ISNULL(COUNT(CREDITO.PKID),0) AS LANCADO, 10 - ISNULL(COUNT(CREDITO.PKID),0) AS ALANC "
  strSql = strSql & " FROM " & _
          " BOLETOARREC " & _
          " INNER JOIN CAIXAARREC ON CAIXAARREC.PKID = BOLETOARREC.CAIXAARRECID " & _
          " INNER JOIN PESSOA ON PESSOA.PKID = CAIXAARREC.ARRECADADORID " & _
          " LEFT JOIN CREDITO ON BOLETOARREC.PKID = CREDITO.BOLETOARRECID " & _
          " WHERE CAIXAARREC.PKID = " & Formata_Dados(lngTURNOCAIXAARRECID, tpDados_Longo) & _
          " GROUP BY BOLETOARREC.PKID "
  '
  Set objRs = RunSPReturnRS(strSql)
  Do While Not objRs.EOF
    
    If objRs.Fields("LANCADO").Value = 10 Then
      strStatus = "F"
    ElseIf objRs.Fields("LANCADO").Value = 0 Then
      strStatus = "C"
    ElseIf objRs.Fields("LANCADO").Value < 10 Then
      strStatus = "P"
    Else
      strStatus = "I"
    End If
    strSql = strSql & "UPDATE BOLETOARREC SET " & _
      " TURNODEVOLID = " & Formata_Dados(lngTURNOFECHAID, tpDados_Longo) & _
      ", DATADEVOL = " & Formata_Dados(strData, tpDados_DataHora) & _
      ", STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & _
      " WHERE PKID = " & Formata_Dados(objRs.Fields("PKID").Value, tpDados_Longo)
    RunSP strSql
    objRs.MoveNext
  Loop
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCaixaArrec.InserirCaixaArrec]", _
            Err.Description
End Sub

Public Sub InserirCaixaArrec(ByRef lngPKID As Long, _
                             lngARRECADADORID As Long, _
                             lngTURNOENTRADAID As Long)
  On Error GoTo trata
  Dim strSql  As String
  Dim objRs   As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO CAIXAARREC (ARRECADADORID, TURNOENTRADAID) VALUES (" & _
    " " & Formata_Dados(lngARRECADADORID, tpDados_Longo) & _
    ", " & Formata_Dados(lngTURNOENTRADAID, tpDados_Longo) & ");"
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngPKID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCaixaArrec.InserirCaixaArrec]", _
            Err.Description
End Sub


