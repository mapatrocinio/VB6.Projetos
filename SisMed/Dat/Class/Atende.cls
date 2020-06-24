VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAtende"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirAtende(lngAtendeId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
'  strSql = "DELETE FROM ATENDE " & _
'    "WHERE ATENDE.PKID = " & _
'    Formata_Dados(lngAtendeId, tpDados_Longo, tpNulo_Aceita)
  strSql = "UPDATE ATENDE " & _
    "SET STATUS = " & Formata_Dados("I", tpDados_Texto) & _
    "WHERE ATENDE.PKID = " & _
    Formata_Dados(lngAtendeId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAtende.ExcluirAtende]", _
            Err.Description
End Function

Public Function SelecionarAtendePeloPkid(lngAtendeId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ATENDE.*, PRONTUARIO.NOME, DIASDASEMANA.DIADASEMANA " & _
    "FROM ATENDE " & _
    " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = ATENDE.PRONTUARIOID " & _
    " INNER JOIN DIASDASEMANA ON DIASDASEMANA.PKID = ATENDE.DIASDASEMANAID " & _
    "WHERE ATENDE.PKID = " & _
    Formata_Dados(lngAtendeId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarAtendePeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAtende.SelecionarAtendePeloPkid]", _
            Err.Description
End Function


Public Sub AlterarAtende(lngPKID As Long, _
                         lngSalaId As Long, _
                         lngProntuarioId As Long, _
                         lngDiasDaSemanaId As Long, _
                         strDataInicio As String, _
                         strDataFim As String, _
                         strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE ATENDE SET " & _
    "SALAID = " & Formata_Dados(lngSalaId, tpDados_Longo) & ", " & _
    "PRONTUARIOID = " & Formata_Dados(lngProntuarioId, tpDados_Longo) & ", " & _
    "DIASDASEMANAID = " & Formata_Dados(lngDiasDaSemanaId, tpDados_Longo) & ", " & _
    "HORAINICIO = " & Formata_Dados(strDataInicio, tpDados_Texto) & ", " & _
    "HORATERMINO = " & Formata_Dados(strDataFim, tpDados_Texto) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAtende.AlterarAtende]", _
            Err.Description
End Sub

Public Sub InserirAtende(lngSalaId As Long, _
                         lngProntuarioId As Long, _
                         lngDiasDaSemanaId As Long, _
                         strDataInicio As String, _
                         strDataFim As String, _
                         strStatus As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO ATENDE (SALAID, PRONTUARIOID, DIASDASEMANAID, HORAINICIO, HORATERMINO, STATUS) VALUES (" & _
    Formata_Dados(lngSalaId, tpDados_Texto) & ", " & _
    Formata_Dados(lngProntuarioId, tpDados_Moeda) & ", " & _
    Formata_Dados(lngDiasDaSemanaId, tpDados_Moeda) & ", " & _
    Formata_Dados(strDataInicio, tpDados_Texto) & ", " & _
    Formata_Dados(strDataFim, tpDados_Texto) & ", " & _
    Formata_Dados("A", tpDados_Texto) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAtende.InserirAtende]", _
            Err.Description
End Sub

Public Function SelecionarAtende(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ATENDE.* " & _
    "FROM ATENDE " & _
    "WHERE ATENDE.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarAtende = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAtende.SelecionarAtende]", _
            Err.Description
End Function

