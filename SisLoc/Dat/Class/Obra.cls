VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function CapturaObra(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql    As String
  Dim strWhere  As String
  
  strWhere = "  WHERE "
  If strWhere <> "  WHERE " Then strWhere = strWhere & " AND "
  strWhere = strWhere & " OBRA.DESCRICAO LIKE " & Formata_Dados(strDescricao & "%", tpDados_Texto)
  '
  strSql = "SELECT OBRA.DESCRICAO, CONTRATO.NUMERO, EMPRESA.NOME AS EMPRESA_NOME " & _
    "FROM CONTRATO " & _
    "INNER JOIN OBRA ON OBRA.CONTRATOID = CONTRATO.PKID " & _
    "INNER JOIN EMPRESA ON CONTRATO.EMPRESAID = EMPRESA.PKID " & _
    strWhere & _
    " ORDER BY OBRA.DESCRICAO"
            
  Set CapturaObra = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsObra.CapturaObra]", _
            Err.Description
End Function

Public Function ExcluirObra(lngOBRAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM OBRA " & _
    "WHERE OBRA.PKID = " & _
    Formata_Dados(lngOBRAID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsObra.ExcluirObra]", _
            Err.Description
End Function

Public Function SelecionarObraPeloPkid(lngOBRAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT OBRA.* " & _
    "FROM OBRA " & _
    "WHERE OBRA.PKID = " & _
    Formata_Dados(lngOBRAID, tpDados_Longo)
  Set SelecionarObraPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsObra.SelecionarObraPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarObra(lngPKID As Long, _
                       strDescricao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE OBRA SET " & _
    "DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsObra.AlterarObra]", _
            Err.Description
End Sub

Public Sub InserirObra(lngCONTRATOID As Long, _
                       strDescricao As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO OBRA (CONTRATOID, DESCRICAO) VALUES (" & _
    Formata_Dados(lngCONTRATOID, tpDados_Longo) & ", " & _
    Formata_Dados(strDescricao, tpDados_Texto) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsObra.InserirObra]", _
            Err.Description
End Sub

Public Function SelecionarObra(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT OBRA.* " & _
    "FROM OBRA " & _
    "WHERE OBRA.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarObra = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsObra.SelecionarObra]", _
            Err.Description
End Function

