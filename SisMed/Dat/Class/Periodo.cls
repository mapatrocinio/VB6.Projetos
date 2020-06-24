VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirPeriodo(strPeriodo As String, _
                          strInicio As String, _
                          strTermino As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO PERIODO (PERIODO, INICIO, TERMINO) VALUES (" & _
    Formata_Dados(strPeriodo, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strInicio, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strTermino, tpDados_Texto, tpNulo_Aceita) & ")"
    
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPeriodo.InserirPeriodo]", _
            Err.Description
End Sub

Public Sub AlterarPeriodo(lngPERIODOID As Long, _
                          strPeriodo As String, _
                          strInicio As String, _
                          strTermino As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE PERIODO SET PERIODO = " & Formata_Dados(strPeriodo, tpDados_Longo, tpNulo_Aceita) & ", " & _
    "INICIO = " & Formata_Dados(strInicio, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "TERMINO = " & Formata_Dados(strTermino, tpDados_Texto, tpNulo_Aceita) & " " & _
    "WHERE PKID = " & lngPERIODOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPeriodo.AlterarPeriodo]", _
            Err.Description
End Sub

Public Function ListarPeriodo(lngPERIODOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PERIODO.* " & _
    "FROM PERIODO " & _
    "WHERE PERIODO.PKID = " & _
    Formata_Dados(lngPERIODOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarPeriodo = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPeriodo.ListarPeriodo]", _
            Err.Description
End Function

Public Sub ExlcuirPeriodo(lngPERIODOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM PERIODO WHERE PKID = " & _
    Formata_Dados(lngPERIODOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPeriodo.ExlcuirPeriodo]", _
            Err.Description
End Sub

