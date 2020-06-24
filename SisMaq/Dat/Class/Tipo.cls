VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirTipo(strTipo As String, _
                       strDescricao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO TIPO (TIPO, DESCRICAO) VALUES (" & _
    Formata_Dados(strTipo, tpDados_Texto) & _
    ", " & Formata_Dados(strDescricao, tpDados_Texto) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipo.InserirTipo]", _
            Err.Description
End Sub

Public Sub AlterarTipo(lngTIPOID As Long, _
                       strTipo As String, _
                       strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE TIPO SET TIPO = " & _
    Formata_Dados(strTipo, tpDados_Texto) & _
    ", DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto) & _
    " WHERE PKID = " & lngTIPOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipo.AlterarTipo]", _
            Err.Description
End Sub

Public Function ListarTipo(lngTIPOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPO.* " & _
    "FROM TIPO " & _
    "WHERE TIPO.PKID = " & _
    Formata_Dados(lngTIPOID, tpDados_Longo)
  Set ListarTipo = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipo.ListarTipo]", _
            Err.Description
End Function

Public Sub ExlcuirTipo(lngTIPOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM TIPO WHERE PKID = " & _
    Formata_Dados(lngTIPOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipo.ExlcuirTipo]", _
            Err.Description
End Sub
