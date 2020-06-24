VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUnidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirUnidade(strUnidade As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO UNIDADE (UNIDADE) VALUES (" & _
    Formata_Dados(strUnidade, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsUnidade.InserirUnidade]", _
            Err.Description
End Sub

Public Sub AlterarUnidade(lngUNIDADEID As Long, _
                          strUnidade As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE UNIDADE SET UNIDADE = " & Formata_Dados(strUnidade, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngUNIDADEID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsUnidade.AlterarUnidade]", _
            Err.Description
End Sub

Public Function ListarUnidade(lngUNIDADEID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT UNIDADE.* " & _
    "FROM UNIDADE " & _
    "WHERE UNIDADE.PKID = " & _
    Formata_Dados(lngUNIDADEID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarUnidade = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsUnidade.ListarUnidade]", _
            Err.Description
End Function

Public Sub ExlcuirUnidade(lngUNIDADEID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM UNIDADE WHERE PKID = " & _
    Formata_Dados(lngUNIDADEID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsUnidade.ExlcuirUnidade]", _
            Err.Description
End Sub


Public Function ListarUnidadePelaDesc(strUnidade As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT UNIDADE.* " & _
    "FROM UNIDADE " & _
    "WHERE UNIDADE.UNIDADE = " & _
    Formata_Dados(strUnidade, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarUnidadePelaDesc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsUnidade.ListarUnidadePelaDesc]", _
            Err.Description
End Function



