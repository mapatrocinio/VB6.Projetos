VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEstadoCivil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirEstadoCivil(strDescricao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO ESTADOCIVIL (DESCRICAO) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEstadoCivil.InserirEstadoCivil]", _
            Err.Description
End Sub

Public Sub AlterarEstadoCivil(lngESTADOCIVILID As Long, _
                              strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE ESTADOCIVIL SET DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngESTADOCIVILID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEstadoCivil.AlterarEstadoCivil]", _
            Err.Description
End Sub

Public Function ListarEstadoCivil(lngESTADOCIVILID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ESTADOCIVIL.* " & _
    "FROM ESTADOCIVIL " & _
    "WHERE ESTADOCIVIL.PKID = " & _
    Formata_Dados(lngESTADOCIVILID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarEstadoCivil = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEstadoCivil.ListarEstadoCivil]", _
            Err.Description
End Function

Public Sub ExlcuirEstadoCivil(lngESTADOCIVILID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM ESTADOCIVIL WHERE PKID = " & _
    Formata_Dados(lngESTADOCIVILID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEstadoCivil.ExlcuirEstadoCivil]", _
            Err.Description
End Sub


Public Function ListarEstadoCivilPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ESTADOCIVIL.* " & _
    "FROM ESTADOCIVIL " & _
    "WHERE ESTADOCIVIL.DESCRICAO = " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarEstadoCivilPelaDesc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEstadoCivil.ListarEstadoCivilPelaDesc]", _
            Err.Description
End Function

