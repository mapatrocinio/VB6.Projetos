VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirIPI(strIPI As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO IPI (IPI) VALUES (" & _
    Formata_Dados(strIPI, tpDados_Moeda) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsIPI.InserirIPI]", _
            Err.Description
End Sub

Public Sub AlterarIPI(lngIPIID As Long, _
                      strIPI As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE IPI SET IPI = " & Formata_Dados(strIPI, tpDados_Moeda) & " " & _
    "WHERE PKID = " & lngIPIID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsIPI.AlterarIPI]", _
            Err.Description
End Sub

Public Function ListarIPI(lngIPIID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT IPI.* " & _
    "FROM IPI " & _
    "WHERE IPI.PKID = " & _
    Formata_Dados(lngIPIID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarIPI = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsIPI.ListarIPI]", _
            Err.Description
End Function

Public Sub ExcluirIPI(lngIPIID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui FAMÍLIA DE PRODUTOS
  strSql = "DELETE FROM IPI WHERE IPI.PKID = " & _
    Formata_Dados(lngIPIID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsIPI.ExcluirIPI]", _
            Err.Description
End Sub






