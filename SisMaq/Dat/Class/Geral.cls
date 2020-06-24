VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ObterConnectionString() As String
  '
  On Error GoTo trata
  '
  ObterConnectionString = GetConnectionString
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGeral.ObterConnectionString]", _
            Err.Description
End Function

Public Function ExecutarSQLRetInteger(ByVal strSql As String, _
                                      params As Variant) As Long 'adInterger is really a VB Long
  '
  On Error GoTo trata
  '
  ExecutarSQLRetInteger = RunSPReturnInteger(strSql, _
                                             params)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGeral.ExecutarSQLRetInteger]", _
            Err.Description
End Function

Public Function ExecutarSQL(ByVal strSql As String) As ADODB.Recordset
  '
  On Error GoTo trata
  '
  Set ExecutarSQL = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGeral.ExecutarSQL]", _
            Err.Description
End Function

Public Sub ExecutarSQLAtualizacao(ByVal strSql As String)
  '
  On Error GoTo trata
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGeral.ExecutarSQLAtualizacao]", _
            Err.Description
End Sub

