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

Public Function ExecutarSQL(ByVal strSql As String) As ADODB.Recordset
  '
  On Error GoTo trata
  '
  Set ExecutarSQL = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, "[" & strClassName & ".clsGeral.ExecutarSQL]", Err.Description
End Function

Public Sub ExecutarSQLAtualizacao(ByVal strSql As String)
  '
  On Error GoTo trata
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsGeral.ExecutarSQLAtualizacao]"
End Sub

