VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormaPgto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub ExcluirFormaPgto(lngFormaPgtoId As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM FORMAPGTO " & _
    " WHERE PKID = " & lngFormaPgtoId
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsFormaPgto.ExcluirFormaPgto]"
End Sub

Public Sub AlterarFormaPgto(lngFormaPgtoId As Long, _
                            strFormaPgto As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE FORMAPGTO SET " & _
    " FORMAPGTO = " & Formata_Dados(strFormaPgto, tpDados_Texto, tpNulo_Aceita) & _
    " WHERE PKID = " & lngFormaPgtoId
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsFormaPgto.AlterarFormaPgto]"
End Sub
Public Sub IncluirFormaPgto(strFormaPgto As String)
  On Error GoTo trata
  '
  Dim strSql          As String
  '
  strSql = "INSERT INTO FORMAPGTO (" & _
    " FORMAPGTO) VALUES (" & _
    Formata_Dados(strFormaPgto, tpDados_Texto, tpNulo_Aceita) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsFormaPgto.IncluirFormaPgto]"
End Sub

Public Function SelecionarFormaPgto(lngFormaPgtoId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT FORMAPGTO.* FROM FORMAPGTO " & _
    " WHERE FORMAPGTO.PKID = " & _
    Formata_Dados(lngFormaPgtoId, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarFormaPgto = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsFormaPgto.SelecionarFormaPgto]"
End Function

