VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoPgto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub ExcluirTipoPgto(lngTipoPgtoId As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM TIPOPGTO " & _
    " WHERE PKID = " & lngTipoPgtoId
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsTipoPgto.ExcluirTipoPgto]"
End Sub

Public Sub AlterarTipoPgto(lngTipoPgtoId As Long, _
                            strTipoPgto As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE TIPOPGTO SET " & _
    " TIPOPGTO = " & Formata_Dados(strTipoPgto, tpDados_Texto, tpNulo_Aceita) & _
    " WHERE PKID = " & lngTipoPgtoId
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsTipoPgto.AlterarTipoPgto]"
End Sub
Public Sub IncluirTipoPgto(strTipoPgto As String)
  On Error GoTo trata
  '
  Dim strSql          As String
  '
  strSql = "INSERT INTO TIPOPGTO (" & _
    " TIPOPGTO) VALUES (" & _
    Formata_Dados(strTipoPgto, tpDados_Texto, tpNulo_Aceita) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsTipoPgto.IncluirTipoPgto]"
End Sub

Public Function SelecionarTipoPgto(lngTipoPgtoId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPOPGTO.* FROM TIPOPGTO " & _
    " WHERE TIPOPGTO.PKID = " & _
    Formata_Dados(lngTipoPgtoId, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarTipoPgto = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsTipoPgto.SelecionarTipoPgto]"
End Function

