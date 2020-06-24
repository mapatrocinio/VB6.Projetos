VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarSaldo(lngSALDOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT SALDO.* " & _
    " FROM SALDO " & _
    " WHERE SALDO.PKID = " & _
    Formata_Dados(lngSALDOID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarSaldo = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsSaldo.SelecionarSaldo]"
End Function
Public Sub ExcluirSaldo(lngSALDOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM SALDO " & _
    " WHERE PKID = " & lngSALDOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsSaldo.ExcluirSaldo]"
End Sub

Public Sub AlterarSaldo(strPercentual As String, _
                        strDescricao As String, _
                        lngSALDOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE SALDO SET " & _
    " DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & _
    ", PERCENTUAL = " & Formata_Dados(strPercentual, tpDados_Moeda, tpNulo_Aceita) & _
    " WHERE PKID = " & lngSALDOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsSaldo.AlterarSaldo]"
End Sub
Public Sub IncluirSaldo(strPercentual As String, _
                        strDescricao As String, _
                        lngParceiroId As Long)
  On Error GoTo trata
  '
  Dim strSql          As String
  '
  strSql = "INSERT INTO SALDO (" & _
    " DESCRICAO, PERCENTUAL, PARCEIROID) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strPercentual, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(lngParceiroId, tpDados_Longo) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsSaldo.IncluirSaldo]"
End Sub

