VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarCor(lngCORID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT COR.* " & _
    " FROM COR " & _
    " WHERE COR.PKID = " & _
    Formata_Dados(lngCORID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarCor = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsCor.SelecionarCor]"
End Function
Public Sub ExcluirCor(lngCORID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM COR " & _
    " WHERE PKID = " & lngCORID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsCor.ExcluirCor]"
End Sub

Public Sub AlterarCor(strSigla As String, _
                      strNome As String, _
                      lngCORID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE COR SET " & _
    " NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & _
    ", SIGLA = " & Formata_Dados(strSigla, tpDados_Texto, tpNulo_Aceita) & _
    " WHERE PKID = " & lngCORID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsCor.AlterarCor]"
End Sub
Public Sub IncluirCor(strSigla As String, _
                      strNome As String)
  On Error GoTo trata
  '
  Dim strSql          As String
  '
  strSql = "INSERT INTO COR (" & _
    " NOME, SIGLA) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strSigla, tpDados_Texto, tpNulo_Aceita) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsCor.IncluirCor]"
End Sub

