VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarVara(lngVARAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT VARA.* " & _
    " FROM VARA " & _
    " WHERE VARA.PKID = " & _
    Formata_Dados(lngVARAID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarVara = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsVara.SelecionarVara]"
End Function
Public Sub ExcluirVara(lngVARAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM VARA " & _
    " WHERE PKID = " & lngVARAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsVara.ExcluirVara]"
End Sub

Public Sub AlterarVara(strValor As String, _
                       strNome As String, _
                       lngVARAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE VARA SET " & _
    " NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & _
    ", VALOR = " & Formata_Dados(strValor, tpDados_Moeda, tpNulo_Aceita) & _
    " WHERE PKID = " & lngVARAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsVara.AlterarVara]"
End Sub
Public Sub IncluirVara(strValor As String, _
                       strNome As String)
  On Error GoTo trata
  '
  Dim strSql          As String
  '
  strSql = "INSERT INTO VARA (" & _
    " NOME, VALOR) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strValor, tpDados_Moeda, tpNulo_Aceita) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsVara.IncluirVara]"
End Sub

