VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirEntrada(lngENTRADAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM ENTRADA " & _
    "WHERE ENTRADA.PKID = " & _
    Formata_Dados(lngENTRADAID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntrada.ExcluirEntrada]", _
            Err.Description
End Function


Public Function SelecionarEntradaPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  'strSql = "SELECT ENTRADA.*, PESSOA.NOME AS DESC_ATENDENTE " & _
    "FROM ENTRADA " & _
    " INNER JOIN PESSOA WHERE PESSOA.PKID = ENTRADA.ATENDENTEID " & _
    "WHERE ENTRADA.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  strSql = "SELECT ENTRADA.* " & _
    "FROM ENTRADA " & _
    "WHERE ENTRADA.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarEntradaPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntrada.SelecionarEntradaPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarEntrada(lngPKID As Long, _
                          strValor As String)

  On Error GoTo trata
  '
  Dim strSql As String
  'Inserir Entrada
  strSql = "UPDATE ENTRADA SET " & _
    "VALOR = " & Formata_Dados(strValor, tpDados_Moeda) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntrada.AlterarEntrada]", _
            Err.Description
End Sub

Public Sub InserirEntrada(lngTURNOID As Long, _
                          strValor As String, _
                          strStatus As String, _
                          strData As String, _
                          lngGERENTEID As Long)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = strSql & "INSERT INTO ENTRADA (TURNOID, VALOR, STATUS, DATA, GERENTEID) VALUES (" & _
    Formata_Dados(lngTURNOID, tpDados_Longo) & ", " & _
    Formata_Dados(strValor, tpDados_Moeda) & ", " & _
    Formata_Dados(strStatus, tpDados_Texto) & ", " & _
    Formata_Dados(strData, tpDados_DataHora) & ", " & _
    Formata_Dados(lngGERENTEID, tpDados_Longo) & ");"
    
  '
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntrada.InserirEntrada]", _
            Err.Description
End Sub





