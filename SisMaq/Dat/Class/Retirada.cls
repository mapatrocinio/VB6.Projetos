VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRetirada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirRetirada(lngRETIRADAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM RETIRADA " & _
    "WHERE RETIRADA.PKID = " & _
    Formata_Dados(lngRETIRADAID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsRetirada.ExcluirRetirada]", _
            Err.Description
End Function


Public Function SelecionarRetiradaPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT RETIRADA.* " & _
    "FROM RETIRADA " & _
    "WHERE RETIRADA.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarRetiradaPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsRetirada.SelecionarRetiradaPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarRetirada(lngPKID As Long, _
                           strValor As String)

  On Error GoTo trata
  '
  Dim strSql As String
  'Inserir Retirada
  strSql = "UPDATE RETIRADA SET " & _
    "VALOR = " & Formata_Dados(strValor, tpDados_Moeda) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsRetirada.AlterarRetirada]", _
            Err.Description
End Sub

Public Sub InserirRetirada(lngTURNOID As Long, _
                           strValor As String, _
                           strStatus As String, _
                           strData As String, _
                           lngGERENTEID As Long)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = strSql & "INSERT INTO RETIRADA (TURNOID, VALOR, STATUS, DATA, GERENTEID) VALUES (" & _
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
            Err.Source & ".[clsRetirada.InserirRetirada]", _
            Err.Description
End Sub





