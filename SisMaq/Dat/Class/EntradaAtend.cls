VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEntradaAtend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirEntradaAtend(lngENTRADAATENDID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM ENTRADAATEND " & _
    "WHERE ENTRADAATEND.PKID = " & _
    Formata_Dados(lngENTRADAATENDID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaAtend.ExcluirEntradaAtend]", _
            Err.Description
End Function


Public Function SelecionarEntradaAtendPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ENTRADAATEND.*, PESSOA.NOME AS DESC_ATENDENTE " & _
    "FROM ENTRADAATEND " & _
    " INNER JOIN CAIXAATEND ON CAIXAATEND.PKID = ENTRADAATEND.CAIXAATENDID " & _
    " INNER JOIN PESSOA ON PESSOA.PKID = CAIXAATEND.ATENDENTEID " & _
    "WHERE ENTRADAATEND.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarEntradaAtendPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaAtend.SelecionarEntradaAtendPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarEntradaAtend(lngPKID As Long, _
                               strValor As String)

  On Error GoTo trata
  '
  Dim strSql As String
  'Inserir EntradaAtend
  strSql = "UPDATE ENTRADAATEND SET " & _
    "VALOR = " & Formata_Dados(strValor, tpDados_Moeda) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaAtend.AlterarEntradaAtend]", _
            Err.Description
End Sub

Public Sub InserirEntradaAtend(lngTURNOID As Long, _
                               lngCAIXAATENDID As Long, _
                               strValor As String, _
                               strData As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = strSql & "INSERT INTO ENTRADAATEND (TURNOENTRADAID, CAIXAATENDID, VALOR, DATA) VALUES (" & _
    Formata_Dados(lngTURNOID, tpDados_Longo) & ", " & _
    Formata_Dados(lngCAIXAATENDID, tpDados_Longo) & ", " & _
    Formata_Dados(strValor, tpDados_Moeda) & ", " & _
    Formata_Dados(strData, tpDados_DataHora) & ");"
    
  '
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaAtend.InserirEntradaAtend]", _
            Err.Description
End Sub





