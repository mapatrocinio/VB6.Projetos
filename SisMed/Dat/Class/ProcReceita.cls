VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcReceita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirProcReceita(lngProcReceitaId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM PROCEDIMENTORECEITA " & _
    "WHERE PROCEDIMENTORECEITA.PKID = " & _
    Formata_Dados(lngProcReceitaId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProcReceita.ExcluirProcReceita]", _
            Err.Description
End Function

Public Function SelecionarProcReceitaPeloPkid(lngProcReceitaId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PROCEDIMENTORECEITA.*, PROCEDIMENTO.PROCEDIMENTO " & _
    "FROM PROCEDIMENTORECEITA INNER JOIN PROCEDIMENTO ON PROCEDIMENTO.PKID = PROCEDIMENTORECEITA.PROCEDIMENTOID " & _
    "WHERE PROCEDIMENTORECEITA.PKID = " & _
    Formata_Dados(lngProcReceitaId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarProcReceitaPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProcReceita.SelecionarProcReceitaPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarProcReceita(lngPKID As Long, _
                              strTipo As String, _
                              strDescricao As String, _
                              strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE PROCEDIMENTORECEITA SET " & _
    "TIPO = " & Formata_Dados(strTipo, tpDados_Texto) & ", " & _
    "DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProcReceita.AlterarProcReceita]", _
            Err.Description
End Sub

Public Sub InserirProcReceita(lngPROCEDIMENTOID As Long, _
                              strTipo As String, _
                              strDescricao As String, _
                              strStatus As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO PROCEDIMENTORECEITA (PROCEDIMENTOID, TIPO, DESCRICAO, STATUS) VALUES (" & _
    Formata_Dados(lngPROCEDIMENTOID, tpDados_Longo) & ", " & _
    Formata_Dados(strTipo, tpDados_Texto) & ", " & _
    Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    Formata_Dados(strStatus, tpDados_Texto) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProcReceita.InserirProcReceita]", _
            Err.Description
End Sub

Public Function SelecionarProcReceita(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PROCEDIMENTORECEITA.* " & _
    "FROM PROCEDIMENTORECEITA " & _
    "WHERE PROCEDIMENTORECEITA.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarProcReceita = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProcReceita.SelecionarProcReceita]", _
            Err.Description
End Function

