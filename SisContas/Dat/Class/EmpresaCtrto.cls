VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpresaCtrto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarEmpresaCtrtoPelaDescricao(lngCONTRATOID As Long, _
                                                    strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT EMPRESACTRTO.* " & _
    "FROM EMPRESACTRTO INNER JOIN EMPRESA " & _
    " ON EMPRESA.PKID = EMPRESACTRTO.EMPRESAID " & _
    "WHERE EMPRESACTRTO.DESCRICAO = " & _
      Formata_Dados(strDescricao, tpDados_Texto) & _
      " AND EMPRESACTRTO.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
      " AND EMPRESACTRTO.CONTRATOID = " & Formata_Dados(lngCONTRATOID, tpDados_Longo)
  Set SelecionarEmpresaCtrtoPelaDescricao = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresaCtrto.SelecionarEmpresaCtrtoPelaDescricao]", _
            Err.Description
End Function

Public Function ExcluirEmpresaCtrto(lngEMPRESACTRTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM EMPRESACTRTO " & _
    "WHERE EMPRESACTRTO.PKID = " & _
    Formata_Dados(lngEMPRESACTRTOID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresaCtrto.ExcluirEmpresaCtrto]", _
            Err.Description
End Function

Public Function SelecionarEmpresaCtrtoPeloPkid(lngEMPRESACTRTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT EMPRESACTRTO.*, EMPRESA.NOME AS NOME_EMPRESA " & _
    "FROM EMPRESACTRTO INNER JOIN EMPRESA ON EMPRESA.PKID = EMPRESACTRTO.EMPRESAID " & _
    "WHERE EMPRESACTRTO.PKID = " & _
    Formata_Dados(lngEMPRESACTRTOID, tpDados_Longo)
  Set SelecionarEmpresaCtrtoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresaCtrto.SelecionarEmpresaCtrtoPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarEmpresaCtrto(lngPKID As Long, _
                               lngEMPRESAID As Long, _
                               strDescricao As String, _
                               strValor As String, _
                               strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE EMPRESACTRTO SET " & _
    "EMPRESAID = " & Formata_Dados(lngEMPRESAID, tpDados_Longo) & ", " & _
    "DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    "VALOR = " & Formata_Dados(strValor, tpDados_Moeda) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresaCtrto.AlterarEmpresaCtrto]", _
            Err.Description
End Sub

Public Sub InserirEmpresaCtrto(lngCONTRATOID As Long, _
                               lngEMPRESAID As Long, _
                               strDescricao As String, _
                               strValor As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO EMPRESACTRTO (CONTRATOID, EMPRESAID, DESCRICAO, VALOR, STATUS) VALUES (" & _
    Formata_Dados(lngCONTRATOID, tpDados_Longo) & ", " & _
    Formata_Dados(lngEMPRESAID, tpDados_Longo) & ", " & _
    Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    Formata_Dados(strValor, tpDados_Moeda) & ", " & _
    Formata_Dados("A", tpDados_Texto) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresaCtrto.InserirEmpresaCtrto]", _
            Err.Description
End Sub

Public Function SelecionarEmpresaCtrto(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT EMPRESACTRTO.* " & _
    "FROM EMPRESACTRTO " & _
    "WHERE EMPRESACTRTO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarEmpresaCtrto = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresaCtrto.SelecionarEmpresaCtrto]", _
            Err.Description
End Function

