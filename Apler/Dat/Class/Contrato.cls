VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit

Public Function ExcluirContrato(lngCONTRATOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM CONTRATO " & _
    "WHERE CONTRATO.PKID = " & _
    Formata_Dados(lngCONTRATOID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContrato.ExcluirContrato]", _
            Err.Description
End Function

Public Function SelecionarContratoPeloPkid(lngCONTRATOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CONTRATO.* " & _
    "FROM CONTRATO " & _
    "WHERE CONTRATO.PKID = " & _
    Formata_Dados(lngCONTRATOID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarContratoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContrato.SelecionarContratoPeloPkid]", _
            Err.Description
End Function

Public Function SelecionarContratoPelaDescricao(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CONTRATO.* " & _
    "FROM CONTRATO " & _
    "WHERE CONTRATO.DESCRICAO = " & _
      Formata_Dados(strDescricao, tpDados_Texto) & _
      " AND CONTRATO.STATUS = " & Formata_Dados("A", tpDados_Texto)
  Set SelecionarContratoPelaDescricao = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContrato.SelecionarContratoPelaDescricao]", _
            Err.Description
End Function


Public Sub AlterarContrato(lngPKID As Long, _
                        strDescricao As String, _
                        strObservacao As String, _
                        strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE CONTRATO SET " & _
    "DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    "OBSERVACAO = " & Formata_Dados(strObservacao, tpDados_Texto) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContrato.AlterarContrato]", _
            Err.Description
End Sub

Public Sub InserirContrato(strDescricao As String, _
                           strObservacao As String, _
                           lngParceiroId As Long)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO CONTRATO (DESCRICAO, OBSERVACAO, STATUS, PARCEIROID) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    Formata_Dados(strObservacao, tpDados_Texto) & ", " & _
    Formata_Dados("A", tpDados_Texto) & ", " & _
    Formata_Dados(lngParceiroId, tpDados_Longo) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContrato.InserirContrato]", _
            Err.Description
End Sub

Public Function SelecionarContrato(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CONTRATO.* " & _
    "FROM CONTRATO " & _
    "WHERE CONTRATO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarContrato = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContrato.SelecionarContrato]", _
            Err.Description
End Function



