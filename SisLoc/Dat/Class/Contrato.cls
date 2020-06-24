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


Public Function CapturaContrato(strCodigo As String, _
                                strEmpresa As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql    As String
  Dim strWhere  As String
  
  strWhere = "  WHERE "
  If Len(strCodigo & "") <> 0 Then
    strWhere = strWhere & " CONTRATO.NUMERO LIKE " & Formata_Dados(strCodigo & "%", tpDados_Texto)
  End If
  If Len(strEmpresa & "") <> 0 Then
    If strWhere <> "  WHERE " Then strWhere = strWhere & " AND "
    strWhere = strWhere & " EMPRESA.NOME LIKE " & Formata_Dados(strEmpresa & "%", tpDados_Texto)
  End If
  '
  strSql = "SELECT EMPRESA.NOME, CONTRATO.NUMERO " & _
    "FROM CONTRATO " & _
    "LEFT JOIN EMPRESA ON EMPRESA.PKID = CONTRATO.EMPRESAID " & _
    strWhere & _
    " ORDER BY EMPRESA.NOME, CONTRATO.NUMERO"
            
  Set CapturaContrato = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsNF.CapturaContrato]", _
            Err.Description
End Function

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

Public Function SelecionarContrato(lngCONTRATOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CONTRATO.*, PESSOA.NOME AS FUNCIONARIO, EMPRESA.NOME AS NOME_EMPRESA " & _
    "FROM CONTRATO LEFT JOIN PESSOA ON PESSOA.PKID = CONTRATO.PESSOAID " & _
    " LEFT JOIN EMPRESA ON EMPRESA.PKID = CONTRATO.EMPRESAID " & _
    " LEFT JOIN OBRA ON CONTRATO.PKID = OBRA.CONTRATOID " & _
    "WHERE CONTRATO.PKID = " & _
    Formata_Dados(lngCONTRATOID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarContrato = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContrato.SelecionarContrato]", _
            Err.Description
End Function

Public Function SelecionarContratoPeloPkid(lngCONTRATOID As Long, _
                                           lngOBRAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CONTRATO.*, PESSOA.NOME AS FUNCIONARIO, EMPRESA.NOME AS NOME_EMPRESA, OBRA.DESCRICAO AS OBRA_DESCRICAO " & _
    "FROM CONTRATO LEFT JOIN PESSOA ON PESSOA.PKID = CONTRATO.PESSOAID " & _
    " LEFT JOIN EMPRESA ON EMPRESA.PKID = CONTRATO.EMPRESAID " & _
    " LEFT JOIN OBRA ON CONTRATO.PKID = OBRA.CONTRATOID " & _
    "WHERE CONTRATO.PKID = " & _
    Formata_Dados(lngCONTRATOID, tpDados_Longo, tpNulo_Aceita) & _
    " AND OBRA.PKID = " & Formata_Dados(lngOBRAID, tpDados_Longo)
  Set SelecionarContratoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContrato.SelecionarContratoPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarContrato(lngPKID As Long, _
                           strDataInicial As String, _
                           strDataFim As String, _
                           strFuncionarioId As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE CONTRATO SET " & _
    "DATAINICIO = " & Formata_Dados(strDataInicial, tpDados_DataHora) & ", " & _
    "DATAFIM = " & Formata_Dados(strDataFim, tpDados_DataHora) & ", " & _
    "PESSOAID = " & Formata_Dados(strFuncionarioId, tpDados_Longo) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContrato.AlterarContrato]", _
            Err.Description
End Sub

Public Sub InserirContrato(strNumero As String, _
                           strSequencial As String, _
                           strAno As String, _
                           strDataInicial As String, _
                           strDataFim As String, _
                           strEmpresaId As String, _
                           strFuncionarioId As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO CONTRATO (NUMERO, SEQUENCIAL, ANO, DATAINICIO, DATAFIM, EMPRESAID, PESSOAID) VALUES (" & _
    Formata_Dados(strNumero, tpDados_Texto) & ", " & _
    Formata_Dados(strSequencial, tpDados_Longo) & ", " & _
    Formata_Dados(strAno, tpDados_Longo) & ", " & _
    Formata_Dados(strDataInicial, tpDados_DataHora) & ", " & _
    Formata_Dados(strDataFim, tpDados_DataHora) & ", " & _
    Formata_Dados(strEmpresaId, tpDados_Longo) & ", " & _
    Formata_Dados(strFuncionarioId, tpDados_Longo) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContrato.InserirContrato]", _
            Err.Description
End Sub



