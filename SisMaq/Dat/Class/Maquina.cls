VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaquina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirEquipamento(lngEQUIPAMENTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM EQUIPAMENTO " & _
    "WHERE EQUIPAMENTO.PKID = " & _
    Formata_Dados(lngEQUIPAMENTOID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEquipamento.ExcluirEquipamento]", _
            Err.Description
End Function


Public Function SelecionarEquipamentoPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT EQUIPAMENTO.*, MAQUINA.TIPOID, TIPO.TIPO AS DESC_TIPO, MAQUINA.INICIO  " & _
    "FROM EQUIPAMENTO INNER JOIN MAQUINA ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
    "         AND MAQUINA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
    " INNER JOIN TIPO ON TIPO.PKID = MAQUINA.TIPOID " & _
    "WHERE EQUIPAMENTO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarEquipamentoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEquipamento.SelecionarEquipamentoPeloPkid]", _
            Err.Description
End Function

Public Function SelecionarEquipamentoPeloNumero(strNumero As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  strSql = "SELECT EQUIPAMENTO.*, TIPO.DESCRICAO AS DESC_TIPO, MAQUINA.INICIO  " & _
    "FROM EQUIPAMENTO INNER JOIN MAQUINA ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
    "         AND MAQUINA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
    " INNER JOIN TIPO ON TIPO.PKID = MAQUINA.TIPOID " & _
    "WHERE EQUIPAMENTO.NUMERO = " & _
    Formata_Dados(strNumero, tpDados_Texto)
  '
  Set SelecionarEquipamentoPeloNumero = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEquipamento.SelecionarEquipamentoPeloNumero]", _
            Err.Description
End Function

Public Sub AlterarMaquina(lngPKID As Long, _
                          lngTIPOID As Long, _
                          strDataTermino As String, _
                          strStatus As String, _
                          Optional strUsuario As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE MAQUINA SET " & _
    "TIPOID = " & Formata_Dados(lngTIPOID, tpDados_Longo) & ", " & _
    "TERMINO = " & Formata_Dados(strDataTermino, tpDados_DataHora) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & ", " & _
    "USUARIO = " & Formata_Dados(strUsuario & "", tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsMaquina.AlterarMaquina]", _
            Err.Description
End Sub

Public Sub AlterarEquipamento(lngPKID As Long, _
                              lngSERIEID As Long, _
                              strNumero As String, _
                              strCoeficiente As String, _
                              strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE EQUIPAMENTO SET " & _
    "SERIEID = " & Formata_Dados(lngSERIEID, tpDados_Longo) & ", " & _
    "NUMERO = " & Formata_Dados(strNumero, tpDados_Texto) & ", " & _
    "COEFICIENTE = " & Formata_Dados(strCoeficiente, tpDados_Moeda) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEquipamento.AlterarEquipamento]", _
            Err.Description
End Sub

Public Sub InserirMaquina(lngEQUIPAMENTOID As Long, _
                          lngTIPOID As Long, _
                          strInicio As String, _
                          strStatus As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = strSql & "INSERT INTO MAQUINA (EQUIPAMENTOID, TIPOID, INICIO, STATUS) VALUES (" & _
    Formata_Dados(lngEQUIPAMENTOID, tpDados_Longo) & ", " & _
    Formata_Dados(lngTIPOID, tpDados_Longo) & ", " & _
    Formata_Dados(strInicio, tpDados_DataHora) & ", " & _
    Formata_Dados(strStatus, tpDados_Texto) & ");"
  '
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsMaquina.InserirMaquina]", _
            Err.Description
End Sub

Public Sub InserirEquipamento(ByRef lngPKID As Long, _
                              lngSERIEID As Long, _
                              strNumero As String, _
                              strCoeficiente As String, _
                              strStatus As String)

  On Error GoTo trata
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO EQUIPAMENTO (SERIEID, NUMERO, COEFICIENTE, STATUS) VALUES (" & _
    Formata_Dados(lngSERIEID, tpDados_Longo) & ", " & _
    Formata_Dados(strNumero, tpDados_Texto) & ", " & _
    Formata_Dados(strCoeficiente, tpDados_Moeda) & ", " & _
    Formata_Dados(strStatus, tpDados_Texto) & ");"
  '
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngPKID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEquipamento.InserirEquipamento]", _
            Err.Description
End Sub





