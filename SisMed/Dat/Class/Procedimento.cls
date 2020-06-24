VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcedimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirProcedimento(strProcedimento As String, _
                               strTipoProcedimentoId As String, _
                               strValor As String, _
                               strAceitaValor As String, _
                               strConsulta As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO PROCEDIMENTO (PROCEDIMENTO, TIPOPROCEDIMENTOID, VALOR, INDACEITAVALOR, INDCONSULTA) VALUES (" & _
    Formata_Dados(strProcedimento, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strTipoProcedimentoId, tpDados_Longo) & _
    ", " & Formata_Dados(strValor, tpDados_Moeda) & _
    ", " & Formata_Dados(strAceitaValor, tpDados_Texto) & _
    ", " & Formata_Dados(strConsulta, tpDados_Texto) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProcedimento.InserirProcedimento]", _
            Err.Description
End Sub

Public Sub AlterarProcedimento(lngPROCEDIMENTOID As Long, _
                               strProcedimento As String, _
                               strTipoProcedimentoId As String, _
                               strValor As String, _
                               strAceitaValor As String, _
                               strConsulta As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE PROCEDIMENTO SET PROCEDIMENTO = " & Formata_Dados(strProcedimento, tpDados_Texto) & _
    ", TIPOPROCEDIMENTOID = " & Formata_Dados(strTipoProcedimentoId, tpDados_Longo) & _
    ", VALOR = " & Formata_Dados(strValor, tpDados_Moeda) & _
    ", INDACEITAVALOR = " & Formata_Dados(strAceitaValor, tpDados_Texto) & _
    ", INDCONSULTA = " & Formata_Dados(strConsulta, tpDados_Texto) & _
    " WHERE PKID = " & lngPROCEDIMENTOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProcedimento.AlterarProcedimento]", _
            Err.Description
End Sub

Public Function ListarProcedimento(lngPROCEDIMENTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PROCEDIMENTO.*, TIPOPROCEDIMENTO.TIPOPROCEDIMENTO " & _
    "FROM PROCEDIMENTO LEFT JOIN TIPOPROCEDIMENTO ON TIPOPROCEDIMENTO.PKID = PROCEDIMENTO.TIPOPROCEDIMENTOID " & _
    "WHERE PROCEDIMENTO.PKID = " & _
    Formata_Dados(lngPROCEDIMENTOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarProcedimento = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProcedimento.ListarProcedimento]", _
            Err.Description
End Function

Public Sub ExlcuirProcedimento(lngPROCEDIMENTOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM PROCEDIMENTO WHERE PKID = " & _
    Formata_Dados(lngPROCEDIMENTOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProcedimento.ExlcuirProcedimento]", _
            Err.Description
End Sub
