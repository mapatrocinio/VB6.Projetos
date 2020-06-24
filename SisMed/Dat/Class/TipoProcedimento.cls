VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoProcedimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirTipoProcedimento(strTipoProcedimento As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO TIPOPROCEDIMENTO (TIPOPROCEDIMENTO) VALUES (" & _
    Formata_Dados(strTipoProcedimento, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoProcedimento.InserirTipoProcedimento]", _
            Err.Description
End Sub

Public Sub AlterarTipoProcedimento(lngTIPOPROCEDIMENTOID As Long, _
                         strTipoProcedimento As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE TIPOPROCEDIMENTO SET TIPOPROCEDIMENTO = " & Formata_Dados(strTipoProcedimento, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngTIPOPROCEDIMENTOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoProcedimento.AlterarTipoProcedimento]", _
            Err.Description
End Sub

Public Function ListarTipoProcedimento(lngTIPOPROCEDIMENTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPOPROCEDIMENTO.* " & _
    "FROM TIPOPROCEDIMENTO " & _
    "WHERE TIPOPROCEDIMENTO.PKID = " & _
    Formata_Dados(lngTIPOPROCEDIMENTOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarTipoProcedimento = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoProcedimento.ListarTipoProcedimento]", _
            Err.Description
End Function

Public Sub ExlcuirTipoProcedimento(lngTIPOPROCEDIMENTOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM TIPOPROCEDIMENTO WHERE PKID = " & _
    Formata_Dados(lngTIPOPROCEDIMENTOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoProcedimento.ExlcuirTipoProcedimento]", _
            Err.Description
End Sub
