VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLeituraFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirLeituraFecha(lngLEITURAFECHAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM LEITURAFECHA " & _
    "WHERE LEITURAFECHA.PKID = " & _
    Formata_Dados(lngLEITURAFECHAID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLeituraFecha.ExcluirLeituraFecha]", _
            Err.Description
End Function


Public Function SelecionarLeituraFechaPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT LEITURAFECHA.* " & _
    "FROM LEITURAFECHA " & _
    "WHERE LEITURAFECHA.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarLeituraFechaPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLeituraFecha.SelecionarLeituraFechaPeloPkid]", _
            Err.Description
End Function



Public Sub InserirLeituraFecha(lngLEITURAFECHAID As Long, _
                               lngCAIXAID As Long, _
                               lngPERIODOID As Long, _
                               strData As String, _
                               strStatus As String)

  On Error GoTo trata
  Dim strSql  As String
  Dim objRs   As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO LEITURAFECHA (CAIXAID, PERIODOID, DATA, STATUS) VALUES (" & _
    Formata_Dados(lngCAIXAID, tpDados_Longo) & _
    ", " & Formata_Dados(lngPERIODOID, tpDados_Longo) & _
    ", " & Formata_Dados(strData, tpDados_DataHora) & _
    ", " & Formata_Dados(strStatus, tpDados_Texto) & _
    ");"
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngLEITURAFECHAID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLeituraFecha.InserirLeituraFecha]", _
            Err.Description
End Sub





