VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLeitura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirLeitura(lngLEITURAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM LEITURA " & _
    "WHERE LEITURA.PKID = " & _
    Formata_Dados(lngLEITURAID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLeitura.ExcluirLeitura]", _
            Err.Description
End Function


Public Function SelecionarLeituraPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT LEITURA.* " & _
    "FROM LEITURA " & _
    "WHERE LEITURA.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarLeituraPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLeitura.SelecionarLeituraPeloPkid]", _
            Err.Description
End Function



Public Sub InserirLeitura(lngLEITURAID As Long, _
                          lngLEITURISTAID As Long, _
                          strData As String)

  On Error GoTo trata
  Dim strSql  As String
  Dim objRs   As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO LEITURA (LEITURISTAID, DATA) VALUES (" & _
    Formata_Dados(lngLEITURISTAID, tpDados_Longo) & _
    ", " & Formata_Dados(strData, tpDados_DataHora) & _
    ");"
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngLEITURAID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLeitura.InserirLeitura]", _
            Err.Description
End Sub





