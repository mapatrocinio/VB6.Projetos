VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirICMS(strDescricao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO ICMS (ICMS) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Moeda, tpNulo_NaoAceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsICMS.InserirICMS]", _
            Err.Description
End Sub

Public Sub AlterarICMS(lngICMSID As Long, _
                       strICMS As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE ICMS SET ICMS = " & Formata_Dados(strICMS, tpDados_Moeda) & " " & _
    "WHERE PKID = " & lngICMSID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsICMS.AlterarICMS]", _
            Err.Description
End Sub

Public Function ListarICMS(lngICMSID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ICMS.* " & _
    "FROM ICMS " & _
    "WHERE ICMS.PKID = " & _
    Formata_Dados(lngICMSID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarICMS = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsICMS.ListarICMS]", _
            Err.Description
End Function

Public Sub ExcluirICMS(lngICMSID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui FAMÍLIA DE PRODUTOS
  strSql = "DELETE FROM ICMS WHERE ICMS.PKID = " & _
    Formata_Dados(lngICMSID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsICMS.ExcluirICMS]", _
            Err.Description
End Sub


