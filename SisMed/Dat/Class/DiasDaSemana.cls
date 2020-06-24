VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiasDaSemana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirDiaDaSemana(strCodigo As String, _
                              strDiaDaSemana As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO DIASDASEMANA (CODIGO, DIADASEMANA) VALUES (" & _
    Formata_Dados(strCodigo, tpDados_Longo, tpNulo_Aceita) & ", " & _
    Formata_Dados(strDiaDaSemana, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDiasDaSemana.InserirDiaDaSemana]", _
            Err.Description
End Sub

Public Sub AlterarDiaDaSemana(lngDIADASEMANAID As Long, _
                              strCodigo As String, _
                              strDiaDaSemana As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE DIASDASEMANA SET CODIGO = " & Formata_Dados(strCodigo, tpDados_Longo, tpNulo_Aceita) & ", " & _
    "DIADASEMANA = " & Formata_Dados(strDiaDaSemana, tpDados_Texto, tpNulo_Aceita) & " " & _
    "WHERE PKID = " & lngDIADASEMANAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDiasDaSemana.AlterarDiaDaSemana]", _
            Err.Description
End Sub

Public Function ListarDiaDaSemana(lngDIADASEMANAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT DIASDASEMANA.* " & _
    "FROM DIASDASEMANA " & _
    "WHERE DIASDASEMANA.PKID = " & _
    Formata_Dados(lngDIADASEMANAID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarDiaDaSemana = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDiasDaSemana.ListarDiaDaSemana]", _
            Err.Description
End Function

Public Sub ExlcuirDiaDaSemana(lngDIADASEMANAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM DIASDASEMANA WHERE PKID = " & _
    Formata_Dados(lngDIADASEMANAID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDiasDaSemana.ExlcuirDiaDaSemana]", _
            Err.Description
End Sub

