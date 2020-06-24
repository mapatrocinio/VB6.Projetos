VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarPacientePeloPkid(lngPRONTUARIOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PACIENTE.* " & _
    "FROM PACIENTE " & _
    "WHERE PACIENTE.PRONTUARIOID = " & _
    Formata_Dados(lngPRONTUARIOID, tpDados_Longo)
  Set SelecionarPacientePeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPaciente.SelecionarPacientePeloPkid]", _
            Err.Description
End Function

Public Sub AlterarPaciente(lngPRONTUARIOID As Long)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE PACIENTE SET " & _
    "DTCADASTRO = " & Formata_Dados(Format(Now, "DD/MM/YYYY hh:mm"), tpDados_DataHora) & " " & _
    "WHERE PRONTUARIOID = " & Formata_Dados(lngPRONTUARIOID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPaciente.AlterarPaciente]", _
            Err.Description
End Sub

Public Sub InserirPaciente(lngPRONTUARIOID As Long)
  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO PACIENTE (PRONTUARIOID, DTCADASTRO) VALUES (" & _
    " " & Formata_Dados(lngPRONTUARIOID, tpDados_Longo) & _
    ", " & Formata_Dados(Format(Now, "DD/MM/YYYY hh:mm"), tpDados_DataHora) & ")"
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPaciente.InserirPaciente]", _
            Err.Description
End Sub

