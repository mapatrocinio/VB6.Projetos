VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMotorista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarMotoristaPeloPkid(lngMOTORISTAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT MOTORISTA.*, BANCO.NOME AS NOME_BANCO " & _
    "FROM MOTORISTA " & _
    " LEFT JOIN BANCO ON BANCO.PKID = MOTORISTA.BANCOID " & _
    "WHERE MOTORISTA.PESSOAID = " & _
    Formata_Dados(lngMOTORISTAID, tpDados_Longo)
  Set SelecionarMotoristaPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsMotorista.SelecionarMotoristaPeloPkid]", _
            Err.Description
End Function

Public Sub AlterarMotorista(lngPESSOAID As Long, _
                            lngBANCOID As Long, _
                            strAgencia As String, _
                            strConta As String, _
                            strExcluido As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE MOTORISTA SET " & _
    "BANCOID = " & Formata_Dados(IIf(lngBANCOID = 0, "", lngBANCOID), tpDados_Longo) & ", " & _
    "AGENCIA = " & Formata_Dados(strAgencia, tpDados_Texto) & ", " & _
    "CONTA = " & Formata_Dados(strConta, tpDados_Texto) & ", " & _
    "INDEXCLUIDO = " & Formata_Dados(strExcluido, tpDados_Texto) & " " & _
    "WHERE PESSOAID = " & Formata_Dados(lngPESSOAID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsMotorista.AlterarMotorista]", _
            Err.Description
End Sub

Public Sub InserirMotorista(lngPESSOAID As Long, _
                            lngBANCOID As Long, _
                            strAgencia As String, _
                            strConta As String, _
                            strExcluido As String)
  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO MOTORISTA (PESSOAID, BANCOID, AGENCIA, CONTA, INDEXCLUIDO) VALUES (" & _
    " " & Formata_Dados(lngPESSOAID, tpDados_Longo) & _
    ", " & Formata_Dados(IIf(lngBANCOID = 0, "", lngBANCOID), tpDados_Longo) & _
    ", " & Formata_Dados(strConta, tpDados_Texto) & _
    ", " & Formata_Dados(strAgencia, tpDados_Texto) & _
    ", " & Formata_Dados(strExcluido, tpDados_Texto) & ")"
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsMotorista.InserirMotorista]", _
            Err.Description
End Sub

