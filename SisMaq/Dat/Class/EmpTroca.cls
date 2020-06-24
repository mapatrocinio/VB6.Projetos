VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpTroca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirEmpTroca(lngEMPTROCAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM EMPTROCA " & _
    "WHERE EMPTROCA.PKID = " & _
    Formata_Dados(lngEMPTROCAID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpTroca.ExcluirEmpTroca]", _
            Err.Description
End Function


Public Function SelecionarEmpTrocaPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT EMPTROCA.*, TIPOPGTO.TIPOPGTO " & _
    "FROM EMPTROCA " & _
    "INNER JOIN TIPOPGTO ON TIPOPGTO.PKID = EMPTROCA.TIPOPGTOID " & _
    "WHERE EMPTROCA.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarEmpTrocaPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpTroca.SelecionarEmpTrocaPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarEmpTroca(lngPKID As Long, _
                           lngTIPOPGTOID As Long, _
                           strValor As String, _
                           strNome As String)

  On Error GoTo trata
  '
  Dim strSql As String
  'Inserir EmpTroca
  strSql = "UPDATE EMPTROCA SET " & _
    "TIPOPGTOID = " & Formata_Dados(lngTIPOPGTOID, tpDados_Longo) & _
    ", VALOR = " & Formata_Dados(strValor, tpDados_Moeda) & _
    ", NOME = " & Formata_Dados(strNome, tpDados_Texto) & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpTroca.AlterarEmpTroca]", _
            Err.Description
End Sub

Public Sub InserirEmpTroca(lngTURNOID As Long, _
                           lngTIPOPGTOID As Long, _
                           strValor As String, _
                           strData As String, _
                           strNome As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = strSql & "INSERT INTO EMPTROCA (TURNOID, TIPOPGTOID, VALOR, DATA, NOME) VALUES (" & _
    Formata_Dados(lngTURNOID, tpDados_Longo) & ", " & _
    Formata_Dados(lngTIPOPGTOID, tpDados_Longo) & ", " & _
    Formata_Dados(strValor, tpDados_Moeda) & ", " & _
    Formata_Dados(strData, tpDados_DataHora) & ", " & _
    Formata_Dados(strNome, tpDados_Texto) & ");"
    
  '
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpTroca.InserirEmpTroca]", _
            Err.Description
End Sub





