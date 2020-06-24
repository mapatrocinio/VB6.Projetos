VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirEmpresa(strNome As String, _
                          strCodigo As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO EMPRESA (NOME, CODIGO) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strCodigo, tpDados_Texto) & ")"
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresa.InserirEmpresa]", _
            Err.Description
End Sub

Public Sub AlterarEmpresa(lngEMPRESAID As Long, _
                          strNome As String, _
                          strCodigo As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE EMPRESA SET NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "CODIGO = " & Formata_Dados(strCodigo, tpDados_Texto) & " " & _
    "WHERE PKID = " & lngEMPRESAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresa.AlterarEmpresa]", _
            Err.Description
End Sub

Public Function ListarEmpresa(lngEMPRESAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT EMPRESA.* " & _
    "FROM EMPRESA " & _
    "WHERE EMPRESA.PKID = " & _
    Formata_Dados(lngEMPRESAID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarEmpresa = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresa.ListarEmpresa]", _
            Err.Description
End Function

Public Sub ExlcuirEmpresa(lngEMPRESAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM EMPRESA WHERE PKID = " & _
    Formata_Dados(lngEMPRESAID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresa.ExlcuirEmpresa]", _
            Err.Description
End Sub


Public Function ListarEmpresaPeloNome(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT EMPRESA.* " & _
    "FROM EMPRESA " & _
    "WHERE EMPRESA.NOME = " & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarEmpresaPeloNome = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresa.ListarEmpresaPeloNome]", _
            Err.Description
End Function

