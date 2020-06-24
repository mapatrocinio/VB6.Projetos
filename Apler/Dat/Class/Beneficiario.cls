VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBeneficiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirBeneficiario(lngBeneficiarioId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM BENEFICIARIO " & _
    "WHERE BENEFICIARIO.PKID = " & _
    Formata_Dados(lngBeneficiarioId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBeneficiario.ExcluirBeneficiario]", _
            Err.Description
End Function

Public Function SelecionarBeneficiarioPeloPkid(lngBeneficiarioId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT BENEFICIARIO.*, GRAUPARENTESCO.DESCRICAO AS DESCR_GRAUPARENTESCO " & _
    "FROM BENEFICIARIO LEFT JOIN GRAUPARENTESCO ON GRAUPARENTESCO.PKID = BENEFICIARIO.GRAUPARENTESCOID " & _
    "WHERE BENEFICIARIO.PKID = " & _
    Formata_Dados(lngBeneficiarioId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarBeneficiarioPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBeneficiario.SelecionarBeneficiarioPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarBeneficiario(lngPKID As Long, _
                               lngGRAUPARENTESCOID As Long, _
                               strNome As String, _
                               strCpf As String, _
                               strIdentidade As String, _
                               strDataNascimento As String, _
                               strSexo As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE BENEFICIARIO SET " & _
    "GRAUPARENTESCOID = " & Formata_Dados(IIf(lngGRAUPARENTESCOID = 0, "", lngGRAUPARENTESCOID), tpDados_Longo) & ", " & _
    "NOME = " & Formata_Dados(strNome, tpDados_Texto) & ", " & _
    "CPF = " & Formata_Dados(strCpf, tpDados_Texto) & ", " & _
    "IDENTIDADE = " & Formata_Dados(strIdentidade, tpDados_Texto) & ", " & _
    "DATANASCIMENTO = " & Formata_Dados(strDataNascimento, tpDados_DataHora) & ", " & _
    "SEXO = " & Formata_Dados(strSexo, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBeneficiario.AlterarBeneficiario]", _
            Err.Description
End Sub

Public Sub InserirBeneficiario(lngTABCONVASSOCID As Long, _
                               lngGRAUPARENTESCOID As Long, _
                               strNome As String, _
                               strCpf As String, _
                               strIdentidade As String, _
                               strDataNascimento As String, _
                               strSexo As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO BENEFICIARIO (TABCONVASSOCID, GRAUPARENTESCOID, NOME, CPF, IDENTIDADE, DATANASCIMENTO, SEXO) VALUES (" & _
    Formata_Dados(lngTABCONVASSOCID, tpDados_Longo) & ", " & _
    Formata_Dados(IIf(lngGRAUPARENTESCOID = 0, "", lngGRAUPARENTESCOID), tpDados_Longo) & ", " & _
    Formata_Dados(strNome, tpDados_Texto) & ", " & _
    Formata_Dados(strCpf, tpDados_Texto) & ", " & _
    Formata_Dados(strIdentidade, tpDados_Texto) & ", " & _
    Formata_Dados(strDataNascimento, tpDados_DataHora) & ", " & _
    Formata_Dados(strSexo, tpDados_Texto) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBeneficiario.InserirBeneficiario]", _
            Err.Description
End Sub

Public Function SelecionarBeneficiario(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT BENEFICIARIO.* " & _
    "FROM BENEFICIARIO " & _
    "WHERE BENEFICIARIO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarBeneficiario = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBeneficiario.SelecionarBeneficiario]", _
            Err.Description
End Function

