VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCartaoDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirCartaoDebito(strNome As String, _
                               strNumero As String, _
                               strPercTxAdm As String, _
                               strConvenio As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO CARTAODEBITO (NOME, NUMERO, PERCTAXAADMIN, INDCONVENIO) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strNumero, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strPercTxAdm, tpDados_Moeda, tpNulo_Aceita) & ", " & _
    Formata_Dados(strConvenio, tpDados_Texto) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoDebito.InserirCartaoDebito]", _
            Err.Description
End Sub

Public Sub AlterarCartaoDebito(lngCARTAODEBITOID As Long, _
                               strNome As String, _
                               strNumero As String, _
                               strPercTxAdm As String, _
                               strConvenio As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE CARTAODEBITO SET NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "NUMERO = " & Formata_Dados(strNumero, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "PERCTAXAADMIN = " & Formata_Dados(strPercTxAdm, tpDados_Moeda, tpNulo_Aceita) & ", " & _
    "INDCONVENIO = " & Formata_Dados(strConvenio, tpDados_Texto) & " " & _
    "WHERE PKID = " & lngCARTAODEBITOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoDebito.AlterarCartaoDebito]", _
            Err.Description
End Sub

Public Function ListarCartaoDebito(lngCARTAODEBITOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CARTAODEBITO.* " & _
    "FROM CARTAODEBITO " & _
    "WHERE CARTAODEBITO.PKID = " & _
    Formata_Dados(lngCARTAODEBITOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarCartaoDebito = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoDebito.ListarCartaoDebito]", _
            Err.Description
End Function

Public Sub ExlcuirCartaoDebito(lngCARTAODEBITOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM CARTAODEBITO WHERE PKID = " & _
    Formata_Dados(lngCARTAODEBITOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoDebito.ExlcuirCartaoDebito]", _
            Err.Description
End Sub

