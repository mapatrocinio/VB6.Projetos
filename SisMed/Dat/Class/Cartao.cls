VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCartao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirCartao(strNome As String, _
                         strNumero As String, _
                         strPercTxAdm As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO CARTAO (NOME, NUMERO, PERCTAXAADMIN) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strNumero, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strPercTxAdm, tpDados_Moeda, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartao.InserirCartao]", _
            Err.Description
End Sub

Public Sub AlterarCartao(lngCARTAOID As Long, _
                         strNome As String, _
                         strNumero As String, _
                         strPercTxAdm As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE CARTAO SET NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "NUMERO = " & Formata_Dados(strNumero, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "PERCTAXAADMIN = " & Formata_Dados(strPercTxAdm, tpDados_Moeda, tpNulo_Aceita) & " " & _
    "WHERE PKID = " & lngCARTAOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartao.AlterarCartao]", _
            Err.Description
End Sub

Public Function ListarCartao(lngCARTAOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CARTAO.* " & _
    "FROM CARTAO " & _
    "WHERE CARTAO.PKID = " & _
    Formata_Dados(lngCARTAOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarCartao = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartao.ListarCartao]", _
            Err.Description
End Function

Public Sub ExlcuirCartao(lngCARTAOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM CARTAO WHERE PKID = " & _
    Formata_Dados(lngCARTAOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartao.ExlcuirCartao]", _
            Err.Description
End Sub

