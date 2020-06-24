VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit
Public Sub InserirBanco(strNome As String, _
                        strNumero As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO BANCO (NOME, NUMERO) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strNumero, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBanco.InserirBanco]", _
            Err.Description
End Sub

Public Sub AlterarBanco(lngBANCOID As Long, _
                        strNome As String, _
                        strNumero As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE BANCO SET NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "NUMERO = " & Formata_Dados(strNumero, tpDados_Texto, tpNulo_Aceita) & " " & _
    "WHERE PKID = " & lngBANCOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBanco.AlterarBanco]", _
            Err.Description
End Sub

Public Function ListarBanco(lngBANCOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT BANCO.* " & _
    "FROM BANCO " & _
    "WHERE BANCO.PKID = " & _
    Formata_Dados(lngBANCOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarBanco = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBanco.ListarBanco]", _
            Err.Description
End Function

Public Sub ExlcuirBanco(lngBANCOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM BANCO WHERE PKID = " & _
    Formata_Dados(lngBANCOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBanco.ExlcuirBanco]", _
            Err.Description
End Sub



