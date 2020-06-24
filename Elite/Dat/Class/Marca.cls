VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMarca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirMarca(strNome As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO MARCA (NOME) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsMarca.InserirMarca]", _
            Err.Description
End Sub


Public Sub AlterarMarca(lngMARCAID As Long, _
                        strNome As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE MARCA SET NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngMARCAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsMarca.AlterarMarca]", _
            Err.Description
End Sub

Public Function ListarMarca(lngMARCAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT MARCA.* " & _
    "FROM MARCA " & _
    "WHERE MARCA.PKID = " & _
    Formata_Dados(lngMARCAID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarMarca = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsMarca.ListarMarca]", _
            Err.Description
End Function

Public Sub ExcluirMarca(lngMARCAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui DOCUMENTO DE SAIDA
  strSql = "DELETE FROM MARCA WHERE MARCA.PKID = " & _
    Formata_Dados(lngMARCAID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsMarca.ExcluirMarca]", _
            Err.Description
End Sub
