VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirModelo(strNome As String, _
                         lngMARCAID As Long)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO MODELO (NOME, MARCAID) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & _
    ", " & Formata_Dados(IIf(lngMARCAID = 0, "", lngMARCAID), tpDados_Longo) & _
    ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsModelo.InserirModelo]", _
            Err.Description
End Sub

Public Sub AlterarModelo(lngMODELOID As Long, _
                         strNome As String, _
                         lngMARCAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE MODELO SET NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    " MARCAID = " & Formata_Dados(lngMARCAID, tpDados_Longo) & " " & _
    " WHERE PKID = " & lngMODELOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsModelo.AlterarModelo]", _
            Err.Description
End Sub

Public Function ListarModelo(lngMODELOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT MODELO.*, MARCA.NOME AS NOME_MARCA " & _
    "FROM MODELO " & _
    " LEFT JOIN MARCA ON MARCA.PKID = MODELO.MARCAID " & _
    "WHERE MODELO.PKID = " & _
    Formata_Dados(lngMODELOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarModelo = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsModelo.ListarModelo]", _
            Err.Description
End Function

Public Sub ExcluirModelo(lngMODELOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui FAMÍLIA DE PRODUTOS
  strSql = "DELETE FROM MODELO WHERE MODELO.PKID = " & _
    Formata_Dados(lngMODELOID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsModelo.ExcluirModelo]", _
            Err.Description
End Sub


