VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrigem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirOrigem(strNome As String, _
                         strIndicador As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO ORIGEM (NOME, IC_ORIGEM) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strIndicador, tpDados_Texto) & " " & _
    ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsOrigem.InserirOrigem]", _
            Err.Description
End Sub


Public Sub AlterarOrigem(lngORIGEMID As Long, _
                         strNome As String, _
                         strIndicador As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE ORIGEM SET NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    " IC_ORIGEM = " & Formata_Dados(strIndicador, tpDados_Texto) & " " & _
    "WHERE PKID = " & lngORIGEMID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsOrigem.AlterarOrigem]", _
            Err.Description
End Sub

Public Function ListarOrigem(lngORIGEMID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ORIGEM.*, CASE IC_ORIGEM WHEN 'O' THEN 'ORIGEM' WHEN 'D' THEN 'DESTINO' ELSE 'AMBOS' END AS IC_ORIGEM_DEST " & _
    "FROM ORIGEM " & _
    "WHERE ORIGEM.PKID = " & _
    Formata_Dados(lngORIGEMID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarOrigem = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsOrigem.ListarOrigem]", _
            Err.Description
End Function

Public Sub ExcluirOrigem(lngORIGEMID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui ORIGEM
  strSql = "DELETE FROM ORIGEM WHERE ORIGEM.PKID = " & _
    Formata_Dados(lngORIGEMID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsOrigem.ExcluirOrigem]", _
            Err.Description
End Sub
