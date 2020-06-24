VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQualificacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirQualificacao(strDescricao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO QUALIFICACAO (DESCRICAO) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsQualificacao.InserirQualificacao]", _
            Err.Description
End Sub

Public Sub AlterarQualificacao(lngQUALIFICACAOID As Long, _
                               strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE QUALIFICACAO SET DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngQUALIFICACAOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsQualificacao.AlterarQualificacao]", _
            Err.Description
End Sub

Public Function ListarQualificacao(lngQUALIFICACAOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT QUALIFICACAO.* " & _
    "FROM QUALIFICACAO " & _
    "WHERE QUALIFICACAO.PKID = " & _
    Formata_Dados(lngQUALIFICACAOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarQualificacao = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsQualificacao.ListarQualificacao]", _
            Err.Description
End Function

Public Sub ExlcuirQualificacao(lngQUALIFICACAOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM QUALIFICACAO WHERE PKID = " & _
    Formata_Dados(lngQUALIFICACAOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsQualificacao.ExlcuirQualificacao]", _
            Err.Description
End Sub


Public Function ListarQualificacaoPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT QUALIFICACAO.* " & _
    "FROM QUALIFICACAO " & _
    "WHERE QUALIFICACAO.DESCRICAO = " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarQualificacaoPelaDesc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsQualificacao.ListarQualificacaoPelaDesc]", _
            Err.Description
End Function

