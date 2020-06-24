VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit
Public Sub InserirTipoPagamento(strDescricao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO TIPOPAGAMENTO (DESCRICAO) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoPagamento.InserirTipoPagamento]", _
            Err.Description
End Sub

Public Sub AlterarTipoPagamento(lngTIPOPAGAMENTOID As Long, _
                                strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE TIPOPAGAMENTO SET DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngTIPOPAGAMENTOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoPagamento.AlterarTipoPagamento]", _
            Err.Description
End Sub

Public Function ListarTipoPagamento(lngTIPOPAGAMENTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPOPAGAMENTO.* " & _
    "FROM TIPOPAGAMENTO " & _
    "WHERE TIPOPAGAMENTO.PKID = " & _
    Formata_Dados(lngTIPOPAGAMENTOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarTipoPagamento = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoPagamento.ListarTipoPagamento]", _
            Err.Description
End Function

Public Sub ExlcuirTipoPagamento(lngTIPOPAGAMENTOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM TIPOPAGAMENTO WHERE PKID = " & _
    Formata_Dados(lngTIPOPAGAMENTOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoPagamento.ExlcuirTipoPagamento]", _
            Err.Description
End Sub



