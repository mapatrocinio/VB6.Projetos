VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit
Public Sub InserirTipoVenda(strDescricao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO TIPOVENDA (DESCRICAO) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoVenda.InserirTipoVenda]", _
            Err.Description
End Sub

Public Sub AlterarTipoVenda(lngTIPOVENDAID As Long, _
                                strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE TIPOVENDA SET DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngTIPOVENDAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoVenda.AlterarTipoVenda]", _
            Err.Description
End Sub

Public Function ListarTipoVenda(lngTIPOVENDAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPOVENDA.* " & _
    "FROM TIPOVENDA " & _
    "WHERE TIPOVENDA.PKID = " & _
    Formata_Dados(lngTIPOVENDAID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarTipoVenda = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoVenda.ListarTipoVenda]", _
            Err.Description
End Function

Public Sub ExlcuirTipoVenda(lngTIPOVENDAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM TIPOVENDA WHERE PKID = " & _
    Formata_Dados(lngTIPOVENDAID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoVenda.ExlcuirTipoVenda]", _
            Err.Description
End Sub



