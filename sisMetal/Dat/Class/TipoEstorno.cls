VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoEstorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit
Public Sub InserirTipoEstorno(strDescricao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO TIPOESTORNO (DESCRICAO) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoEstorno.InserirTipoEstorno]", _
            Err.Description
End Sub

Public Sub AlterarTipoEstorno(lngTIPOESTORNOID As Long, _
                                strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE TIPOESTORNO SET DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngTIPOESTORNOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoEstorno.AlterarTipoEstorno]", _
            Err.Description
End Sub

Public Function ListarTipoEstorno(lngTIPOESTORNOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPOESTORNO.* " & _
    "FROM TIPOESTORNO " & _
    "WHERE TIPOESTORNO.PKID = " & _
    Formata_Dados(lngTIPOESTORNOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarTipoEstorno = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoEstorno.ListarTipoEstorno]", _
            Err.Description
End Function

Public Sub ExlcuirTipoEstorno(lngTIPOESTORNOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM TIPOESTORNO WHERE PKID = " & _
    Formata_Dados(lngTIPOESTORNOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoEstorno.ExlcuirTipoEstorno]", _
            Err.Description
End Sub



