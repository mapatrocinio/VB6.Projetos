VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirTipoConvenio(strDescricao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO TIPOCONVENIO (DESCRICAO) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoConvenio.InserirTipoConvenio]", _
            Err.Description
End Sub

Public Sub AlterarTipoConvenio(lngTIPOCONVENIOID As Long, _
                               strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE TIPOCONVENIO SET DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngTIPOCONVENIOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoConvenio.AlterarTipoConvenio]", _
            Err.Description
End Sub

Public Function ListarTipoConvenio(lngTIPOCONVENIOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPOCONVENIO.* " & _
    "FROM TIPOCONVENIO " & _
    "WHERE TIPOCONVENIO.PKID = " & _
    Formata_Dados(lngTIPOCONVENIOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarTipoConvenio = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoConvenio.ListarTipoConvenio]", _
            Err.Description
End Function

Public Sub ExlcuirTipoConvenio(lngTIPOCONVENIOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM TIPOCONVENIO WHERE PKID = " & _
    Formata_Dados(lngTIPOCONVENIOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoConvenio.ExlcuirTipoConvenio]", _
            Err.Description
End Sub


Public Function ListarTipoConvenioPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPOCONVENIO.* " & _
    "FROM TIPOCONVENIO " & _
    "WHERE TIPOCONVENIO.DESCRICAO = " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarTipoConvenioPelaDesc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoConvenio.ListarTipoConvenioPelaDesc]", _
            Err.Description
End Function

