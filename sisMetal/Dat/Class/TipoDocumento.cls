VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit
Public Sub InserirTipoDocumento(strDescricao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO TIPODOCUMENTO (DESCRICAO) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoDocumento.InserirTipoDocumento]", _
            Err.Description
End Sub

Public Sub AlterarTipoDocumento(lngTIPODOCUMENTOID As Long, _
                                strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE TIPODOCUMENTO SET DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngTIPODOCUMENTOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoDocumento.AlterarTipoDocumento]", _
            Err.Description
End Sub

Public Function ListarTipoDocumento(lngTIPODOCUMENTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPODOCUMENTO.* " & _
    "FROM TIPODOCUMENTO " & _
    "WHERE TIPODOCUMENTO.PKID = " & _
    Formata_Dados(lngTIPODOCUMENTOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarTipoDocumento = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoDocumento.ListarTipoDocumento]", _
            Err.Description
End Function

Public Sub ExlcuirTipoDocumento(lngTIPODOCUMENTOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM TIPODOCUMENTO WHERE PKID = " & _
    Formata_Dados(lngTIPODOCUMENTOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoDocumento.ExlcuirTipoDocumento]", _
            Err.Description
End Sub



