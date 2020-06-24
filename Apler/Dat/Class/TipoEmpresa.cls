VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit
Public Sub InserirTipoEmpresa(strDescricao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO TIPOEMPRESA (DESCRICAO) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoEmpresa.InserirTipoEmpresa]", _
            Err.Description
End Sub

Public Sub AlterarTipoEmpresa(lngTIPOEMPRESAID As Long, _
                              strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE TIPOEMPRESA SET DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngTIPOEMPRESAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoEmpresa.AlterarTipoEmpresa]", _
            Err.Description
End Sub

Public Function ListarTipoEmpresa(lngTIPOEMPRESAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPOEMPRESA.* " & _
    "FROM TIPOEMPRESA " & _
    "WHERE TIPOEMPRESA.PKID = " & _
    Formata_Dados(lngTIPOEMPRESAID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarTipoEmpresa = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoEmpresa.ListarTipoEmpresa]", _
            Err.Description
End Function

Public Sub ExlcuirTipoEmpresa(lngTIPOEMPRESAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM TIPOEMPRESA WHERE PKID = " & _
    Formata_Dados(lngTIPOEMPRESAID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoEmpresa.ExlcuirTipoEmpresa]", _
            Err.Description
End Sub



