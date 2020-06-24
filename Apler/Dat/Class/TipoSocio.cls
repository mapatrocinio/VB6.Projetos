VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirTipoSocio(strDescricao As String, _
                            strTipo As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO TIPOSOCIO (DESCRICAO, TIPO) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strTipo, tpDados_Texto) & ")"
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoSocio.InserirTipoSocio]", _
            Err.Description
End Sub

Public Sub AlterarTipoSocio(lngTIPOSOCIOID As Long, _
                            strDescricao As String, _
                            strTipo As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE TIPOSOCIO SET DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "TIPO = " & Formata_Dados(strTipo, tpDados_Texto) & " " & _
    "WHERE PKID = " & lngTIPOSOCIOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoSocio.AlterarTipoSocio]", _
            Err.Description
End Sub

Public Function ListarTipoSocio(lngTIPOSOCIOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPOSOCIO.* " & _
    "FROM TIPOSOCIO " & _
    "WHERE TIPOSOCIO.PKID = " & _
    Formata_Dados(lngTIPOSOCIOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarTipoSocio = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoSocio.ListarTipoSocio]", _
            Err.Description
End Function

Public Sub ExlcuirTipoSocio(lngTIPOSOCIOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM TIPOSOCIO WHERE PKID = " & _
    Formata_Dados(lngTIPOSOCIOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoSocio.ExlcuirTipoSocio]", _
            Err.Description
End Sub


Public Function ListarTipoSocioPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPOSOCIO.* " & _
    "FROM TIPOSOCIO " & _
    "WHERE TIPOSOCIO.DESCRICAO = " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarTipoSocioPelaDesc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoSocio.ListarTipoSocioPelaDesc]", _
            Err.Description
End Function

