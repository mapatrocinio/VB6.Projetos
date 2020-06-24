VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTabFichaClieLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarTabFichaClieLoc(lngLocacaoId As Long, _
                                          strTipoCliente As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TAB_FICHACLIELOC.* " & _
    "FROM TAB_FICHACLIELOC " & _
    "WHERE TAB_FICHACLIELOC.LOCACAOID = " & _
    Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_NaoAceita) & _
    " AND TIPOCLIENTE = " & _
    Formata_Dados(strTipoCliente, tpDados_Texto, tpNulo_NaoAceita)
  Set SelecionarTabFichaClieLoc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTabFichaClieLoc.SelecionarTabFichaClieLoc]", _
            Err.Description
End Function

Public Sub ExcluirTabFichaClieLoc(lngTabFichaClieLocId As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  '
  strSql = "DELETE FROM TAB_FICHACLIELOC " & _
    " WHERE TAB_FICHACLIELOC.PKID = " & Formata_Dados(lngTabFichaClieLocId, tpDados_Longo)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTabFichaClieLoc.ExcluirTabFichaClieLoc]", _
            Err.Description
End Sub

Public Sub AlterarTabFichaClieLoc(lngTabFichaClieLocId As Long, _
                                  lngFichaClienteId As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  '
  strSql = "UPDATE TAB_FICHACLIELOC SET " & _
    " FICHACLIENTEID = " & Formata_Dados(lngFichaClienteId, tpDados_Longo, tpNulo_Aceita) & _
    " WHERE TAB_FICHACLIELOC.PKID = " & Formata_Dados(lngTabFichaClieLocId, tpDados_Longo)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTabFichaClieLoc.AlterarTabFichaClieLoc]", _
            Err.Description
End Sub

Public Sub IncluirTabFichaClieLoc(lngLocacoId As Long, _
                                  lngFichaClienteId As Long, _
                                  strTipoCliente As String)
  On Error GoTo trata
  '
  Dim strSql As String
  '
  strSql = "INSERT INTO TAB_FICHACLIELOC (" & _
    "LOCACAOID, FICHACLIENTEID, TIPOCLIENTE) VALUES (" & _
    Formata_Dados(lngLocacoId, tpDados_Longo, tpNulo_Aceita) & _
    ",  " & Formata_Dados(lngFichaClienteId, tpDados_Longo, tpNulo_Aceita) & _
    ",  " & Formata_Dados(strTipoCliente, tpDados_Texto, tpNulo_Aceita) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTabFichaClieLoc.IncluirTabFichaClieLoc]", _
            Err.Description
End Sub




