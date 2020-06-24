VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoAjuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarTipoAjuste(lngTIPO_AJUSTEID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPO_AJUSTE.*, " & _
    " Case TIPO_AJUSTE.TIPO " & _
    " WHEN 'PA' THEN 'Peça (Adicionar)' " & _
    " WHEN 'PR' THEN 'Peça (Retirar)' " & _
    " WHEN 'TO' THEN 'Total' " & _
    " Else '' END AS TIPO_AJUSTE " & _
    " FROM TIPO_AJUSTE " & _
    " WHERE TIPO_AJUSTE.PKID = " & _
    Formata_Dados(lngTIPO_AJUSTEID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarTipoAjuste = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsTipoAjuste.SelecionarTipoAjuste]"
End Function
Public Sub ExcluirTipoAjuste(lngTIPO_AJUSTEID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM TIPO_AJUSTE " & _
    " WHERE PKID = " & lngTIPO_AJUSTEID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsTipoAjuste.ExcluirTipoAjuste]"
End Sub

Public Sub AlterarTipoAjuste(strTipo As String, _
                      strDescricao As String, _
                      lngTIPO_AJUSTEID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE TIPO_AJUSTE SET " & _
    " DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto) & _
    ", TIPO = " & Formata_Dados(strTipo, tpDados_Texto, tpNulo_Aceita) & _
    " WHERE PKID = " & lngTIPO_AJUSTEID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsTipoAjuste.AlterarTipoAjuste]"
End Sub
Public Sub IncluirTipoAjuste(strTipo As String, _
                      strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql          As String
  '
  strSql = "INSERT INTO TIPO_AJUSTE (" & _
    " DESCRICAO, TIPO) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strTipo, tpDados_Texto, tpNulo_Aceita) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsTipoAjuste.IncluirTipoAjuste]"
End Sub

