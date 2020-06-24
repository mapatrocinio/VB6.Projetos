VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubGrupoDespesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarTipoGrupo(strCodigoGrupo As String, _
                                    strCodigoSubGrupo As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT GRUPODESPESA.TIPO " & _
    " FROM GRUPODESPESA INNER JOIN SUBGRUPODESPESA ON GRUPODESPESA.PKID = SUBGRUPODESPESA.GRUPODESPESAID " & _
    " WHERE GRUPODESPESA.CODIGO = " & _
    Formata_Dados(strCodigoGrupo, tpDados_Texto, tpNulo_NaoAceita) & _
    " AND SUBGRUPODESPESA.CODIGO = " & _
    Formata_Dados(strCodigoSubGrupo, tpDados_Texto, tpNulo_NaoAceita)
  Set SelecionarTipoGrupo = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsSubGrupoDespesa.SelecionarTipoGrupo]"
End Function
Public Function SelecionarSubGrupoDespesa(lngSUBGRUPODESPESAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT SUBGRUPODESPESA.* " & _
    " FROM SUBGRUPODESPESA " & _
    " WHERE SUBGRUPODESPESA.PKID = " & _
    Formata_Dados(lngSUBGRUPODESPESAID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarSubGrupoDespesa = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsSubGrupoDespesa.SelecionarSubGrupoDespesa]"
End Function
Public Sub ExcluirSubGrupoDespesa(lngSUBGRUPODESPESAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM SUBGRUPODESPESA " & _
    " WHERE PKID = " & lngSUBGRUPODESPESAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsSubGrupoDespesa.ExcluirSubGrupoDespesa]"
End Sub

Public Sub AlterarSubGrupoDespesa(lngSUBGRUPODESPESAID As Long, _
                                  strCodigo As String, _
                                  strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE SUBGRUPODESPESA SET " & _
    " CODIGO = " & Formata_Dados(strCodigo, tpDados_Texto, tpNulo_Aceita) & _
    ", DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & _
    " WHERE PKID = " & lngSUBGRUPODESPESAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsSubGrupoDespesa.AlterarSubGrupoDespesa]"
End Sub
Public Sub IncluirSubGrupoDespesa(lngGrupoDespesaId As Long, _
                                  strCodigo As String, _
                                  strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql          As String
  '
  strSql = "INSERT INTO SUBGRUPODESPESA (" & _
    " CODIGO, DESCRICAO, GRUPODESPESAID) VALUES (" & _
    Formata_Dados(strCodigo, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(lngGrupoDespesaId, tpDados_Longo, tpNulo_Aceita) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsSubGrupoDespesa.IncluirSubGrupoDespesa]"
End Sub

