VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReceita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarReceita(lngRECEITAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT LIVRO.NUMEROLIVRO, FORMAPGTO.FORMAPGTO AS DESCRFORMAPGTO, GRUPODESPESA.CODIGO AS CODIGOGRUPODESPESA, SUBGRUPODESPESA.CODIGO AS CODIGOSUBGRUPODESPESA, EMPRESACTRTO.DESCRICAO AS DESCR_CONTRATO, RECEITA.* FROM RECEITA " & _
    " LEFT JOIN SUBGRUPODESPESA ON SUBGRUPODESPESA.PKID = RECEITA.SUBGRUPODESPESAID " & _
    " LEFT JOIN GRUPODESPESA ON GRUPODESPESA.PKID = SUBGRUPODESPESA.GRUPODESPESAID " & _
    " LEFT JOIN FORMAPGTO ON FORMAPGTO.PKID = RECEITA.FORMAPGTOID " & _
    " LEFT JOIN LIVRO ON LIVRO.PKID = RECEITA.LIVROID " & _
    " LEFT JOIN EMPRESACTRTO ON EMPRESACTRTO.PKID = RECEITA.EMPRESACTRTOID " & _
    " WHERE RECEITA.PKID = " & _
    Formata_Dados(lngRECEITAID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarReceita = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsReceita.SelecionarReceita]"
End Function
Public Sub ExcluirReceita(lngRECEITAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM RECEITA " & _
    " WHERE PKID = " & lngRECEITAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsReceita.ExcluirReceita]"
End Sub

Public Sub AlterarReceita(strDtRecebimento As String, _
                          strDtVencimento As String, _
                          strValorReceber As String, _
                          lngLivroId As Long, _
                          strNumeroCheque As String, _
                          lngRECEITAID As Long, _
                          strDescricao As String, _
                          strValorRecebido As String, _
                          lngSubGrupoDespesaId As Long, _
                          lngFormaPgtoId As Long, _
                          strUsuarioAlterouDescricao As String, _
                          lngEMPRESACTRTOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE RECEITA SET " & _
    " DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & _
    ", VR_RECEBIDO = " & Formata_Dados(strValorRecebido, tpDados_Moeda, tpNulo_Aceita) & _
    ", SUBGRUPODESPESAID = " & Formata_Dados(lngSubGrupoDespesaId, tpDados_Longo, tpNulo_Aceita) & _
    ", FORMAPGTOID = " & Formata_Dados(IIf(lngFormaPgtoId = 0, "", lngFormaPgtoId), tpDados_Longo, tpNulo_Aceita) & _
    ", DT_VENCIMENTO = " & Formata_Dados(strDtVencimento, tpDados_DataHora, tpNulo_Aceita) & _
    ", DT_RECEBIMENTO = " & Formata_Dados(strDtRecebimento, tpDados_DataHora, tpNulo_Aceita) & _
    ", VR_RECEBER = " & Formata_Dados(strValorReceber, tpDados_Moeda, tpNulo_Aceita) & _
    ", LIVROID = " & Formata_Dados(IIf(lngLivroId = 0, "", lngLivroId), tpDados_Longo, tpNulo_Aceita) & _
    ", NUMEROCHEQUE = " & Formata_Dados(strNumeroCheque, tpDados_Texto, tpNulo_Aceita) & _
    ", USUARIOALTDESCR = " & Formata_Dados(strUsuarioAlterouDescricao, tpDados_Texto, tpNulo_Aceita) & _
    ", EMPRESACTRTOID = " & Formata_Dados(lngEMPRESACTRTOID, tpDados_Longo) & _
    " WHERE PKID = " & lngRECEITAID
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsReceita.AlterarReceita]"
End Sub
Public Sub IncluirReceita(strDtVencimento As String, _
                          strValorReceber As String, _
                          lngLivroId As Long, _
                          strNumeroCheque As String, _
                          strTipo As String, _
                          strDtRecebimento As String, _
                          strDescricao As String, _
                          strValorRecebido As String, _
                          lngSubGrupoDespesaId As Long, _
                          lngFormaPgtoId As Long, _
                          strUsuario As String, _
                          ByRef lngSEQUENCIALEXTERNO As Long, _
                          lngEMPRESACTRTOID As Long, _
                          lngParceiroId As Long)
  On Error GoTo trata
  '
  Dim strSql          As String
  'Dim lngSequencial   As Long
  '
  'lngSequencial = RetonaProximoSequencial("RECEITA", "SEQUENCIAL")
  '
  strSql = "INSERT INTO RECEITA (" & _
    " DESCRICAO, VR_RECEBIDO, SUBGRUPODESPESAID, FORMAPGTOID" & _
    ", DT_RECEBIMENTO, SEQUENCIAL, TIPO, DT_VENCIMENTO, VR_RECEBER, LIVROID, NUMEROCHEQUE, USUARIO, EMPRESACTRTOID, PARCEIROID) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strValorRecebido, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(lngSubGrupoDespesaId, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(IIf(lngFormaPgtoId = 0, "", lngFormaPgtoId), tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strDtRecebimento, tpDados_DataHora, tpNulo_Aceita) & _
    ", " & Formata_Dados(lngSEQUENCIALEXTERNO, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strTipo, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strDtVencimento, tpDados_DataHora, tpNulo_Aceita) & _
    ", " & Formata_Dados(strValorReceber, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(IIf(lngLivroId = 0, "", lngLivroId), tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strNumeroCheque, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strUsuario, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(lngEMPRESACTRTOID, tpDados_Longo) & _
    ", " & Formata_Dados(lngParceiroId, tpDados_Longo) & _
    ")"

  RunSP strSql
  'lngSEQUENCIALEXTERNO = lngSequencial
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsReceita.IncluirReceita]"
End Sub

