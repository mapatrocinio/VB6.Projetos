VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDespesaCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarDespesa(lngDESPESAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT LIVRO.NUMEROLIVRO, FORMAPGTO.FORMAPGTO AS DESCRFORMAPGTO, GRUPODESPESA.CODIGO AS CODIGOGRUPODESPESA, SUBGRUPODESPESA.CODIGO AS CODIGOSUBGRUPODESPESA, DESPESA.* FROM DESPESA " & _
    " LEFT JOIN SUBGRUPODESPESA ON SUBGRUPODESPESA.PKID = DESPESA.SUBGRUPODESPESAID " & _
    " LEFT JOIN GRUPODESPESA ON GRUPODESPESA.PKID = SUBGRUPODESPESA.GRUPODESPESAID " & _
    " LEFT JOIN FORMAPGTO ON FORMAPGTO.PKID = DESPESA.FORMAPGTOID " & _
    " LEFT JOIN LIVRO ON LIVRO.PKID = DESPESA.LIVROID " & _
    " WHERE DESPESA.PKID = " & _
    Formata_Dados(lngDESPESAID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarDespesa = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsDespesa.SelecionarDespesa]"
End Function
Public Sub ExcluirDespesa(lngDESPESAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM DESPESA " & _
    " WHERE PKID = " & lngDESPESAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsDespesa.ExcluirDespesa]"
End Sub

Public Sub AlterarDespesa(strDtPagamento As String, _
                          strDtVencimento As String, _
                          strValorPagar As String, _
                          lngLivroId As Long, _
                          strNumeroCheque As String, _
                          lngDESPESAID As Long, _
                          strDescricao As String, _
                          strValorPago As String, _
                          strVale As String, _
                          lngSUBGRUPODESPESAID As Long, _
                          lngFormaPgtoId As Long, _
                          strUsuarioAlterouDescricao As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE DESPESA SET " & _
    " DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & _
    ", VR_PAGO = " & Formata_Dados(strValorPago, tpDados_Moeda, tpNulo_Aceita) & _
    ", VALE = " & Formata_Dados(strVale, tpDados_Texto, tpNulo_Aceita) & _
    ", SUBGRUPODESPESAID = " & Formata_Dados(lngSUBGRUPODESPESAID, tpDados_Longo, tpNulo_Aceita) & _
    ", FORMAPGTOID = " & Formata_Dados(IIf(lngFormaPgtoId = 0, "", lngFormaPgtoId), tpDados_Longo, tpNulo_Aceita) & _
    ", DT_VENCIMENTO = " & Formata_Dados(strDtVencimento, tpDados_DataHora, tpNulo_Aceita) & _
    ", DT_PAGAMENTO = " & Formata_Dados(strDtPagamento, tpDados_DataHora, tpNulo_Aceita) & _
    ", VR_PAGAR = " & Formata_Dados(strValorPagar, tpDados_Moeda, tpNulo_Aceita) & _
    ", LIVROID = " & Formata_Dados(IIf(lngLivroId = 0, "", lngLivroId), tpDados_Longo, tpNulo_Aceita) & _
    ", NUMEROCHEQUE = " & Formata_Dados(strNumeroCheque, tpDados_Texto, tpNulo_Aceita) & _
    ", USUARIOALTDESCR = " & Formata_Dados(strUsuarioAlterouDescricao, tpDados_Texto, tpNulo_Aceita) & _
    " WHERE PKID = " & lngDESPESAID
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsDespesa.AlterarDespesa]"
End Sub
Public Sub IncluirDespesa(strDtVencimento As String, _
                          strValorPagar As String, _
                          lngLivroId As Long, _
                          strNumeroCheque As String, _
                          strTipo As String, _
                          strDtPagamento As String, _
                          strDescricao As String, _
                          strValorPago As String, _
                          strVale As String, _
                          lngSUBGRUPODESPESAID As Long, _
                          lngFormaPgtoId As Long, _
                          strUsuario As String, _
                          ByRef lngSEQUENCIALEXTERNO As Long, _
                          lngTurnoId As Long)
  On Error GoTo trata
  '
  Dim strSql          As String
  'Dim lngSequencial   As Long
  '
  'lngSequencial = RetonaProximoSequencial("DESPESA", "SEQUENCIAL")
  '
  strSql = "INSERT INTO DESPESA (" & _
    " DESCRICAO, VR_PAGO, VALE, SUBGRUPODESPESAID, FORMAPGTOID" & _
    ", DT_PAGAMENTO, SEQUENCIAL, TIPO, DT_VENCIMENTO, VR_PAGAR, LIVROID, NUMEROCHEQUE, USUARIO, TURNOID) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strValorPago, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(strVale, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(lngSUBGRUPODESPESAID, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(IIf(lngFormaPgtoId = 0, "", lngFormaPgtoId), tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strDtPagamento, tpDados_DataHora, tpNulo_Aceita) & _
    ", " & Formata_Dados(lngSEQUENCIALEXTERNO, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strTipo, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strDtVencimento, tpDados_DataHora, tpNulo_Aceita) & _
    ", " & Formata_Dados(strValorPagar, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(IIf(lngLivroId = 0, "", lngLivroId), tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strNumeroCheque, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strUsuario, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(IIf(lngTurnoId = 0, "", lngTurnoId), tpDados_Longo, tpNulo_Aceita) & _
    ")"

  RunSP strSql
  'lngSEQUENCIALEXTERNO = lngSequencial
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsDespesa.IncluirDespesa]"
End Sub

