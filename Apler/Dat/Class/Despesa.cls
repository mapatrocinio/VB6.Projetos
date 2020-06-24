VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDespesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarDespesaPeloSequencial(lngSEQUENCIAL As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT FORMAPGTO.FORMAPGTO AS DESCRFORMAPGTO, GRUPODESPESA.CODIGO AS CODIGOGRUPODESPESA, SUBGRUPODESPESA.CODIGO AS CODIGOSUBGRUPODESPESA, DESPESA.* FROM ((DESPESA " & _
    " LEFT JOIN SUBGRUPODESPESA ON SUBGRUPODESPESA.PKID = DESPESA.SUBGRUPODESPESAID) " & _
    " LEFT JOIN GRUPODESPESA ON GRUPODESPESA.PKID = SUBGRUPODESPESA.GRUPODESPESAID) " & _
    " LEFT JOIN FORMAPGTO ON FORMAPGTO.PKID = DESPESA.FORMAPGTOID " & _
    " WHERE DESPESA.SEQUENCIAL = " & _
    Formata_Dados(lngSEQUENCIAL, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarDespesaPeloSequencial = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDespesa.SelecionarDespesaPeloSequencial]", _
            Err.Description
End Function

Public Function SelecionarDespesa(lngDESPESAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT FUNCIONARIO.NOME AS NOME_FUNCIONARIO, FORMAPGTO.FORMAPGTO AS DESCRFORMAPGTO, GRUPODESPESA.CODIGO AS CODIGOGRUPODESPESA, SUBGRUPODESPESA.CODIGO AS CODIGOSUBGRUPODESPESA, DESPESA.*, " & _
    " TOT_PGTO_OUTRA_FORMA = isnull((SELECT COUNT(*) FROM CONTACORRENTE WHERE DESPESAID = " & _
    Formata_Dados(lngDESPESAID, tpDados_Longo) & _
    " AND STATUSCC <> 'ES'),0) " & _
    " FROM DESPESA " & _
    " LEFT JOIN SUBGRUPODESPESA ON SUBGRUPODESPESA.PKID = DESPESA.SUBGRUPODESPESAID " & _
    " LEFT JOIN GRUPODESPESA ON GRUPODESPESA.PKID = SUBGRUPODESPESA.GRUPODESPESAID " & _
    " LEFT JOIN FORMAPGTO ON FORMAPGTO.PKID = DESPESA.FORMAPGTOID " & _
    " LEFT JOIN FUNCIONARIO ON FUNCIONARIO.PKID = DESPESA.FUNCIONARIOID " & _
    " WHERE DESPESA.PKID = " & _
    Formata_Dados(lngDESPESAID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarDespesa = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDespesa.SelecionarDespesa]", _
            Err.Description
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
  Err.Raise Err.Number, _
            Err.Source & ".[clsDespesa.ExcluirDespesa]", _
            Err.Description
End Sub

Public Sub AlterarDespesa(lngDESPESAID As Long, _
                          strDescricao As String, _
                          strValorPago As String, _
                          strVale As String, _
                          lngSubGrupoDespesaId As Long, _
                          lngFormaPgtoId As Long, _
                          strUsuarioAutorizacao As String, _
                          lngFUNCIONARIOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE DESPESA SET " & _
    " DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & _
    ", VR_PAGO = " & Formata_Dados(strValorPago, tpDados_Moeda, tpNulo_Aceita) & _
    ", VALE = " & Formata_Dados(strVale, tpDados_Texto, tpNulo_Aceita) & _
    ", SUBGRUPODESPESAID = " & Formata_Dados(IIf(lngSubGrupoDespesaId = 0, "", lngSubGrupoDespesaId), tpDados_Longo, tpNulo_Aceita) & _
    ", FORMAPGTOID = " & Formata_Dados(lngFormaPgtoId, tpDados_Longo, tpNulo_Aceita) & _
    ", USUARIOAUTORIZACAO = " & Formata_Dados(strUsuarioAutorizacao, tpDados_Texto, tpNulo_Aceita) & _
    ", FUNCIONARIOID = " & Formata_Dados(IIf(lngFUNCIONARIOID = 0, "", lngFUNCIONARIOID), tpDados_Longo) & _
    " WHERE PKID = " & lngDESPESAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDespesa.AlterarDespesa]", _
            Err.Description
End Sub
Public Function IncluirDespesa(strSequencial As String, _
                               strTipo As String, _
                               lngTurnoId As Long, _
                               strDtPagamento As String, _
                               strDescricao As String, _
                               strValorPago As String, _
                               strVale As String, _
                               lngSubGrupoDespesaId As Long, _
                               lngFormaPgtoId As Long, _
                               strUsuario As String, _
                               strUsuarioAutorizacao As String, _
                               lngFUNCIONARIOID As Long) As Long
  On Error GoTo trata
  '
  Dim strSql          As String
  Dim lngDESPESAID    As Long
  Dim objRs           As ADODB.Recordset
  '
  strSql = "INSERT INTO DESPESA (" & _
    " DESCRICAO, VR_PAGO, VALE, SUBGRUPODESPESAID, FORMAPGTOID" & _
    ", TURNOID, DT_PAGAMENTO, SEQUENCIAL, TIPO, USUARIO, USUARIOAUTORIZACAO, FUNCIONARIOID) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strValorPago, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(strVale, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(IIf(lngSubGrupoDespesaId = 0, "", lngSubGrupoDespesaId), tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(lngFormaPgtoId, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(lngTurnoId, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strDtPagamento, tpDados_DataHora, tpNulo_Aceita) & _
    ", " & Formata_Dados(strSequencial, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strTipo, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strUsuario, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strUsuarioAutorizacao, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(IIf(lngFUNCIONARIOID = 0, "", lngFUNCIONARIOID), tpDados_Longo) & _
    ")"

  RunSP strSql
  '
  Set objRs = SelecionarDespesaPeloSequencial(CLng(strSequencial))
  If Not objRs.EOF Then
    lngDESPESAID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  IncluirDespesa = lngDESPESAID
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDespesa.IncluirDespesa]", _
            Err.Description
End Function

