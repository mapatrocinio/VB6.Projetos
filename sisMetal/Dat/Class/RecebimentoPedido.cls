VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecebimentoPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'''Public Function SelecionarUltCCDaLoc(lngPEDIDOVENDAID As Long, _
'''                                     strStatusLancamento As String) As ADODB.Recordset
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "SELECT MAX(CONTACORRENTE.PKID) AS MAXPKID " & _
'''    " FROM CONTACORRENTE " & _
'''    " WHERE CONTACORRENTE."
'''
'''  Select Case strStatusLancamento & ""
'''  Case "RE", "CC", "RC", "DP"
'''    strSql = strSql & "LOCACAOID"
'''  Case "DE"
'''    strSql = strSql & "DESPESAID"
'''  Case "VD"
'''    strSql = strSql & "VENDAID"
'''  Case "EX"
'''    strSql = strSql & "EXTRAID"
'''  Case Else
'''    strSql = strSql & "" 'Para forçar erro
'''  End Select
'''
'''
'''  strSql = strSql & " = " & Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo)
'''
'''  Set SelecionarUltCCDaLoc = RunSPReturnRS(strSql)
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsContaCorrente.SelecionarUltCCDaLoc]", _
'''            Err.Description
'''End Function
'''
'''Public Function SelecionarPagamentos(strStatusLanc As String, _
'''                                     lngLOCDESPVDAEXT As Long) As ADODB.Recordset
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "SELECT CONTACORRENTE.*, " & _
'''    " case when STATUSCC = 'CC' then 'Cartão de Crédito' " & _
'''    " when STATUSCC = 'CD' then 'Cartão de Débito' " & _
'''    " when STATUSCC = 'ES' then 'Espécie' " & _
'''    " when STATUSCC = 'PH' then 'Penhor' " & _
'''    " when STATUSCC = 'CH' AND NOT DESPESAID IS NULL then 'Cheque Resgatado' " & _
'''    " when STATUSCC = 'CH' then 'Cheque' " & _
'''    " when STATUSCC = 'FT' then 'Fatura' end AS DESC_STATUSCC " & _
'''    " FROM CONTACORRENTE "
'''  Select Case strStatusLanc
'''  Case "RE", "CC", "RC", "DP"
'''    strSql = strSql & " WHERE CONTACORRENTE.LOCACAOID = "
'''  Case "DE"
'''    strSql = strSql & " WHERE CONTACORRENTE.DESPESAID = "
'''  Case "VD"
'''    strSql = strSql & " WHERE CONTACORRENTE.VENDAID = "
'''  Case "EX"
'''    strSql = strSql & " WHERE CONTACORRENTE.EXTRAID = "
'''  End Select
'''  strSql = strSql & Formata_Dados(lngLOCDESPVDAEXT, tpDados_Longo)
'''
'''  Set SelecionarPagamentos = RunSPReturnRS(strSql)
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsContaCorrente.SelecionarPagamentos]", _
'''            Err.Description
'''End Function

Public Function SelecionarContaCorrente(lngContaCorrentePedId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT CONTACORRENTEPED.*, CARTAO.NOME AS NOME_CARTAO, BANCO.NOME AS NOME_BANCO, CARTAODEBITO.NOME AS NOME_CARTAODEBITO, " & _
    " case INDDEBITOCREDITO when 'D' then 'Débito' else 'Crédito' end AS DESC_INDDEBCRED, " & _
    " case STATUSCC when 'CC' then 'Cartão de Crédito' " & _
    " when 'CD' then 'Cartão de Débito' " & _
    " when 'ES' then 'Espécie' " & _
    " when 'PH' then 'Penhor' " & _
    " when 'CH' then 'Cheque' " & _
    " when 'FT' then 'Fatura' end AS DESC_STATUSCC " & _
    " FROM CONTACORRENTEPED " & _
    " LEFT JOIN CARTAO ON CARTAO.PKID = CONTACORRENTEPED.CARTAOID " & _
    " LEFT JOIN BANCO ON BANCO.PKID = CONTACORRENTEPED.BANCOID " & _
    " LEFT JOIN CARTAODEBITO ON CARTAODEBITO.PKID = CONTACORRENTEPED.CARTAODEBITOID " & _
    " WHERE CONTACORRENTEPED.PKID = " & _
    Formata_Dados(lngContaCorrentePedId, tpDados_Longo, tpNulo_NaoAceita)

  Set SelecionarContaCorrente = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsRecebimentoPedido.SelecionarContaCorrente]", _
            Err.Description
End Function

Public Sub ExcluirContaCorrente(lngContaCorrentePedId As Long)
  On Error GoTo trata
  '
  Dim strSql          As String
  strSql = "DELETE FROM CONTACORRENTEPED WHERE PKID = " & _
    Formata_Dados(lngContaCorrentePedId, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsRecebimentoPedido.ExcluirContaCorrente]", _
            Err.Description
End Sub

'''Public Sub ExcluirContaCorrenteToda(strStatusLanc As String, _
'''                                    lngPEDIDOVENDAID As Long)
'''  On Error GoTo trata
'''  '
'''  Dim strSql          As String
'''  strSql = "DELETE CONTACORRENTEPED WHERE "
'''  Select Case strStatusLanc & ""
'''  Case "RE", "CC", "RC", "DP"
'''    strSql = strSql & "LOCACAOID"
'''  Case "DE"
'''    strSql = strSql & "DESPESAID"
'''  Case "VD"
'''    strSql = strSql & "VENDAID"
'''  Case "EX"
'''    strSql = strSql & "EXTRAID"
'''  Case Else
'''    strSql = strSql & "" 'Para forçar erro
'''  End Select
'''
'''
'''  strSql = strSql & " = " & _
'''    Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo, tpNulo_NaoAceita)
'''  RunSP strSql
'''  '
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsRecebimentoPedido.ExcluirContaCorrente]", _
'''            Err.Description
'''End Sub

Public Function InserirCC(lngPEDIDOVENDAID As Long, _
                          strDtHora As String, _
                          strValor As String, _
                          strIndDebitoCredito As String, _
                          strStatusCC As String, _
                          strCartaoId As String, _
                          strBancoId As String, _
                          strResponsavel As String, _
                          strAgencia As String, _
                          strConta As String, _
                          strNroCheque As String, _
                          strCPF As String, _
                          strCliente As String, _
                          strLote As String, _
                          strCartaoDebId As String, _
                          strVrCalcTroco As String, _
                          strNroParcelas As String, _
                          strDtPrimeiraParc As String) As Long

  On Error GoTo trata
  '
  Dim objRs As ADODB.Recordset
  Dim strSql As String
  '
  strSql = "SET NOCOUNT ON; "

  strSql = strSql & "INSERT INTO CONTACORRENTEPED ("

  strSql = strSql & "PEDIDOVENDAID, DTHORACC, VALOR, INDDEBITOCREDITO, STATUSCC, CARTAOID, BANCOID, RESPONSAVEL, AGENCIA, CONTA, NROCHEQUE, CPF, CLIENTE, LOTE, CARTAODEBITOID, VRTROCO, NROPARCELAS, DTPRIMEIRAPARCELA) VALUES (" & _
    Formata_Dados(lngPEDIDOVENDAID, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strDtHora, tpDados_DataHora, tpNulo_Aceita) & _
    ", " & Formata_Dados(strValor, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(strIndDebitoCredito, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strStatusCC, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strCartaoId, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strBancoId, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strResponsavel, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strAgencia, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strConta, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strNroCheque, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strCPF, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strCliente, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strLote, tpDados_Texto, tpNulo_Aceita) & ", " & Formata_Dados(strCartaoDebId, tpDados_Longo, tpNulo_Aceita) & ", " & Formata_Dados(strVrCalcTroco, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(strNroParcelas, tpDados_Longo) & _
    ", " & Formata_Dados(strDtPrimeiraParc, tpDados_DataHora) & _
    ");"
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    InserirCC = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  'InserirCC = 0
  'InserirCC = ExecSPReturnInteger("sp_conta_corrente_inserir", "")
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsRecebimentoPedido.InserirCC]", _
            Err.Description
End Function

'''Public Sub AssociaCCDespesa(strDespesaId As String, _
'''                            strCCId As String)
'''
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "UPDATE CONTACORRENTE SET " & _
'''    "DESPESAID = " & Formata_Dados(strDespesaId, tpDados_Longo) & _
'''    " WHERE PKID = " & Formata_Dados(strCCId, tpDados_Longo)
'''  RunSP strSql
'''  '
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsContaCorrente.AssociaCCDespesa]", _
'''            Err.Description
'''End Sub
'''
'''Public Sub AlterarCCValor(lngCCId As Long, _
'''                          strValor As String)
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "UPDATE CONTACORRENTE SET " & _
'''    "VALOR = " & Formata_Dados(strValor, tpDados_Moeda, tpNulo_Aceita) & _
'''    " WHERE PKID = " & Formata_Dados(lngCCId, tpDados_Longo, tpNulo_Aceita)
'''  RunSP strSql
'''  '
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsContaCorrente.AlterarCCValor]", _
'''            Err.Description
'''End Sub

Public Sub AlterarCC(strStatusCC As String, _
                     lngCCId As Long, _
                     strValor As String, _
                     strCartaoId As String, _
                     strBancoId As String, _
                     strAgencia As String, _
                     strConta As String, _
                     strNroCheque As String, _
                     strCPF As String, _
                     strCliente As String, _
                     strLote As String, _
                     strCartaoDebId As String, _
                     strVrCalcTroco As String, _
                     strNroParcelas As String, _
                     strDtPrimeiraParc As String)

  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE CONTACORRENTEPED SET " & _
    "VALOR = " & Formata_Dados(strValor, tpDados_Moeda, tpNulo_Aceita) & _
    ", STATUSCC = " & Formata_Dados(strStatusCC, tpDados_Texto) & _
    ", CARTAOID = " & Formata_Dados(strCartaoId, tpDados_Longo, tpNulo_Aceita) & _
    ", BANCOID = " & Formata_Dados(strBancoId, tpDados_Longo, tpNulo_Aceita) & _
    ", AGENCIA = " & Formata_Dados(strAgencia, tpDados_Texto, tpNulo_Aceita) & _
    ", CONTA = " & Formata_Dados(strConta, tpDados_Texto, tpNulo_Aceita) & _
    ", NROCHEQUE = " & Formata_Dados(strNroCheque, tpDados_Texto, tpNulo_Aceita) & _
    ", CPF = " & Formata_Dados(strCPF, tpDados_Texto, tpNulo_Aceita) & _
    ", CLIENTE = " & Formata_Dados(strCliente, tpDados_Texto, tpNulo_Aceita) & _
    ", LOTE = " & Formata_Dados(strLote, tpDados_Texto, tpNulo_Aceita) & _
    ", CARTAODEBITOID = " & Formata_Dados(strCartaoDebId, tpDados_Longo, tpNulo_Aceita) & _
    ", VRTROCO = " & Formata_Dados(strVrCalcTroco, tpDados_Moeda, tpNulo_Aceita) & _
    ", NROPARCELAS = " & Formata_Dados(strNroParcelas, tpDados_Longo) & _
    ", DTPRIMEIRAPARCELA = " & Formata_Dados(strDtPrimeiraParc, tpDados_DataHora) & _
    " WHERE PKID = " & Formata_Dados(lngCCId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsRecebimentoPedido.AlterarCC]", _
            Err.Description
End Sub


