VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContaCorrente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'strNomeUsu = gsNomeUsu
'icTipoGR = tpIcTipoGR_Prest

Public Sub InserirFinanceiro(lngGRID As Long, _
                             icTipoGR As Integer, _
                             strNomeUsu As String)
  On Error GoTo trata
  Dim objCC             As datSisMed.clsContaCorrente
  Dim objGeral          As datSisMed.clsGeral
  Dim lngPRESTADORID    As Long
  Dim lngGRPAGAMENTOID  As Long
  Dim lngFINANCEIROID   As Long
  Dim strStatus         As String
  Dim strStatusIncAlt   As String
  Dim strStatusFinIncAlt  As String
  Dim datData           As Date
  Dim strDataIni        As String
  Dim strDataFim        As String
  Dim objRs             As ADODB.Recordset
  Dim strSql            As String
  Dim objGRPagamento    As datSisMed.clsGRPagamento
  Dim objGRPgto         As datSisMed.clsGRPgto
  '
  Set objGeral = New datSisMed.clsGeral
  '
  'Verifica status
  Select Case icTipoGR
  Case tpIcTipoGR_DonoRX: strStatus = "DR"
  Case tpIcTipoGR_DonoUltra: strStatus = "DU"
  Case tpIcTipoGR_Prest: strStatus = "PG"
  Case tpIcTipoGR_TecRX: strStatus = "TR"
  Case tpIcTipoGR_CancPont: strStatus = "CP"
  Case tpIcTipoGR_CancAut: strStatus = "CA"
  Case Else: strStatus = ""
  End Select
  'PRESTADOR
  lngPRESTADORID = 0
  strSql = "SELECT PRESTADOR.PKID " & _
    "FROM GR INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
      " INNER JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
      " INNER JOIN PRONTUARIO AS PRESTADOR ON PRESTADOR.PKID = ATENDE.PRONTUARIOID " & _
    "WHERE GR.PKID = " & Formata_Dados(lngGRID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngPRESTADORID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'Verifica se irá incluir ou alterar grpagaento
  datData = Now
  strDataIni = Format(datData, "DD/MM/YYYY")
  strDataFim = Format(datData, "DD/MM/YYYY 23:59:59")
  strSql = "SELECT * FROM GRPAGAMENTO " & _
    " WHERE GRPAGAMENTO.PRESTADORID = " & Formata_Dados(lngPRESTADORID, tpDados_Longo) & _
    " AND GRPAGAMENTO.DATAINICIO = " & Formata_Dados(strDataIni, tpDados_DataHora) & _
    " AND GRPAGAMENTO.STATUS = " & Formata_Dados(strStatus, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  lngGRPAGAMENTOID = 0
  If Not objRs.EOF Then
    'Já cadastrado
    strStatusIncAlt = "A"
    lngGRPAGAMENTOID = objRs.Fields("PKID").Value
  Else
    'Não cadastrado
    strStatusIncAlt = "I"
  End If
  objRs.Close
  Set objRs = Nothing
  'Verifica se irá incluir ou alterar grpagaento
  strSql = "SELECT * FROM FINANCEIRO " & _
    " WHERE FINANCEIRO.GRID = " & Formata_Dados(lngGRID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  lngFINANCEIROID = 0
  If Not objRs.EOF Then
    'Já cadastrado
    strStatusFinIncAlt = "A"
    lngFINANCEIROID = objRs.Fields("PKID").Value
  Else
    'Não cadastrado
    strStatusFinIncAlt = "I"
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set objGeral = Nothing
  '
  'Inserir FINANCEIRO
  'Obeter os valores para inserir na tabela de financeiro
  'Inserir/Alterar FINANCEIRO
  Set objCC = New datSisMed.clsContaCorrente
  If strStatusIncAlt = "A" Then
    'Alterar FINANCEIRO
    'objCC.AlterarFinanceiro lngFINANCEIROID, _
                                      lngPRESTADORID, _
                                      strDataIni, _
                                      strDataFim
    '
  Else
    'Inserir FINANCEIRO
    'objCC.InserirFinanceiro lngGRID, _
                                      lngPRESTADORID, _
                                      strDataIni, _
                                      strDataFim, _
                                      strStatus, _
                                      "N", _
                                      strNomeUsu

    '
  End If
  Set objCC = Nothing
  '
  'Inserir/Alterar GRPAGAMENTO
  Set objGRPagamento = New datSisMed.clsGRPagamento
  If strStatusIncAlt = "A" Then
    'Alterar GRPAGAMENTO
    'objGRPagamento.AlterarGRPagamento lngGRPAGAMENTOID, _
                                      lngPRESTADORID, _
                                      strDataIni, _
                                      strDataFim
    '
  Else
    'Inserir GRPAGAMENTO
    objGRPagamento.InserirGRPagamento lngGRPAGAMENTOID, _
                                      lngPRESTADORID, _
                                      strDataIni, _
                                      strDataFim, _
                                      strStatus, _
                                      "N", _
                                      strNomeUsu

    '
  End If
  Set objGRPagamento = Nothing
  'Associar GR a GRPAGAMENTO
  Set objGRPgto = New datSisMed.clsGRPgto
  '
  objGRPgto.InserirGRPGTOGR lngGRID, _
                            lngGRPAGAMENTOID
  Set objGRPgto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Function SelecionarPagamentos(strStatusLanc As String, _
                                     lngGR As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objCC As datSisMed.clsContaCorrente
  Set objCC = New datSisMed.clsContaCorrente
  
  'Selecionar pagamentos da Conta Corrente
  Set SelecionarPagamentos = objCC.SelecionarPagamentos(strStatusLanc, _
                                                        lngGR)
  '
  Set objCC = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarContaCorrente(lngContaCorrenteId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objCC As datSisMed.clsContaCorrente
  Set objCC = New datSisMed.clsContaCorrente
  
  'Listar Conta Corrente
  Set SelecionarContaCorrente = objCC.SelecionarContaCorrente(lngContaCorrenteId)
  '
  Set objCC = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirContaCorrente(strStatusLanc As String, _
                                lngContaCorrenteId As Long)

  On Error GoTo trata
  Dim objContaCorrente  As datSisMed.clsContaCorrente
  Dim objRs             As ADODB.Recordset
  Set objContaCorrente = New datSisMed.clsContaCorrente
  
  'Excluir ContaCorrente
  objContaCorrente.ExcluirContaCorrente strStatusLanc, _
                                        lngContaCorrenteId
  '
  Set objContaCorrente = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ExcluirContaCorrenteToda(strStatusLanc As String, _
                                    lngGRID As Long)

  On Error GoTo trata
  Dim objContaCorrente  As datSisMed.clsContaCorrente
  Dim objRs             As ADODB.Recordset
  Set objContaCorrente = New datSisMed.clsContaCorrente
  
  'Excluir Toda ContaCorrente
  objContaCorrente.ExcluirContaCorrenteToda strStatusLanc, _
                                            lngGRID
  '
  Set objContaCorrente = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function InserirCC(lngGRID As Long, _
                          lngTURNOCCID As Long, strDtHora As String, strValor As String, strIndDebitoCredito As String, strStatusCC As String, _
                          strStatusLancamento As String, _
                          strCartaoId As String, _
                          strBancoId As String, _
                          strFuncionarioId As String, _
                          strResponsavel As String, _
                          strAgencia As String, _
                          strConta As String, _
                          strNroCheque As String, _
                          strCPF As String, _
                          strCliente As String, _
                          strObjeto As String, _
                          strDocumentoPenhor As String, _
                          strLote As String, _
                          strCartaoDebId As String, _
                          strVrGorjeta As String, _
                          strVrCalcTroco As String, _
                          strNroParcelas As String, _
                          strDtPrimeiraParc As String) As Long
  On Error GoTo trata
  Dim clsCC   As datSisMed.clsContaCorrente
  Dim objRs   As ADODB.Recordset
  '
  Set clsCC = New datSisMed.clsContaCorrente
  InserirCC = 0
  'Inserir Conta Corrente
  
  InserirCC = clsCC.InserirCC(lngGRID, _
                              lngTURNOCCID, _
                              strDtHora, _
                              strValor, _
                              strIndDebitoCredito, _
                              strStatusCC, _
                              strStatusLancamento, _
                              strCartaoId, _
                              strBancoId, _
                              strFuncionarioId, _
                              strResponsavel, _
                              strAgencia, _
                              strConta, _
                              strNroCheque, _
                              strCPF, _
                              strCliente, _
                              strObjeto, _
                              strDocumentoPenhor, _
                              strLote, _
                              strCartaoDebId, _
                              strVrGorjeta, _
                              strVrCalcTroco, _
                              strNroParcelas, _
                              strDtPrimeiraParc)
  
  Set objRs = clsCC.SelecionarUltCCDaGR(lngGRID, _
                                        strStatusLancamento)
  If Not objRs.EOF Then
    If IsNumeric(objRs.Fields("MAXPKID").Value) Then
      InserirCC = objRs.Fields("MAXPKID").Value
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set clsCC = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Sub AlterarCCValor(lngCCId As Long, _
                          strValor As String)
  On Error GoTo trata
  Dim clsCC As datSisMed.clsContaCorrente
  '
  Set clsCC = New datSisMed.clsContaCorrente
  
  'Alterar Conta Corrente
  
  clsCC.AlterarCCValor lngCCId, _
                       strValor
  '
  Set clsCC = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub




Public Sub AlterarCC(strStatusCC As String, _
                     lngCCId As Long, _
                     strValor As String, _
                     strCartaoId As String, _
                     strBancoId As String, _
                     strFuncionarioId As String, _
                     strAgencia As String, _
                     strConta As String, _
                     strNroCheque As String, _
                     strCPF As String, _
                     strCliente As String, _
                     strObjeto As String, _
                     strDocumentoPenhor As String, _
                     strLote As String, _
                     strCartaoDebId As String, _
                     strVrGorjeta As String, _
                     strVrCalcTroco As String, _
                     strNroParcelas As String, _
                     strDtPrimeiraParc As String)
  On Error GoTo trata
  Dim clsCC As datSisMed.clsContaCorrente
  '
  Set clsCC = New datSisMed.clsContaCorrente
  
  'Alterar Conta Corrente
  
  clsCC.AlterarCC strStatusCC, _
                  lngCCId, _
                  strValor, _
                  strCartaoId, _
                  strBancoId, _
                  strFuncionarioId, _
                  strAgencia, _
                  strConta, _
                  strNroCheque, _
                  strCPF, _
                  strCliente, _
                  strObjeto, _
                  strDocumentoPenhor, _
                  strLote, _
                  strCartaoDebId, _
                  strVrGorjeta, _
                  strVrCalcTroco, _
                  strNroParcelas, _
                  strDtPrimeiraParc

  '
  Set clsCC = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


