VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLocacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub AlterarLocRecCC(lngLocacaoId As Long, _
                           lngTurnoRecebeId As Long, _
                           blnRecebimento As Boolean, _
                           blnAtualizarTurnoRecCC As Boolean)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Alterar Locação - RECEBIMENTO
  
  clsLoc.AlterarLocRecCC lngLocacaoId, _
                         lngTurnoRecebeId, _
                         blnRecebimento, _
                         blnAtualizarTurnoRecCC

  '
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarLocRecebimento(lngLocacaoId As Long, _
                                 lngTurnoRecebeId As Long, _
                                 strGarcomId As String, _
                                 strDtHoraReceb As String, _
                                 strResponsavel As String, _
                                 blnCobranca As Boolean, _
                                 strPgtoEspecie As String, _
                                 strGorjeta As String, _
                                 strPgtoCartao As String, _
                                 strPgtoCartaoDebito As String, _
                                 strCartaoId As String, strCartaoDebId As String, _
                                 strLote As String, _
                                 strPgtoCheque As String, _
                                 strCPF As String, _
                                 strNroCheque As String, _
                                 strBancoId As String, _
                                 strAgencia As String, _
                                 strConta As String, _
                                 strPgtoPenhor As String, _
                                 strCliente As String, _
                                 strDocumentoPenhor As String, _
                                 strObjeto As String, _
                                 strVrCalcTotalPago As String, _
                                 strVrCalcTroco As String, _
                                 blnRecebimento As Boolean)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Alterar Locação - RECEBIMENTO
  
  clsLoc.AlterarLocRecebimento lngLocacaoId, _
                               lngTurnoRecebeId, _
                               strGarcomId, _
                               strDtHoraReceb, _
                               strResponsavel, _
                               blnCobranca, _
                               strPgtoEspecie, _
                               strGorjeta, _
                               strPgtoCartao, _
                               strPgtoCartaoDebito, _
                               strCartaoId, strCartaoDebId, _
                               strLote, _
                               strPgtoCheque, _
                               strCPF, _
                               strNroCheque, _
                               strBancoId, _
                               strAgencia, _
                               strConta, _
                               strPgtoPenhor, _
                               strCliente, _
                               strDocumentoPenhor, _
                               strObjeto, _
                               strVrCalcTotalPago, _
                               strVrCalcTroco, _
                               blnRecebimento

  '
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub AlterarLocLimpeza(lngLocacaoId As Long, _
                             blnLimpeza As Boolean, _
                             blnOcupado As Boolean, _
                             lngCamareiraLimpezaId As Long)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Alterar Locação - LIMPEZA

  clsLoc.AlterarLocLimpeza lngLocacaoId, _
                           blnLimpeza, _
                           blnOcupado, _
                           lngCamareiraLimpezaId
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub AlterarLocEncerrar(lngLocacaoId As Long, _
                              blnOcupado As Boolean)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Alterar Locação - ENCERRAR

  clsLoc.AlterarLocEncerrar lngLocacaoId, _
                            blnOcupado
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarLocPeriodoSomaVenc(lngLocacaoId As Long)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Alterar Locação - PERIODO SOMA VENCIMENTO

  clsLoc.AlterarLocPeriodoSomaVenc lngLocacaoId
  
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarLocLiberacao(lngLocacaoId As Long, _
                               blnLimpo As Boolean, _
                               blnLiberado As Boolean, _
                               lngCamareiraLiberaId As Long)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Alterar Locação - LIMPEZA

  clsLoc.AlterarLocLiberacao lngLocacaoId, _
                             blnLimpo, _
                             blnLiberado, _
                             lngCamareiraLiberaId
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub AlterarLocMovAposFecha(lngLocacaoId As Long, _
                                  blnMOVAPOSFECHA As Boolean, _
                                  blnMOVAPOSFECHAOK As Boolean)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Alterar Locação - MOVIMENTO APÓS O FECHAMENTO

  clsLoc.AlterarLocMovAposFecha lngLocacaoId, _
                                blnMOVAPOSFECHA, _
                                blnMOVAPOSFECHAOK
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarLocOcupado(lngLocacaoId As Long, _
                             blnOcupado As Boolean)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Alterar Locação - OCUPADO

  clsLoc.AlterarLocOcupado lngLocacaoId, _
                           blnOcupado
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarLocEntradaParc(lngLocacaoId As Long, _
                                strPlacaVeiculo As String, _
                                blnCortesia As Boolean, _
                                strFaixa As String, _
                                strDataHoraEntra As String, _
                                strDataPrevEntrada As String, _
                                strDataPrevSaida As String, _
                                strVrDeposito As String, _
                                strApartamentoId As String, _
                                strTurnoEntradaId As String, _
                                strGrupoPeriodoId As String, _
                                blnDeposito As Boolean, _
                                strOcupantes As String, _
                                strAutoriacao As String, _
                                strIdentPeriodos As String, _
                                strCheque As String, _
                                strCartao As String, _
                                strDocumento As String, _
                                strHospede As String, _
                                lngCARTAOPROMOCIONALID As Long, _
                                strMotivoDesconto As String, strPercentualDesconto As String, strDescCortesia As String, strSobreTotal As String, _
                                strReservaId As String, strPacoteId As String)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Alterar Locação - Entrada Parcial

  clsLoc.AlterarLocEntradaParc lngLocacaoId, _
                                strPlacaVeiculo, _
                                blnCortesia, _
                                strFaixa, _
                                strDataHoraEntra, _
                                strDataPrevEntrada, _
                                strDataPrevSaida, _
                                strVrDeposito, _
                                strApartamentoId, _
                                strTurnoEntradaId, _
                                strGrupoPeriodoId, _
                                blnDeposito, _
                                strOcupantes, _
                                strAutoriacao, _
                                strIdentPeriodos, _
                                strCheque, _
                                strCartao, _
                                strDocumento, _
                                strHospede, _
                                lngCARTAOPROMOCIONALID, _
                            strMotivoDesconto, strPercentualDesconto, strDescCortesia, strSobreTotal, _
                            strReservaId, strPacoteId
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarLocSaida(lngLocacaoId As Long, _
                           strDtHoraSaida As String, _
                           lngTurnoSaidaId As Long, _
                           blnSaiu As Boolean)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Alterar Locação - SAIDA

  clsLoc.AlterarLocSaida lngLocacaoId, _
                         strDtHoraSaida, _
                         lngTurnoSaidaId, _
                         blnSaiu
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub TratarLocCortesia(lngLocacaoId As Long)

  On Error GoTo trata
  Dim objGeral      As datSisContas.clsGeral
  Dim objLoc        As datSisContas.clsLocacao
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  Dim lngTOTALCORT  As Long
  Dim blnCortesia   As Boolean
  '
  Set objGeral = New datSisContas.clsGeral
  'Tratar Locação - CORTESIA
  strSql = "SELECT COUNT(*) AS TOTALCORT FROM CORTESIA " & _
      "WHERE LOCACAOID = " & lngLocacaoId
  Set objRs = objGeral.ExecutarSQL(strSql)
  If objRs.EOF Then
    lngTOTALCORT = 0
  ElseIf Not IsNumeric(objRs.Fields("TOTALCORT").Value) Then
    lngTOTALCORT = 0
  Else
    lngTOTALCORT = objRs.Fields("TOTALCORT").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  If lngTOTALCORT = 0 Then
    blnCortesia = False
  Else
    blnCortesia = True
  End If
  Set objGeral = Nothing
  '
  Set objLoc = New datSisContas.clsLocacao
  '
  objLoc.AlterarLocCortesia lngLocacaoId, _
                            blnCortesia
  '
  Set objLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarLocLiberarCheque(lngLocacaoId As Long, _
                                   strCPF As String)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Alterar Locação - LIBERAR CHEQUE

  
  clsLoc.AlterarLocLiberarCheque lngLocacaoId, _
                                 strCPF
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarLocVrEmprPrincAssoc(lngLocacaoId As Long)

  On Error GoTo trata
  Dim strValor      As String
  Dim strLocacaoId  As String
  Dim strSql        As String
  Dim objGeral      As datSisContas.clsGeral
  Dim clsLoc        As datSisContas.clsLocacao
  Dim objRs         As ADODB.Recordset
  '
  Set clsLoc = New datSisContas.clsLocacao
  Set objGeral = New datSisContas.clsGeral
  strValor = ""
  strSql = "SELECT MIN(ASSOCIACAO.LOCACAOID) AS PKIDLOC, SUM(VRCALCTOTALEMPRESA) AS VALOR " & _
    "FROM (ASSOCIACAO " & _
    " INNER JOIN TAB_LOCASSOC ON ASSOCIACAO.PKID = TAB_LOCASSOC.ASSOCIACAOID) " & _
    " INNER JOIN LOCACAO ON LOCACAO.PKID = TAB_LOCASSOC.LOCACAOID " & _
    " WHERE LOCACAO.PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    If IsNumeric(objRs.Fields("VALOR").Value) Then
      strLocacaoId = objRs.Fields("PKIDLOC").Value
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  If strLocacaoId <> "" Then
    strSql = "SELECT MIN(ASSOCIACAO.LOCACAOID) AS PKIDLOC, SUM(VRCALCTOTALEMPRESA) AS VALOR " & _
      "FROM (ASSOCIACAO " & _
      " INNER JOIN TAB_LOCASSOC ON ASSOCIACAO.PKID = TAB_LOCASSOC.ASSOCIACAOID) " & _
      " INNER JOIN LOCACAO ON LOCACAO.PKID = TAB_LOCASSOC.LOCACAOID " & _
      " WHERE ASSOCIACAO.LOCACAOID = " & Formata_Dados(strLocacaoId, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    
    If Not objRs.EOF Then
      If IsNumeric(objRs.Fields("VALOR").Value) Then
        'Soma
        strValor = Format(objRs.Fields("VALOR").Value, "###,##0.00")
      End If
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  Set objGeral = Nothing
  'Alterar Locação - Valor Empresa Principal Associada
  If strValor <> "" Then
    clsLoc.AlterarLocVrEmprPrincAssoc CLng(strLocacaoId), _
                                      strValor
  End If
  '
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarLocFechamento(lngLocacaoId As Long, strVrDesconto As String, strPercDesconto As String, strMotivoDesconto As String, strQtdFecha As String, _
                                blnFechamento As Boolean, _
                                strDtHoraFecha As String, _
                                blnMOVAPOSFECHAOK As Boolean, _
                                strDtHoraPriFecha As String, _
                                lngTurnoFechaId As Long, _
                                strVrHrExced As String, _
                                strQtdHrExced As String, _
                                strQtdPerExced As String, _
                                strVrPerExced As String, _
                                strQtdTotalPer As String, _
                                strVrCalcDiaria As String, _
                                strVrCalcConsumo As String, _
                                strVrCalcTelefonema As String, _
                                strVrCalcExcedPes As String, _
                                strVrCalcExtra As String, _
                                strVrCalcDesconto As String, _
                                strVrCalcTotal As String, _
                                lngCARTAOPROMOCIONALID As Long, _
                                blnCortesia As Boolean, _
                                strAutorizacao As String, _
                                strDescCortesia As String, _
                                strSobreTotal As String, _
                                strVrCalcDesc As String, strVrCalcCortDez As String, strVrCalcCortDia As String, strVrCalcCortDiaTot As String, strVrCalcCortTot As String, strVrCalcCortCons As String, _
                                strVrTotalCC As String, strVrExtraEmpr As String, strVrConsumoEmpr As String, strVrTelefoneEmpr As String, strQtdPerEmpr As String, strVrDiariaEmpr As String, strVrExcedPesEmpr As String, strVrTotalEmpresa As String, strVrTaxaEmpr As String, strVrTaxa As String, blnCobrarTaxa As Boolean)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Alterar Locação - FECHAMENTO

  
  clsLoc.AlterarLocFechamento lngLocacaoId, strVrDesconto, strPercDesconto, strMotivoDesconto, strQtdFecha, blnFechamento, _
                              strDtHoraFecha, _
                              blnMOVAPOSFECHAOK, _
                              strDtHoraPriFecha, _
                              lngTurnoFechaId, _
                              strVrHrExced, _
                              strQtdHrExced, _
                              strQtdPerExced, _
                              strVrPerExced, _
                              strQtdTotalPer, _
                              strVrCalcDiaria, _
                              strVrCalcConsumo, _
                              strVrCalcTelefonema, _
                              strVrCalcExcedPes, _
                              strVrCalcExtra, _
                              strVrCalcDesconto, _
                              strVrCalcTotal, _
                              lngCARTAOPROMOCIONALID, _
                              blnCortesia, _
                              strAutorizacao, _
                              strDescCortesia, _
                              strSobreTotal, _
                              strVrCalcDesc, strVrCalcCortDez, strVrCalcCortDia, strVrCalcCortDiaTot, strVrCalcCortTot, strVrCalcCortCons, _
                              strVrTotalCC, strVrExtraEmpr, strVrConsumoEmpr, strVrTelefoneEmpr, strQtdPerEmpr, strVrDiariaEmpr, strVrExcedPesEmpr, strVrTotalEmpresa, strVrTaxaEmpr, strVrTaxa, blnCobrarTaxa
  '
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarLocDtHoraCalcFecha(lngLocacaoId As Long, _
                                     strDtHoraCalcFecha As String)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Alterar Locação - FECHAMENTO

  
  clsLoc.AlterarLocDtHoraCalcFecha lngLocacaoId, _
                            strDtHoraCalcFecha
  '
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarLocEntradaPes(lngLocacaoId As Long, _
                                strQtdPessoasEntrada As String)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Alterar Locação - ENTRADA PESSOAS NA UNIDADE

  
  clsLoc.AlterarLocEntradaPes lngLocacaoId, _
                              strQtdPessoasEntrada
  '
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarLocDeposito(lngLocacaoId As Long, _
                              strVrDeposito As String, _
                              blnDeposito As Boolean, _
                              strCheque As String, _
                              strCartao As String, _
                              strDocumento As String, _
                              strHospede As String)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Alterar Locação - DEPÓSITO

  
  clsLoc.AlterarLocDeposito lngLocacaoId, _
                            strVrDeposito, _
                            blnDeposito, _
                            strCheque, _
                            strCartao, _
                            strDocumento, _
                            strHospede
  '
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub ExcluirLocacao(lngLocacaoId As Long, _
                          gbTrabSuiteAptoLimpo As Boolean)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Excluir Locação
  
  clsLoc.ExcluirLocacao lngLocacaoId, _
                        gbTrabSuiteAptoLimpo
  '
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarLocTransf(lngLocacaoId As Long, _
                            strApartamentoId As String, _
                            strGrupoPeriodoId As String, _
                            strDataPrevSaida As String, _
                            strNumeroAptoAnterior As String, _
                            gbTrabSuiteAptoLimpo As Boolean)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Alterar Locação - TRANSFERENCIA

  
  clsLoc.AlterarLocTransf lngLocacaoId, _
                          strApartamentoId, _
                          strGrupoPeriodoId, _
                          strDataPrevSaida, _
                          strNumeroAptoAnterior, _
                          gbTrabSuiteAptoLimpo
  '
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function InserirLocacao(blnOcupado As Boolean, _
                               strPlacaVeiculo As String, _
                               blnCortesia As Boolean, _
                               strFaixa As String, _
                               strDataHoraEntra As String, _
                               strDataPrevEntrada As String, _
                               strDataPrevSaida As String, _
                               strVrDeposito As String, _
                               strApartamentoId As String, _
                               strTurnoEntradaId As String, _
                               strSequencial As String, _
                               strGrupoPeriodoId As String, _
                               blnDeposito As Boolean, _
                               strOcupantes As String, _
                               strAutoriacao As String, _
                               strIdentPeriodos As String, _
                               strCheque As String, _
                               strCartao As String, _
                               strDocumento As String, _
                               strHospede As String, _
                               lngCARTAOPROMOCIONALID As Long, _
                               strMotivoDesconto As String, _
                               strPercentualDesconto As String, _
                               strDescCortesia As String, strSobreTotal As String, _
                               strReservaId As String, strPacoteId As String, lngCONFIGURACAOID As Long)

  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  
  'Inserir Locação

  
  clsLoc.InserirLocacao blnOcupado, _
                        strPlacaVeiculo, _
                        blnCortesia, _
                        strFaixa, _
                        strDataHoraEntra, _
                        strDataPrevEntrada, _
                        strDataPrevSaida, _
                        strVrDeposito, _
                        strApartamentoId, _
                        strTurnoEntradaId, _
                        strSequencial, _
                        strGrupoPeriodoId, _
                        blnDeposito, _
                        strOcupantes, _
                        strAutoriacao, _
                        strIdentPeriodos, _
                        strCheque, _
                        strCartao, _
                        strDocumento, _
                        strHospede, _
                        lngCARTAOPROMOCIONALID, _
                        strMotivoDesconto, _
                        strPercentualDesconto, _
                        strDescCortesia, strSobreTotal, _
                        strReservaId, strPacoteId, lngCONFIGURACAOID
  '
  InserirLocacao = clsLoc.SelecionarLocacaoPeloSeq(strSequencial)
  Set clsLoc = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarLocacao(lngLocacaoId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  Set clsLoc = New datSisContas.clsLocacao
  
  'Listar Locação
  Set SelecionarLocacao = clsLoc.SelecionarLocacao(lngLocacaoId)
  '
  Set clsLoc = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function ListarLocacao(lngTurnoId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  Set clsLoc = New datSisContas.clsLocacao
  
  'Listar Locação
  Set ListarLocacao = clsLoc.ListarLocacao(lngTurnoId)
  '
  Set clsLoc = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ListarLocacaoPorUnidade(lngTurnoId As Long, _
                                        strSql As String) As ADODB.Recordset
  On Error GoTo trata
  Dim clsLoc As datSisContas.clsLocacao
  Set clsLoc = New datSisContas.clsLocacao
  
  'Listar Locação
  Set ListarLocacaoPorUnidade = clsLoc.ListarLocacaoPorUnidade(lngTurnoId, _
                                                               strSql)
  '
  Set clsLoc = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub CapturaAptoSuite(ByVal strNumeroSuiteApto As String, _
                                  strNomeSuiteApto As String, _
                                  lngLocacaoId As Long)
  '
  On Error GoTo trata
    
  Dim objRs As ADODB.Recordset
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
    
  If lngLocacaoId > 0 Then
    Set objRs = clsLoc.SelecionarLocacao(lngLocacaoId)
  Else
    Set objRs = clsLoc.CapturaAptoSuite(strNumeroSuiteApto)
  End If
  '
  lngLocacaoId = IIf(Not IsNumeric(objRs.Fields("PKID")), 0, objRs.Fields("PKID"))
  strNomeSuiteApto = objRs.Fields("NOME") & ""
  '
  objRs.Close
  Set objRs = Nothing
  Set clsLoc = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub GravarMovAposFecha(lngLocacaoId As Long)
  On Error GoTo trata
  Dim objRs   As ADODB.Recordset
  Dim strSql  As String
  Dim clsLoc As datSisContas.clsLocacao
  '
  Set clsLoc = New datSisContas.clsLocacao
  '
  Set objRs = clsLoc.CapturaLocacao(lngLocacaoId)
  If IIf(IsNull(objRs.Fields("FECHAMENTO").Value), False, objRs.Fields("FECHAMENTO").Value) Then
    clsLoc.GravarMovAposFecha (lngLocacaoId)
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set clsLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

