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
  '
  Dim strSql As String
      
  strSql = "UPDATE LOCACAO SET " & _
    "RECEBIMENTO = " & Formata_Dados(blnRecebimento, tpDados_Boolean) & _
    ", TURNORECEBEID = " & Formata_Dados(lngTurnoRecebeId, tpDados_Longo) & _
    " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  If blnAtualizarTurnoRecCC = True Then
    'Atualizar também Turno da conta corrente
    strSql = "UPDATE CONTACORRENTE SET " & _
      "TURNOCCID = " & Formata_Dados(lngTurnoRecebeId, tpDados_Longo) & _
      " WHERE LOCACAOID = " & Formata_Dados(lngLocacaoId, tpDados_Longo)
  
    RunSP strSql
    
  End If
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocRecCC]", _
            Err.Description
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
  '
  Dim strSql As String
      
  strSql = "UPDATE LOCACAO SET " & _
    " TURNORECEBEID = " & Formata_Dados(lngTurnoRecebeId, tpDados_Longo, tpNulo_Aceita) & _
    ", GARCOMID = " & Formata_Dados(strGarcomId, tpDados_Longo, tpNulo_Aceita) & _
    ", DTHORARECEB = " & Formata_Dados(strDtHoraReceb, tpDados_DataHora, tpNulo_Aceita) & _
    ", RESPONSAVEL = " & Formata_Dados(strResponsavel, tpDados_Texto, tpNulo_Aceita) & _
    ", COBRANCA = " & Formata_Dados(blnCobranca, tpDados_Boolean, tpNulo_Aceita) & _
    ", PGTOESPECIE = " & Formata_Dados(strPgtoEspecie, tpDados_Moeda, tpNulo_Aceita) & _
    ", VRCALCGORJETA = " & Formata_Dados(strGorjeta, tpDados_Moeda, tpNulo_Aceita) & _
    ", PGTOCARTAO = " & Formata_Dados(strPgtoCartao, tpDados_Moeda, tpNulo_Aceita) & ", PGTOCARTAODEB = " & Formata_Dados(strPgtoCartaoDebito, tpDados_Moeda, tpNulo_Aceita) & _
    ", CARTAOID = " & Formata_Dados(strCartaoId, tpDados_Longo, tpNulo_Aceita) & ", CARTAODEBITOID = " & Formata_Dados(strCartaoDebId, tpDados_Longo, tpNulo_Aceita) & _
    ", LOTE = " & Formata_Dados(strLote, tpDados_Texto, tpNulo_Aceita) & _
    ", PGTOCHEQUE = " & Formata_Dados(strPgtoCheque, tpDados_Moeda, tpNulo_Aceita) & _
    ", CPF = " & Formata_Dados(strCPF, tpDados_Texto, tpNulo_Aceita) & _
    ", NROCHEQUE = " & Formata_Dados(strNroCheque, tpDados_Texto, tpNulo_Aceita) & _
    ", BANCOID = " & Formata_Dados(strBancoId, tpDados_Longo, tpNulo_Aceita) & _
    ", AGENCIA = " & Formata_Dados(strAgencia, tpDados_Texto, tpNulo_Aceita) & _
    ", CONTA = " & Formata_Dados(strConta, tpDados_Texto, tpNulo_Aceita) & _
    ", PGTOPENHOR = " & Formata_Dados(strPgtoPenhor, tpDados_Moeda, tpNulo_Aceita) & _
    ", CLIENTE = " & Formata_Dados(strCliente, tpDados_Texto, tpNulo_Aceita) & _
    ", DOCUMENTOPENHOR = " & Formata_Dados(strDocumentoPenhor, tpDados_Texto, tpNulo_Aceita) & _
    ", DESCOBJETO = " & Formata_Dados(strObjeto, tpDados_Texto, tpNulo_Aceita) & _
    ", VRCALCTOTALPAGO = " & Formata_Dados(strVrCalcTotalPago, tpDados_Moeda, tpNulo_Aceita) & _
    ", VRCALCTROCO = " & Formata_Dados(strVrCalcTroco, tpDados_Moeda, tpNulo_Aceita) & _
    ", RECEBIMENTO = " & Formata_Dados(blnRecebimento, tpDados_Boolean, tpNulo_Aceita) & _
    " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocRecebimento]", _
            Err.Description
End Sub
Public Sub AlterarLocLimpeza(lngLocacaoId As Long, _
                             blnLimpeza As Boolean, _
                             blnOcupado As Boolean, _
                             lngCamareiraLimpezaId As Long)

  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE LOCACAO SET " & _
    " LIMPO = " & Formata_Dados(blnLimpeza, tpDados_Boolean, tpNulo_Aceita) & _
    ", OCUPADO = " & Formata_Dados(blnOcupado, tpDados_Boolean, tpNulo_Aceita) & _
    ", CAMAREIRALIMPEZAID = " & Formata_Dados(lngCamareiraLimpezaId, tpDados_Longo, tpNulo_Aceita) & _
    " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocLimpeza]", _
            Err.Description
End Sub
Public Sub AlterarLocEncerrar(lngLocacaoId As Long, _
                              blnOcupado As Boolean)

  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE LOCACAO SET " & _
    " OCUPADO = " & Formata_Dados(blnOcupado, tpDados_Boolean, tpNulo_Aceita) & _
    " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocEncerrar]", _
            Err.Description
End Sub

Public Sub AlterarLocPeriodoSomaVenc(lngLocacaoId As Long)

  On Error GoTo trata
  '
  Dim strSql As String
        
  strSql = "UPDATE LOCACAO SET " & _
    " PeriodoSomaVenc = case when PeriodoSomaVenc is null then 1 else (PeriodoSomaVenc + 1) end " & _
    " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocPeriodoSomaVenc]", _
            Err.Description
End Sub

Public Sub AlterarLocLiberacao(lngLocacaoId As Long, _
                               blnLimpo As Boolean, _
                               blnLiberado As Boolean, _
                               lngCamareiraLiberaId As Long)

  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE LOCACAO SET " & _
    " LIMPO = " & Formata_Dados(blnLimpo, tpDados_Boolean, tpNulo_Aceita) & _
    ", LIBERADO = " & Formata_Dados(blnLiberado, tpDados_Boolean, tpNulo_Aceita) & _
    ", CAMAREIRALIBERAID = " & Formata_Dados(IIf(lngCamareiraLiberaId = 0, "", lngCamareiraLiberaId), tpDados_Longo) & _
    " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocLiberacao]", _
            Err.Description
End Sub
Public Sub AlterarLocMovAposFecha(lngLocacaoId As Long, _
                                  blnMOVAPOSFECHA As Boolean, _
                                  blnMOVAPOSFECHAOK As Boolean)

  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE LOCACAO SET " & _
    " MOVAPOSFECHA = " & Formata_Dados(blnMOVAPOSFECHA, tpDados_Boolean, tpNulo_Aceita) & "," & _
    " MOVAPOSFECHAOK = " & Formata_Dados(blnMOVAPOSFECHAOK, tpDados_Boolean, tpNulo_Aceita) & _
    " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocMovAposFecha]", _
            Err.Description
End Sub

Public Sub AlterarLocOcupado(lngLocacaoId As Long, _
                             blnOcupado As Boolean)

  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE LOCACAO SET " & _
    " OCUPADO = " & Formata_Dados(blnOcupado, tpDados_Boolean, tpNulo_Aceita) & _
    " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocOcupado]", _
            Err.Description
End Sub
Public Sub AlterarLocSaida(lngLocacaoId As Long, _
                           strDtHoraSaida As String, _
                           lngTurnoSaidaId As Long, _
                           blnSaiu As Boolean)

  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE LOCACAO SET " & _
    " DTHORASAIDA = " & Formata_Dados(strDtHoraSaida, tpDados_DataHora, tpNulo_Aceita) & _
    ", TURNOSAIDAID = " & Formata_Dados(lngTurnoSaidaId, tpDados_Longo, tpNulo_Aceita) & _
    ", SAIU = " & Formata_Dados(blnSaiu, tpDados_Boolean, tpNulo_Aceita) & _
    " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocSaida]", _
            Err.Description
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
  '
  Dim strSql As String

  strSql = "UPDATE LOCACAO SET " & _
    " PLACAVEICULO = " & Formata_Dados(strPlacaVeiculo, tpDados_Texto, tpNulo_Aceita) & _
    ", CORTESIA = " & Formata_Dados(blnCortesia, tpDados_Boolean, tpNulo_Aceita) & _
    ", VRFXGRPERIODO = " & Formata_Dados(strFaixa, tpDados_Texto, tpNulo_Aceita) & _
    ", DATAHORAENTRA = " & Formata_Dados(strDataHoraEntra, tpDados_DataHora, tpNulo_Aceita) & _
    ", DTPREVENTRADA = " & Formata_Dados(strDataPrevEntrada, tpDados_DataHora, tpNulo_Aceita) & _
    ", DTPREVSAIDA = " & Formata_Dados(strDataPrevSaida, tpDados_DataHora, tpNulo_Aceita) & _
    ", VRDEPOSITO = " & Formata_Dados(strVrDeposito, tpDados_Moeda, tpNulo_Aceita) & _
    ", APARTAMENTOID = " & Formata_Dados(strApartamentoId, tpDados_Longo, tpNulo_Aceita) & _
    ", TURNOENTRADAID = " & Formata_Dados(strTurnoEntradaId, tpDados_Longo, tpNulo_Aceita) & _
    ", GRUPOPERIODOID = " & Formata_Dados(strGrupoPeriodoId, tpDados_Longo, tpNulo_Aceita) & _
    ", DEPOSITO = " & Formata_Dados(blnDeposito, tpDados_Boolean, tpNulo_Aceita) & _
    ", OCUPANTES = " & Formata_Dados(strOcupantes, tpDados_Longo, tpNulo_Aceita) & _
    ", AUTORIZACAO = " & Formata_Dados(strAutoriacao, tpDados_Texto, tpNulo_Aceita) & _
    ", IDENTIFPERIODO = " & Formata_Dados(strIdentPeriodos, tpDados_Texto, tpNulo_Aceita) & _
    ", CHEQUE = " & Formata_Dados(strCheque, tpDados_Texto, tpNulo_Aceita) & _
    ", CARTAO = " & Formata_Dados(strCartao, tpDados_Texto, tpNulo_Aceita) & _
    ", DOCUMENTO = " & Formata_Dados(strDocumento, tpDados_Texto, tpNulo_Aceita) & _
    ", HOSPEDE = " & Formata_Dados(strHospede, tpDados_Texto, tpNulo_Aceita) & _
    ", CARTAOPROMOCIONALID = " & Formata_Dados(IIf(lngCARTAOPROMOCIONALID = 0, "", lngCARTAOPROMOCIONALID), tpDados_Longo, tpNulo_Aceita) & _
    ", MOTIVODESC = " & Formata_Dados(strMotivoDesconto, tpDados_Texto, tpNulo_Aceita) & _
    ", PERCDESCONTO = " & Formata_Dados(strPercentualDesconto, tpDados_Moeda, tpNulo_Aceita) & _
    ", DESCCORTESIA = " & Formata_Dados(strDescCortesia, tpDados_Texto, tpNulo_Aceita)
    
  strSql = strSql & ", SOBRETOTAL = " & Formata_Dados(strSobreTotal, tpDados_Texto, tpNulo_Aceita) & _
    ", RESERVAID = " & Formata_Dados(strReservaId, tpDados_Longo, tpNulo_Aceita) & _
    ", PACOTEID = " & Formata_Dados(strPacoteId, tpDados_Longo, tpNulo_Aceita) & _
    ", PERIODOSOMAVENC = " & Formata_Dados(1, tpDados_Longo, tpNulo_Aceita)
    
  strSql = strSql & ", FECHAMENTO = " & Formata_Dados(False, tpDados_Boolean) & _
    ", RECEBIMENTO = " & Formata_Dados(False, tpDados_Boolean) & _
    ", LIBERADO = " & Formata_Dados(False, tpDados_Boolean) & _
    ", SAIU = " & Formata_Dados(False, tpDados_Boolean) & _
    ", LIMPO = " & Formata_Dados(False, tpDados_Boolean) & _
    ", MOVAPOSFECHA = " & Formata_Dados(False, tpDados_Boolean) & _
    ", MOVAPOSFECHAOK = " & Formata_Dados(False, tpDados_Boolean) & _
    " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  'Atualizar apartamento
  strSql = "UPDATE APARTAMENTO SET " & _
    "OCUPADOPARC = " & Formata_Dados(False, tpDados_Longo) & _
    " WHERE PKID = " & Formata_Dados(strApartamentoId, tpDados_Longo)
  RunSP strSql
  
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocEntradaParc]", _
            Err.Description
End Sub

Public Sub AlterarLocLiberarCheque(lngLocacaoId As Long, _
                                   strCPF As String)

  On Error GoTo trata
  '
  Dim strSql As String
      
  strSql = "UPDATE LOCACAO SET " & _
    " CPF = " & Formata_Dados(strCPF, tpDados_Texto, tpNulo_Aceita) & _
    " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocLiberarCheque]", _
            Err.Description
End Sub

Public Sub AlterarLocVrEmprPrincAssoc(lngLocacaoId As Long, _
                                      strValor As String)

  On Error GoTo trata
  '
  Dim strSql As String
      
  strSql = "UPDATE LOCACAO SET " & _
    " VRCALCTOTALEMPRESASASSOC = " & Formata_Dados(strValor, tpDados_Moeda)
  strSql = strSql & " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocVrEmprPrincAssoc]", _
            Err.Description
End Sub

Public Sub AlterarLocFechamento(lngLocacaoId As Long, strVrDesconto As String, strPercDesconto As String, strMotivoDesconto As String, strQtdFecha As String, blnFechamento As Boolean, strDtHoraFecha As String, blnMOVAPOSFECHAOK As Boolean, _
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
  '
  Dim strSql As String
      
  strSql = "UPDATE LOCACAO SET " & _
    " VRDESCONTO = " & Formata_Dados(strVrDesconto, tpDados_Moeda, tpNulo_Aceita) & ", PERCDESCONTO = " & Formata_Dados(strPercDesconto, tpDados_Moeda, tpNulo_Aceita) & ", MOTIVODESC = " & Formata_Dados(strMotivoDesconto, tpDados_Texto, tpNulo_Aceita) & ", QTDFECHA = " & Formata_Dados(strQtdFecha, tpDados_Longo, tpNulo_Aceita) & ", FECHAMENTO = " & Formata_Dados(blnFechamento, tpDados_Boolean, tpNulo_Aceita) & _
    ", DTHORAFECHA = " & Formata_Dados(strDtHoraFecha, tpDados_DataHora, tpNulo_Aceita) & ", MOVAPOSFECHAOK = " & Formata_Dados(blnMOVAPOSFECHAOK, tpDados_Boolean, tpNulo_Aceita) & _
    ", VRHREXCEDENTE = " & Formata_Dados(strVrHrExced, tpDados_Moeda, tpNulo_Aceita) & _
    ", QTDHREXCEDENTE = " & Formata_Dados(strQtdHrExced, tpDados_Longo, tpNulo_Aceita) & _
    ", QTDPERIODOSEXCEDENTE = " & Formata_Dados(strQtdPerExced, tpDados_Longo, tpNulo_Aceita) & _
    ", VRHRPERIODOSEXCEDENTE = " & Formata_Dados(strVrPerExced, tpDados_Moeda, tpNulo_Aceita) & _
    ", QTDTOTALPERIODOS = " & Formata_Dados(strQtdTotalPer, tpDados_Longo, tpNulo_Aceita) & _
    ", VRCALCDIARIA = " & Formata_Dados(strVrCalcDiaria, tpDados_Moeda, tpNulo_Aceita) & ", VRCALCCONSUMO = " & Formata_Dados(strVrCalcConsumo, tpDados_Moeda, tpNulo_Aceita) & ", VRCALCTELEFONEMA = " & Formata_Dados(strVrCalcTelefonema, tpDados_Moeda, tpNulo_Aceita) & ", VRCALCEXCEDPES = " & Formata_Dados(strVrCalcExcedPes, tpDados_Moeda, tpNulo_Aceita) & _
    ", VRCALCEXTRA = " & Formata_Dados(strVrCalcExtra, tpDados_Moeda, tpNulo_Aceita) & ", VRCALCDESCONTO = " & Formata_Dados(strVrCalcDesconto, tpDados_Moeda, tpNulo_Aceita) & ", VRCALCTOTAL = " & Formata_Dados(strVrCalcTotal, tpDados_Moeda, tpNulo_Aceita) & ", CARTAOPROMOCIONALID = " & Formata_Dados(IIf(lngCARTAOPROMOCIONALID = 0, "", lngCARTAOPROMOCIONALID), tpDados_Longo, tpNulo_Aceita) & ", CORTESIA = " & Formata_Dados(blnCortesia, tpDados_Boolean, tpNulo_Aceita) & _
    ", AUTORIZACAO = " & Formata_Dados(strAutorizacao, tpDados_Texto, tpNulo_Aceita) & ", DESCCORTESIA = " & Formata_Dados(strDescCortesia, tpDados_Texto, tpNulo_Aceita) & ", SOBRETOTAL = " & Formata_Dados(strSobreTotal, tpDados_Texto, tpNulo_Aceita) & ", VRCALCDESC = " & Formata_Dados(strVrCalcDesc, tpDados_Moeda, tpNulo_Aceita) & _
    ", VRCALCCORTDEZ = " & Formata_Dados(strVrCalcCortDez, tpDados_Moeda, tpNulo_Aceita) & _
    ", VRCALCCORTDIA = " & Formata_Dados(strVrCalcCortDia, tpDados_Moeda, tpNulo_Aceita) & _
    ", VRCALCCORTDIATOT = " & Formata_Dados(strVrCalcCortDiaTot, tpDados_Moeda, tpNulo_Aceita) & _
    ", VRCALCCORTTOT = " & Formata_Dados(strVrCalcCortTot, tpDados_Moeda, tpNulo_Aceita) & _
    ", VRCALCCORTCONS = " & Formata_Dados(strVrCalcCortCons, tpDados_Moeda, tpNulo_Aceita) & _
    ", VRCALCCC = " & Formata_Dados(strVrTotalCC, tpDados_Moeda) & _
    ", VRCALCEXTRAEMPR = " & Formata_Dados(strVrExtraEmpr, tpDados_Moeda) & _
    ", VRCALCCONSUMOEMPR = " & Formata_Dados(strVrConsumoEmpr, tpDados_Moeda) & _
    ", VRCALCTELEFONEMAEMPR = " & Formata_Dados(strVrTelefoneEmpr, tpDados_Moeda) & _
    ", QTDTOTALPERIODOSEMPR = " & Formata_Dados(strQtdPerEmpr, tpDados_Longo) & _
    ", VRCALCDIARIAEMPR = " & Formata_Dados(strVrDiariaEmpr, tpDados_Moeda) & _
    ", VRCALCEXCEDPESEMPR = " & Formata_Dados(strVrExcedPesEmpr, tpDados_Moeda) & _
    ", VRCALCTOTALEMPRESA = " & Formata_Dados(strVrTotalEmpresa, tpDados_Moeda)
  strSql = strSql & _
    ", VRCALCTAXAEMPR = " & Formata_Dados(strVrTaxaEmpr, tpDados_Moeda) & _
    ", VRCALCTAXA = " & Formata_Dados(strVrTaxa, tpDados_Moeda) & _
    ", COBRARTAXA = " & Formata_Dados(blnCobrarTaxa, tpDados_Boolean)
  If CInt(strQtdFecha) = 1 Then
    strSql = strSql & ", DTHORAPRIFECHA = " & Formata_Dados(strDtHoraPriFecha, tpDados_DataHora, tpNulo_Aceita) & ", TURNOFECHAID = " & Formata_Dados(lngTurnoFechaId, tpDados_Longo, tpNulo_Aceita)
  End If
  strSql = strSql & " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocFechamento]", _
            Err.Description
End Sub

Public Sub ExcluirLocacao(lngLocacaoId As Long, _
                          gbTrabSuiteAptoLimpo As Boolean)

  On Error GoTo trata
  '
  Dim strSql            As String
  Dim lngAPARTAMENTOID  As Long
  Dim objRs             As ADODB.Recordset
  '
  'Obter lngAPARTAMENTOID da locação que será excluida
  lngAPARTAMENTOID = 0
  strSql = "SELECT APARTAMENTOID FROM LOCACAO WHERE LOCACAO.PKID = " & _
          Formata_Dados(lngLocacaoId, tpDados_Longo)
  
  Set objRs = RunSPReturnRS(strSql)
  '
  If Not objRs.EOF Then
    lngAPARTAMENTOID = objRs.Fields("APARTAMENTOID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  strSql = "DELETE FROM TAB_RESPLOCACAO " & _
    " WHERE LOCACAOID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  strSql = "DELETE FROM TAB_FICHACLIELOC " & _
    " WHERE LOCACAOID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  strSql = "DELETE FROM VIAGEM " & _
    " WHERE LOCACAOID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  strSql = "DELETE FROM LEMBRETE " & _
    " WHERE ALOCACAOID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  strSql = "DELETE FROM CORTESIA " & _
    " WHERE LOCACAOID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  strSql = "DELETE FROM ENTRADASAIDA " & _
    " WHERE LOCACAOID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  strSql = "DELETE FROM LOCACAO " & _
    " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  If gbTrabSuiteAptoLimpo Then
    'Entrar com unidade em estado de limpeza
    strSql = "UPDATE APARTAMENTO SET LIMPEZA = " & Formata_Dados(True, tpDados_Boolean) & _
      " WHERE APARTAMENTO.PKID = " & _
          Formata_Dados(lngAPARTAMENTOID, tpDados_Longo) & _
      " AND EXCLUIDO = " & Formata_Dados(False, tpDados_Boolean)
    RunSP strSql
  End If
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.ExcluirLocacao]", _
            Err.Description
End Sub

Public Sub AlterarLocTransf(lngLocacaoId As Long, _
                            strApartamentoId As String, _
                            strGrupoPeriodoId As String, _
                            strDataPrevSaida As String, _
                            strNumeroAptoAnterior As String, _
                            gbTrabSuiteAptoLimpo As Boolean)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE LOCACAO SET " & _
    " APARTAMENTOID = " & Formata_Dados(strApartamentoId, tpDados_Longo, tpNulo_Aceita) & _
    ", GRUPOPERIODOID = " & Formata_Dados(strGrupoPeriodoId, tpDados_Longo, tpNulo_Aceita) & _
    ", DTPREVSAIDA = " & Formata_Dados(strDataPrevSaida, tpDados_DataHora, tpNulo_Aceita) & _
    " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  '
  If strNumeroAptoAnterior <> "" And gbTrabSuiteAptoLimpo Then
    'Entrar com unidade em estado de limpeza
    strSql = "UPDATE APARTAMENTO SET LIMPEZA = " & Formata_Dados(True, tpDados_Boolean) & _
      " WHERE APARTAMENTO.NUMERO = " & Formata_Dados(strNumeroAptoAnterior, tpDados_Texto, tpNulo_Aceita) & _
      " AND EXCLUIDO = " & Formata_Dados(False, tpDados_Boolean)
    RunSP strSql
  End If
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocTransf]", _
            Err.Description
End Sub
Public Sub AlterarLocDtHoraCalcFecha(lngLocacaoId As Long, _
                                     strDtHoraCalcFecha As String)

  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE LOCACAO SET " & _
    " DTHORACALCFECHA = " & Formata_Dados(strDtHoraCalcFecha, tpDados_DataHora, tpNulo_Aceita) & _
    " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocDtHoraCalcFecha]", _
            Err.Description
End Sub

Public Sub AlterarLocCortesia(lngLocacaoId As Long, _
                              blnCortesia As Boolean)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE LOCACAO SET " & _
    " CORTESIA = " & Formata_Dados(blnCortesia, tpDados_Boolean, tpNulo_Aceita) & " " & _
    " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocCortesia]", _
            Err.Description
End Sub

Public Sub AlterarLocDeposito(lngLocacaoId As Long, _
                              strVrDeposito As String, _
                              blnDeposito As Boolean, _
                              strCheque As String, _
                              strCartao As String, _
                              strDocumento As String, _
                              strHospede As String)


  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE LOCACAO SET " & _
    " VRDEPOSITO = " & Formata_Dados(strVrDeposito, tpDados_Moeda, tpNulo_Aceita) & _
    ", DEPOSITO = " & Formata_Dados(blnDeposito, tpDados_Boolean, tpNulo_Aceita) & _
    ", CHEQUE = " & Formata_Dados(strCheque, tpDados_Texto, tpNulo_Aceita) & _
    ", CARTAO = " & Formata_Dados(strCartao, tpDados_Texto, tpNulo_Aceita) & _
    ", DOCUMENTO = " & Formata_Dados(strDocumento, tpDados_Texto, tpNulo_Aceita) & _
    ", HOSPEDE = " & Formata_Dados(strHospede, tpDados_Texto, tpNulo_Aceita) & _
    " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocDeposito]", _
            Err.Description
End Sub

Public Sub AlterarLocEntradaPes(lngLocacaoId As Long, _
                                strQuantidadePes As String)


  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE LOCACAO SET " & _
    " OCUPANTES = OCUPANTES + " & Formata_Dados(strQuantidadePes, tpDados_Longo) & _
    " WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  'Insere na tabela ENTRADASAIDA
  
  strSql = "INSERT INTO ENTRADASAIDA (LOCACAOID, QTDPESSOAS, INDENTRADASAIDA, DTHORA) VALUES( " & _
    Formata_Dados(lngLocacaoId, tpDados_Longo) & _
    ", " & Formata_Dados(IIf(CInt(strQuantidadePes) < 0, CInt(strQuantidadePes) * (-1), CInt(strQuantidadePes)), tpDados_Longo) & _
    ", " & Formata_Dados(IIf(CInt(strQuantidadePes) < 0, "S", "E"), tpDados_Texto) & _
    ", " & Formata_Dados(Format(Now, "DD/MM/YYYY hh:mm:ss"), tpDados_DataHora) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.AlterarLocEntradaPes]", _
            Err.Description
End Sub

Public Sub InserirLocacao(blnOcupado As Boolean, _
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
  '
  Dim strSql As String
  
  strSql = "INSERT INTO LOCACAO(OCUPADO, PLACAVEICULO, CORTESIA, VRFXGRPERIODO, DATAHORAENTRA, DTPREVENTRADA, DTPREVSAIDA, VRDEPOSITO, APARTAMENTOID, TURNOENTRADAID, SEQUENCIAL, GRUPOPERIODOID, DEPOSITO, OCUPANTES, OCUPANTESINIC, AUTORIZACAO, IDENTIFPERIODO, CHEQUE, CARTAO, DOCUMENTO, HOSPEDE, CARTAOPROMOCIONALID, MOTIVODESC, PERCDESCONTO, DESCCORTESIA, SOBRETOTAL, RESERVAID, PACOTEID, PERIODOSOMAVENC, FECHAMENTO, RECEBIMENTO, LIBERADO, SAIU, LIMPO, MOVAPOSFECHA, MOVAPOSFECHAOK, CONFIGURACAOID) VALUES (" & _
    Formata_Dados(blnOcupado, tpDados_Boolean, tpNulo_Aceita) & _
    ", " & Formata_Dados(strPlacaVeiculo, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(blnCortesia, tpDados_Boolean, tpNulo_Aceita) & _
    ", " & Formata_Dados(strFaixa, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strDataHoraEntra, tpDados_DataHora, tpNulo_Aceita) & _
    ", " & Formata_Dados(strDataPrevEntrada, tpDados_DataHora, tpNulo_Aceita) & _
    ", " & Formata_Dados(strDataPrevSaida, tpDados_DataHora, tpNulo_Aceita) & _
    ", " & Formata_Dados(strVrDeposito, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(strApartamentoId, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strTurnoEntradaId, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strSequencial, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strGrupoPeriodoId, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(blnDeposito, tpDados_Boolean, tpNulo_Aceita) & _
    ", " & Formata_Dados(strOcupantes, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strOcupantes, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strAutoriacao, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strIdentPeriodos, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strCheque, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strCartao, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strDocumento, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strHospede, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(IIf(lngCARTAOPROMOCIONALID = 0, "", lngCARTAOPROMOCIONALID), tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strMotivoDesconto, tpDados_Texto, tpNulo_Aceita) & ", " & Formata_Dados(strPercentualDesconto, tpDados_Moeda, tpNulo_Aceita) & ", " & Formata_Dados(strDescCortesia, tpDados_Texto, tpNulo_Aceita) & ", " & Formata_Dados(strSobreTotal, tpDados_Texto, tpNulo_Aceita) & ", " & Formata_Dados(strReservaId, tpDados_Longo, tpNulo_Aceita) & ", " & Formata_Dados(strPacoteId, tpDados_Longo, tpNulo_Aceita) & ", " & Formata_Dados(1, tpDados_Longo, tpNulo_Aceita)
    
  strSql = strSql & ", " & Formata_Dados(False, tpDados_Boolean) & _
    ", " & Formata_Dados(False, tpDados_Boolean) & _
    ", " & Formata_Dados(False, tpDados_Boolean) & _
    ", " & Formata_Dados(False, tpDados_Boolean) & _
    ", " & Formata_Dados(False, tpDados_Boolean) & _
    ", " & Formata_Dados(False, tpDados_Boolean) & _
    ", " & Formata_Dados(False, tpDados_Boolean) & _
    ", " & Formata_Dados(lngCONFIGURACAOID, tpDados_Longo) & _
    ")"
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.InserirLocacao]", _
            Err.Description
End Sub

Public Function SelecionarLocacaoPeloSeq(strSequencial As String) As Long
  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  strSql = "SELECT TOP 1 LOCACAO.PKID " & _
    " FROM LOCACAO " & _
    " WHERE LOCACAO.SEQUENCIAL = " & _
    Formata_Dados(strSequencial, tpDados_Longo, tpNulo_NaoAceita) & _
    " ORDER BY LOCACAO.PKID DESC;"
  
  Set objRs = RunSPReturnRS(strSql)
  '
  If objRs.EOF Then
    SelecionarLocacaoPeloSeq = 0
  Else
    SelecionarLocacaoPeloSeq = objRs.Fields("PKID").Value
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.SelecionarLocacao]", _
            Err.Description
End Function

Public Function SelecionarLocacao(lngLocacaoId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT LOCACAO.*, EMPRESA.NOME AS NOMEEMPRESA, GRUPOPERIODO.PERIODO, FXGRUPOPERIODO.VALOR, FXGRUPOPERIODO.DESCRICAO AS DESCGRUPOPERIODO, APARTAMENTO.NUMERO, GRUPO.NOME, CARTAOPROMOCIONAL.NUMERO AS NUMERO_CARTAO, GARCOM.NOME AS NOME_GARCOM, CARTAO.NOME AS NOME_CARTAO, BANCO.NOME AS NOME_BANCO, PACOTE.DESCRICAO AS DESCPACOTE, CARTAODEBITO.NOME AS NOME_CARTAO_DEBITO, " & _
    " VR_DEPOSITO_NOVO = isnull((SELECT SUM(VALOR) FROM CONTACORRENTE WHERE CONTACORRENTE.LOCACAOID = LOCACAO.PKID AND CONTACORRENTE.STATUSLANCAMENTO = " & Formata_Dados("DP", tpDados_Texto) & "),0) " & _
    " FROM LOCACAO " & _
    " LEFT JOIN GRUPOPERIODO ON LOCACAO.GRUPOPERIODOID = GRUPOPERIODO.PKID " & _
    " LEFT JOIN FXGRUPOPERIODO ON FXGRUPOPERIODO.PKID = GRUPOPERIODO.FXGRUPOPERIODOID " & _
    " LEFT JOIN APARTAMENTO ON LOCACAO.APARTAMENTOID = APARTAMENTO.PKID " & _
    " LEFT JOIN GRUPO ON GRUPO.PKID = APARTAMENTO.GRUPOID " & _
    " LEFT JOIN CARTAOPROMOCIONAL ON CARTAOPROMOCIONAL.PKID = LOCACAO.CARTAOPROMOCIONALID " & _
    " LEFT JOIN GARCOM ON GARCOM.PKID = LOCACAO.GARCOMID " & _
    " LEFT JOIN CARTAO ON CARTAO.PKID = LOCACAO.CARTAOID " & _
    " LEFT JOIN BANCO ON BANCO.PKID = LOCACAO.BANCOID " & _
    " LEFT JOIN PACOTE ON PACOTE.PKID = LOCACAO.PACOTEID " & _
    " LEFT JOIN VIAGEM ON LOCACAO.PKID = VIAGEM.LOCACAOID " & _
    " LEFT JOIN EMPRESA ON EMPRESA.PKID = VIAGEM.EMPRESAID " & _
    " LEFT JOIN CARTAODEBITO ON CARTAODEBITO.PKID = LOCACAO.CARTAODEBITOID " & _
    " WHERE LOCACAO.PKID = " & _
    Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_NaoAceita)
  
  Set SelecionarLocacao = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.SelecionarLocacao]", _
            Err.Description
End Function

Public Function ListarLocacao(lngTurnoId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT LOCACAO.OCUPADO, LOCACAO.PKID, APARTAMENTO.NUMERO, GRUPO.NOME " & _
    " FROM ((LOCACAO " & _
    " LEFT JOIN APARTAMENTO ON LOCACAO.APARTAMENTOID = APARTAMENTO.PKID) " & _
    " LEFT JOIN GRUPO ON GRUPO.PKID = APARTAMENTO.GRUPOID) " & _
    " LEFT JOIN TURNO ON TURNO.PKID = LOCACAO.TURNOENTRADAID " & _
    " WHERE TURNO.PKID = " & _
    Formata_Dados(lngTurnoId, tpDados_Longo, tpNulo_NaoAceita) & " ORDER BY LOCACAO.PKID DESC"
  
  Set ListarLocacao = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.ListarLocacao]", _
            Err.Description
End Function
Public Function ListarLocacaoPorUnidade(lngTurnoId As Long, _
                                        strWhere As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT LOCACAO.OCUPADO, LOCACAO.PKID, APARTAMENTO.NUMERO, GRUPO.NOME " & _
    " FROM ((LOCACAO " & _
    " LEFT JOIN APARTAMENTO ON LOCACAO.APARTAMENTOID = APARTAMENTO.PKID) " & _
    " LEFT JOIN GRUPO ON GRUPO.PKID = APARTAMENTO.GRUPOID) " & _
    " LEFT JOIN TURNO ON TURNO.PKID = LOCACAO.TURNOENTRADAID " & _
    " WHERE TURNO.PKID = " & _
    Formata_Dados(lngTurnoId, tpDados_Longo, tpNulo_NaoAceita)
  If strWhere & "" <> "" Then
    strSql = strSql & " AND " & strWhere
  End If
  strSql = strSql & " ORDER BY LOCACAO.PKID DESC"
  
  Set ListarLocacaoPorUnidade = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.ListarLocacaoPorUnidade]", _
            Err.Description
End Function

Public Sub GravarMovAposFecha(ByVal lngLocacaoId As Long)
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "UPDATE LOCACAO SET MOVAPOSFECHA = " & Formata_Dados(True, tpDados_Boolean) & _
      " , MOVAPOSFECHAOK = " & Formata_Dados(True, tpDados_Boolean) & _
      " WHERE LOCACAO.PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.GravarMovAposFecha]", _
            Err.Description
End Sub

Public Function CapturaAptoSuite(ByVal strNumeroSuiteApto As String) As ADODB.Recordset
  '
  On Error GoTo trata
  
  Dim sSql    As String
  '
  sSql = "Select Locacao.PKID, Apartamento.NUMERO, Grupo.NOME FROM GRUPO INNER JOIN (APARTAMENTO INNER JOIN LOCACAO ON APARTAMENTO.PKID = LOCACAO.APARTAMENTOID) ON GRUPO.PKID = APARTAMENTO.GRUPOID Where Numero = " & Formata_Dados(strNumeroSuiteApto, tpDados_Texto, tpNulo_NaoAceita) & " And Locacao.Ocupado = " & Formata_Dados(True, tpDados_Boolean)
  Set CapturaAptoSuite = RunSPReturnRS(sSql)
  '
  If CapturaAptoSuite.EOF Then
    CapturaAptoSuite.Close
    Set CapturaAptoSuite = Nothing
    Err.Raise lngCteErroData, "[" & strClassName & ".clsLocacao.CapturaAptoSuite]", "Apartamento/Suite não possui Locação Associada."
  End If
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.CapturaAptoSuite]", _
            Err.Description
End Function

Public Function CapturaLocacao(ByVal lngLocacaoId As Long) As ADODB.Recordset
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "SELECT * FROM LOCACAO WHERE PKID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_NaoAceita)
  Set CapturaLocacao = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLocacao.CapturaLocacao]", _
            Err.Description
End Function



