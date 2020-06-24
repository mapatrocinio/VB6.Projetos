VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConfiguracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''''Public Sub CadastrarConfiguracaoMesa(lngCONFIGURACAOID As Long, _
''''                                      strVALORINICIAL As String, strVALORFINAL As String, strVALORSEPARACAO As String, _
''''                                      strDATAINICIAL As String, _
''''                                      strDATAFINAL As String, _
''''                                      strDATAMASCARA As String, _
''''                                      strHORAINICIAL As String, _
''''                                      strHORAFINAL As String, _
''''                                      strHORAMASCARA As String, _
''''                                      strDURACAOINICIAL As String, _
''''                                      strDURACAOFINAL As String, _
''''                                      strDURACAOMASCARA As String, _
''''                                      strNUMEROINICIAL As String, _
''''                                      strNUMEROFINAL As String, _
''''                                      strRAMALINICIAL As String, _
''''                                      strRAMALFINAL As String, _
''''                                      strREGIAOINICIAL As String, _
''''                                      strREGIAOFINAL As String, _
''''                                      strDESCRICAOINICIAL As String, _
''''                                      strDESCRICAOFINAL As String, _
''''                                      strINTERURBANOINICIAL As String, _
''''                                      strINTERURBANOFINAL As String, _
''''                                      strINTERURBANOSEPARACAO As String, _
''''                                      strCABECALHO As String)
''''
''''  Dim objConfiguracao As datSisMetal.clsConfiguracao
''''  Dim objRs As ADODB.Recordset
''''  '
''''  Set objConfiguracao = New datSisMetal.clsConfiguracao
''''
''''  'Inserir Configuracao - Locação
''''  Set objRs = objConfiguracao.ListarConfiguracaoMesa(lngCONFIGURACAOID)
''''  If objRs.EOF Then
''''    'Inserir ConfiguraçãoMesa
''''    objConfiguracao.InserirConfiguracaoMesa lngCONFIGURACAOID, _
''''                                            strVALORINICIAL, strVALORFINAL, strVALORSEPARACAO, _
''''                                            strDATAINICIAL, _
''''                                            strDATAFINAL, _
''''                                            strDATAMASCARA, _
''''                                            strHORAINICIAL, _
''''                                            strHORAFINAL, _
''''                                            strHORAMASCARA, _
''''                                            strDURACAOINICIAL, _
''''                                            strDURACAOFINAL, _
''''                                            strDURACAOMASCARA, _
''''                                            strNUMEROINICIAL, _
''''                                            strNUMEROFINAL, _
''''                                            strRAMALINICIAL, _
''''                                            strRAMALFINAL, _
''''                                            strREGIAOINICIAL, _
''''                                            strREGIAOFINAL, _
''''                                            strDESCRICAOINICIAL, _
''''                                            strDESCRICAOFINAL, _
''''                                            strINTERURBANOINICIAL, _
''''                                            strINTERURBANOFINAL, _
''''                                            strINTERURBANOSEPARACAO, _
''''                                            strCABECALHO
''''  Else
''''    'Alterar ConfiguraçãoMesa
''''    objConfiguracao.AlterarConfiguracaoMesa lngCONFIGURACAOID, _
''''                                            strVALORINICIAL, strVALORFINAL, strVALORSEPARACAO, _
''''                                            strDATAINICIAL, _
''''                                            strDATAFINAL, _
''''                                            strDATAMASCARA, _
''''                                            strHORAINICIAL, _
''''                                            strHORAFINAL, _
''''                                            strHORAMASCARA, _
''''                                            strDURACAOINICIAL, _
''''                                            strDURACAOFINAL, _
''''                                            strDURACAOMASCARA, _
''''                                            strNUMEROINICIAL, _
''''                                            strNUMEROFINAL, _
''''                                            strRAMALINICIAL, _
''''                                            strRAMALFINAL, _
''''                                            strREGIAOINICIAL, _
''''                                            strREGIAOFINAL, _
''''                                            strDESCRICAOINICIAL, _
''''                                            strDESCRICAOFINAL, _
''''                                            strINTERURBANOINICIAL, _
''''                                            strINTERURBANOFINAL, _
''''                                            strINTERURBANOSEPARACAO, _
''''                                            strCABECALHO
''''  End If
''''  '
''''  objRs.Close
''''  Set objRs = Nothing
''''  Set objConfiguracao = Nothing
''''  Exit Sub
''''trata:
''''  Err.Raise Err.Number, Err.Source, Err.Description
''''End Sub
''''
''''Public Sub ExcluirConfiguracaoMesa(ByVal lngCONFIGURACAOID As Long)
''''  Dim objConf As datSisMetal.clsConfiguracao
''''  '
''''  On Error GoTo trata
''''  Set objConf = New datSisMetal.clsConfiguracao
''''  '
''''  objConf.ExcluirConfiguracaoMesa lngCONFIGURACAOID
''''  '
''''  Set objConf = Nothing
''''  Exit Sub
''''trata:
''''  Err.Raise Err.Number, Err.Source, Err.Description
''''End Sub

Public Sub ExcluirConfiguracao(ByVal lngCONFIGURACAOID As Long)
  Dim objConf As datSisMetal.clsConfiguracao
  '
  On Error GoTo trata
  Set objConf = New datSisMetal.clsConfiguracao
  '
  objConf.ExcluirConfiguracao lngCONFIGURACAOID
  '
  Set objConf = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function VerificaExclusaoConfiguracao(lngCONFIGURACAOID As Long, _
                                             ByRef strDescricaoErro As String) As Boolean
  On Error GoTo trata
  Dim objConfig As datSisMetal.clsConfiguracao
  Set objConfig = New datSisMetal.clsConfiguracao

  'Verificar
  VerificaExclusaoConfiguracao = objConfig.VerificaExclusaoConfiguracao(lngCONFIGURACAOID, _
                                                                        strDescricaoErro)
  '
  Set objConfig = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub AlterarConfiguracaoGeral(lngCONFIGURACAOID As Long, _
                                    strCaminho As String, _
                                    strQtdDiasVenda As String)


  Dim blnConfiguracao As datSisMetal.clsConfiguracao
  Set blnConfiguracao = New datSisMetal.clsConfiguracao

  'Inserir Configuracao - Cortesia
  blnConfiguracao.AlterarConfiguracaoGeral lngCONFIGURACAOID, _
                                           strCaminho, _
                                           strQtdDiasVenda
  '
  Set blnConfiguracao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

''''Public Sub AlterarConfiguracaoLocacao(lngCONFIGURACAOID As Long, _
''''                                      blnSuiteAptoLimpo As Boolean, blnTrabComLiberacao As Boolean, blnTrabSaida As Boolean, blnTrabDepAssocEntra As Boolean, _
''''                                      blnDepositoObriga As Boolean, blnTrabExtraUnidade As Boolean, blnTrabOcupantesveiculo As Boolean, blnLiberarHrEntrada As Boolean, _
''''                                      blnTrabCampoHospede As Boolean, blnTrabComEntradaPessoa As Boolean, blnTrabcomexedPessoa As Boolean, blnTrabComTroco As Boolean, _
''''                                      strCompr As String, strHoraEntra As String, blnPortCancEntra As Boolean, blnRecAcessaFinanc As Boolean, blnLiberaFechaComFechaSemRec As Boolean, _
''''                                      blnPedirSenhaDescCortFecha As Boolean, blnPedirSenhaSupApos2Fecha As Boolean, strMinTolera As String, strMinReceb As String, _
''''                                      strMinTransfSuiteMenVr As String, _
''''                                      blnTrabComMesaTel As Boolean, _
''''                                      strCaminho As String, _
''''                                      strQtdDiasParaCompensar As String, _
''''                                      strQtdChequesBons As String, _
''''                                      blnPedirSenhaSupLibChqReceb As Boolean, _
''''                                      blnTrabComChequesBons As Boolean, _
''''                                      strTrabComTaxa As String, _
''''                                      strPercTaxa As String, _
''''                                      blnPedirSenhaSupDespesaTelefonia As Boolean, _
''''                                      blnPedirSenhaSupCancPed As Boolean, _
''''                                      blnPedirSenhaSupCliProb As Boolean, _
''''                                      blnPedirSenhaSupTransfUnid As Boolean, _
''''                                      blnPedirSenhaSupSangria As Boolean, blnTrabComFatura As Boolean, blnPedirSenhaSupExcDep As Boolean)
''''
''''  Dim blnConfiguracao As datSisMetal.clsConfiguracao
''''  Set blnConfiguracao = New datSisMetal.clsConfiguracao
''''
''''  'Inserir Configuracao - Locação
''''  blnConfiguracao.AlterarConfiguracaoLocacao lngCONFIGURACAOID, _
''''                                             blnSuiteAptoLimpo, blnTrabComLiberacao, blnTrabSaida, blnTrabDepAssocEntra, _
''''                                             blnDepositoObriga, blnTrabExtraUnidade, blnTrabOcupantesveiculo, blnLiberarHrEntrada, _
''''                                             blnTrabCampoHospede, blnTrabComEntradaPessoa, blnTrabcomexedPessoa, blnTrabComTroco, _
''''                                             strCompr, strHoraEntra, blnPortCancEntra, blnRecAcessaFinanc, blnLiberaFechaComFechaSemRec, _
''''                                             blnPedirSenhaDescCortFecha, blnPedirSenhaSupApos2Fecha, strMinTolera, strMinReceb, _
''''                                             strMinTransfSuiteMenVr, _
''''                                             blnTrabComMesaTel, _
''''                                             strCaminho, _
''''                                             strQtdDiasParaCompensar, _
''''                                             strQtdChequesBons, _
''''                                             blnPedirSenhaSupLibChqReceb, _
''''                                             blnTrabComChequesBons, _
''''                                             strTrabComTaxa, _
''''                                             strPercTaxa, _
''''                                             blnPedirSenhaSupDespesaTelefonia, _
''''                                             blnPedirSenhaSupCancPed, _
''''                                             blnPedirSenhaSupCliProb, _
''''                                             blnPedirSenhaSupTransfUnid, blnPedirSenhaSupSangria, blnTrabComFatura, blnPedirSenhaSupExcDep
''''
''''  '
''''  Set blnConfiguracao = Nothing
''''  Exit Sub
''''trata:
''''  Err.Raise Err.Number, Err.Source, Err.Description
''''End Sub
Public Sub InserirConfiguracao()


  Dim objConfiguracao As datSisMetal.clsConfiguracao
  Set objConfiguracao = New datSisMetal.clsConfiguracao

  'Inserir Configuracao
  objConfiguracao.InserirConfiguracao
  '
  Set objConfiguracao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarConfiguracaoCadastro(lngCONFIGURACAOID As Long, _
                                       strEmpresa As String, _
                                       strCnpj As String, _
                                       strInscrMunicipal As String, _
                                       strEndereco As String, _
                                       strBairro As String, _
                                       strCidade As String, _
                                       strEstado As String, _
                                       strCep As String, _
                                       strTel As String, _
                                       strTitulo As String)


  Dim blnConfiguracao As datSisMetal.clsConfiguracao
  Set blnConfiguracao = New datSisMetal.clsConfiguracao

  'Inserir Configuracao - Cadastro
  blnConfiguracao.AlterarConfiguracaoCadastro lngCONFIGURACAOID, _
                                              strEmpresa, _
                                              strCnpj, _
                                              strInscrMunicipal, _
                                              strEndereco, _
                                              strBairro, _
                                              strCidade, _
                                              strEstado, _
                                              strCep, _
                                              strTel, _
                                              strTitulo

  '
  Set blnConfiguracao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
''''Public Sub AlterarConfiguracaoTelEntrPed(lngCONFIGURACAOID As Long, _
''''                                        strVrTarifaMin As String, _
''''                                        strVrTarifaMinInter As String, _
''''                                        strQtdMinDir As String, _
''''                                        strQtdMinGer As String, _
''''                                        strQtdMinRec As String, _
''''                                        strQtdMinPort As String, _
''''                                        strQtdPedMinDir As String, _
''''                                        strQtdPedMinGer As String, _
''''                                        strQtdPedMinRec As String, _
''''                                        strVrTarifaMinInternac As String, _
''''                                        strDDDLocal As String)
''''
''''
''''  Dim blnConfiguracao As datSisMetal.clsConfiguracao
''''  Set blnConfiguracao = New datSisMetal.clsConfiguracao
''''
''''  'Inserir Configuracao - TelEntrPed
''''  blnConfiguracao.AlterarConfiguracaoTelEntrPed lngCONFIGURACAOID, _
''''                                                strVrTarifaMin, _
''''                                                strVrTarifaMinInter, _
''''                                                strQtdMinDir, _
''''                                                strQtdMinGer, _
''''                                                strQtdMinRec, _
''''                                                strQtdMinPort, _
''''                                                strQtdPedMinDir, _
''''                                                strQtdPedMinGer, _
''''                                                strQtdPedMinRec, _
''''                                                strVrTarifaMinInternac, _
''''                                                strDDDLocal
''''
''''  '
''''  Set blnConfiguracao = Nothing
''''  Exit Sub
''''trata:
''''  Err.Raise Err.Number, Err.Source, Err.Description
''''End Sub
''''Public Sub AlterarConfiguracaoDespDiaCaixa(lngCONFIGURACAOID As Long, _
''''                                           blnHabServDespDir As Boolean, _
''''                                           blnHabServDespGer As Boolean, _
''''                                           blnHabServDespRec As Boolean, _
''''                                           blnHabServDespPort As Boolean, _
''''                                           blnHabDiariaDir As Boolean, _
''''                                           blnHabDiariaGer As Boolean, _
''''                                           blnHabDiariaRec As Boolean, _
''''                                           blnHabDiariaPor As Boolean, _
''''                                           strIntervBase As String, _
''''                                           strIntervOper As String, _
''''                                           blnHabMovCaixaDir As Boolean, _
''''                                           blnHabMovCaixaGer As Boolean, _
''''                                           blnHabMovCaixaRec As Boolean, _
''''                                           blnHabMovCaixaPort As Boolean, _
''''                                           strTetoMovCaixa As String)
''''
''''
''''  Dim blnConfiguracao As datSisMetal.clsConfiguracao
''''  Set blnConfiguracao = New datSisMetal.clsConfiguracao
''''
''''  'Inserir Configuracao - DespDiaCaixa
''''  blnConfiguracao.AlterarConfiguracaoDespDiaCaixa lngCONFIGURACAOID, _
''''                                                  blnHabServDespDir, _
''''                                                  blnHabServDespGer, _
''''                                                  blnHabServDespRec, _
''''                                                  blnHabServDespPort, _
''''                                                  blnHabDiariaDir, _
''''                                                  blnHabDiariaGer, _
''''                                                  blnHabDiariaRec, _
''''                                                  blnHabDiariaPor, _
''''                                                  strIntervBase, _
''''                                                  strIntervOper, _
''''                                                  blnHabMovCaixaDir, _
''''                                                  blnHabMovCaixaGer, _
''''                                                  blnHabMovCaixaRec, _
''''                                                  blnHabMovCaixaPort, _
''''                                                  strTetoMovCaixa
''''
''''  '
''''  Set blnConfiguracao = Nothing
''''  Exit Sub
''''trata:
''''  Err.Raise Err.Number, Err.Source, Err.Description
''''End Sub
''''Public Sub AlterarConfiguracaoDiversos(lngCONFIGURACAOID As Long, _
''''                                       blnPedirSenhaFechaTurno As Boolean, _
''''                                       blnPedirSenhaVdaDiretoria As Boolean, _
''''                                       blnEstCadastraCardapio As Boolean, _
''''                                       blnGerModifConfig As Boolean, _
''''                                       blnTrabComEstorno As Boolean, _
''''                                       blnTrabDepSan As Boolean, _
''''                                       blnTrabComVendasCob As Boolean, _
''''                                       strVrInicialCaixa As String, _
''''                                       strSequencial As String, _
''''                                       strDataInicialRelMov As String, _
''''                                       blnTrabComEstInter As Boolean, _
''''                                       blnTrabComEstNegativo As Boolean, _
''''                                       strDiaFechaFolha As String, _
''''                                       strVrISS As String, _
''''                                       blnTrabComExcPenhor As Boolean, _
''''                                       blnTrabComAltVda As Boolean)
''''
''''
''''
''''  Dim blnConfiguracao As datSisMetal.clsConfiguracao
''''  Set blnConfiguracao = New datSisMetal.clsConfiguracao
''''
''''  'Inserir Configuracao - Diversos
''''  blnConfiguracao.AlterarConfiguracaoDiversos lngCONFIGURACAOID, _
''''                                              blnPedirSenhaFechaTurno, _
''''                                              blnPedirSenhaVdaDiretoria, _
''''                                              blnEstCadastraCardapio, _
''''                                              blnGerModifConfig, _
''''                                              blnTrabComEstorno, _
''''                                              blnTrabDepSan, _
''''                                              blnTrabComVendasCob, _
''''                                              strVrInicialCaixa, _
''''                                              strSequencial, _
''''                                              strDataInicialRelMov, _
''''                                              blnTrabComEstInter, _
''''                                              blnTrabComEstNegativo, _
''''                                              strDiaFechaFolha, _
''''                                              strVrISS, _
''''                                              blnTrabComExcPenhor, _
''''                                              blnTrabComAltVda
''''
''''  '
''''  Set blnConfiguracao = Nothing
''''  Exit Sub
''''trata:
''''  Err.Raise Err.Number, Err.Source, Err.Description
''''End Sub
''''
''''
''''Public Sub AlterarConfiguracaoImpressao(lngCONFIGURACAOID As Long, _
''''                                        blnImprimirVendasFechaTurno As Boolean, _
''''                                        blnImprimirTurnoFecha As Boolean, _
''''                                        blnImprimirCartaoFechaTur As Boolean, _
''''                                        blnImprimirPagamentos As Boolean, _
''''                                        blnImprimirContas As Boolean, _
''''                                        blnImprimirPedFrigo As Boolean, _
''''                                        blnImprimirComprovLoc As Boolean, _
''''                                        blnTrabComImpFiscal As Boolean, _
''''                                        blnImpRecebeTurnosAnt As Boolean, _
''''                                        blnImprimirCabec As Boolean, _
''''                                        strPontosBol As String, _
''''                                        strTipoImpFiscal As String, _
''''                                        blnTrabTabData As Boolean)
''''
''''
''''  Dim blnConfiguracao As datSisMetal.clsConfiguracao
''''  Set blnConfiguracao = New datSisMetal.clsConfiguracao
''''
''''  'Inserir Configuracao - Impressão
''''  blnConfiguracao.AlterarConfiguracaoImpressao lngCONFIGURACAOID, _
''''                                               blnImprimirVendasFechaTurno, _
''''                                               blnImprimirTurnoFecha, _
''''                                               blnImprimirCartaoFechaTur, _
''''                                               blnImprimirPagamentos, _
''''                                               blnImprimirContas, _
''''                                               blnImprimirPedFrigo, _
''''                                               blnImprimirComprovLoc, _
''''                                               blnTrabComImpFiscal, _
''''                                               blnImpRecebeTurnosAnt, _
''''                                               blnImprimirCabec, _
''''                                               strPontosBol, _
''''                                               strTipoImpFiscal, _
''''                                               blnTrabTabData
''''
''''  '
''''  Set blnConfiguracao = Nothing
''''  Exit Sub
''''trata:
''''  Err.Raise Err.Number, Err.Source, Err.Description
''''End Sub
''''
''''Public Sub AlterarConfiguracaoFechamento(lngCONFIGURACAOID As Long, _
''''                                        intTipo As Integer, _
''''                                        intTipoFechamento As Integer, _
''''                                        blnTrabComFechaCC As Boolean)
''''
''''
''''  Dim blnConfiguracao As datSisMetal.clsConfiguracao
''''  Set blnConfiguracao = New datSisMetal.clsConfiguracao
''''
''''  'Inserir Configuracao - Fechamento
''''  blnConfiguracao.AlterarConfiguracaoFechamento lngCONFIGURACAOID, _
''''                                                intTipo, _
''''                                                intTipoFechamento, _
''''                                                blnTrabComFechaCC
''''
''''
''''  '
''''  Set blnConfiguracao = Nothing
''''  Exit Sub
''''trata:
''''  Err.Raise Err.Number, Err.Source, Err.Description
''''End Sub
''''
''''Public Function ListarConfiguracaoLocacao(lngCONFIGURACAOID As Long) As ADODB.Recordset
''''  On Error GoTo trata
''''  Dim objConfiguracao As datSisMetal.clsConfiguracao
''''  Set objConfiguracao = New datSisMetal.clsConfiguracao
''''
''''  'Listar Configuracao - Locacao
''''  Set ListarConfiguracaoLocacao = objConfiguracao.ListarConfiguracaoLocacao(lngCONFIGURACAOID)
''''  '
''''  Set objConfiguracao = Nothing
''''  Exit Function
''''trata:
''''  Err.Raise Err.Number, Err.Source, Err.Description
''''End Function
''''
Public Function ListarConfiguracaoCadastro(lngCONFIGURACAOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objConfiguracao As datSisMetal.clsConfiguracao
  Set objConfiguracao = New datSisMetal.clsConfiguracao

  'Listar Configuracao - Cadastro
  Set ListarConfiguracaoCadastro = objConfiguracao.ListarConfiguracaoCadastro(lngCONFIGURACAOID)
  '
  Set objConfiguracao = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ListarConfiguracao() As ADODB.Recordset
  On Error GoTo trata
  Dim objConfiguracao As datSisMetal.clsConfiguracao
  Set objConfiguracao = New datSisMetal.clsConfiguracao

  'Listar Configuracao
  Set ListarConfiguracao = objConfiguracao.ListarConfiguracao
  '
  Set objConfiguracao = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


''''Public Function ListarConfiguracaoTelEntrPed(lngCONFIGURACAOID As Long) As ADODB.Recordset
''''  On Error GoTo trata
''''  Dim objConfiguracao As datSisMetal.clsConfiguracao
''''  Set objConfiguracao = New datSisMetal.clsConfiguracao
''''
''''  'Listar Configuracao - TelEntrPed
''''  Set ListarConfiguracaoTelEntrPed = objConfiguracao.ListarConfiguracaoTelEntrPed(lngCONFIGURACAOID)
''''  '
''''  Set objConfiguracao = Nothing
''''  Exit Function
''''trata:
''''  Err.Raise Err.Number, Err.Source, Err.Description
''''End Function
''''
''''Public Function ListarConfiguracaoDespDiaCaixa(lngCONFIGURACAOID As Long) As ADODB.Recordset
''''  On Error GoTo trata
''''  Dim objConfiguracao As datSisMetal.clsConfiguracao
''''  Set objConfiguracao = New datSisMetal.clsConfiguracao
''''
''''  'Listar Configuracao - DespDiaCaixa
''''  Set ListarConfiguracaoDespDiaCaixa = objConfiguracao.ListarConfiguracaoDespDiaCaixa(lngCONFIGURACAOID)
''''  '
''''  Set objConfiguracao = Nothing
''''  Exit Function
''''trata:
''''  Err.Raise Err.Number, Err.Source, Err.Description
''''End Function
''''
''''Public Function ListarConfiguracaoDiversos(lngCONFIGURACAOID As Long) As ADODB.Recordset
''''  On Error GoTo trata
''''  Dim objConfiguracao As datSisMetal.clsConfiguracao
''''  Set objConfiguracao = New datSisMetal.clsConfiguracao
''''
''''  'Listar Configuracao - Diversos
''''  Set ListarConfiguracaoDiversos = objConfiguracao.ListarConfiguracaoDiversos(lngCONFIGURACAOID)
''''  '
''''  Set objConfiguracao = Nothing
''''  Exit Function
''''trata:
''''  Err.Raise Err.Number, Err.Source, Err.Description
''''End Function
''''
''''Public Function ListarConfiguracaoImpressao(lngCONFIGURACAOID As Long) As ADODB.Recordset
''''  On Error GoTo trata
''''  Dim objConfiguracao As datSisMetal.clsConfiguracao
''''  Set objConfiguracao = New datSisMetal.clsConfiguracao
''''
''''  'Listar Configuracao - Impressao
''''  Set ListarConfiguracaoImpressao = objConfiguracao.ListarConfiguracaoImpressao(lngCONFIGURACAOID)
''''  '
''''  Set objConfiguracao = Nothing
''''  Exit Function
''''trata:
''''  Err.Raise Err.Number, Err.Source, Err.Description
''''End Function
Public Function ListarConfiguracaoGeral(lngCONFIGURACAOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objConfiguracao As datSisMetal.clsConfiguracao
  Set objConfiguracao = New datSisMetal.clsConfiguracao

  'Listar Configuracao - Cortesia
  Set ListarConfiguracaoGeral = objConfiguracao.ListarConfiguracaoGeral(lngCONFIGURACAOID)
  '
  Set objConfiguracao = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

''''Public Function ListarConfiguracaoFechamento(lngCONFIGURACAOID As Long) As ADODB.Recordset
''''  On Error GoTo trata
''''  Dim objConfiguracao As datSisMetal.clsConfiguracao
''''  Set objConfiguracao = New datSisMetal.clsConfiguracao
''''
''''  'Listar Configuracao - Fechamento
''''  Set ListarConfiguracaoFechamento = objConfiguracao.ListarConfiguracaoFechamento(lngCONFIGURACAOID)
''''  '
''''  Set objConfiguracao = Nothing
''''  Exit Function
''''trata:
''''  Err.Raise Err.Number, Err.Source, Err.Description
''''End Function
''''
