Attribute VB_Name = "mdlUserImpressoraFiscal"
Option Explicit

Public Declare Function Bematech_FI_NumeroSerie Lib "BemaFi32.dll" (ByVal NumeroSerie As String) As Integer
Public Declare Function Bematech_FI_SubTotal Lib "BemaFi32.dll" (ByVal SubTotal As String) As Integer
Public Declare Function Bematech_FI_NumeroCupom Lib "BemaFi32.dll" (ByVal NumeroCupom As String) As Integer
Public Declare Function Bematech_FI_ResetaImpressora Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_AbrePortaSerial Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_LeituraX Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_LeituraXSerial Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_AbreCupom Lib "BemaFi32.dll" (ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FI_VendeItem Lib "BemaFi32.dll" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal TipoQuantidade As String, ByVal Quantidade As String, ByVal CasasDecimais As Integer, ByVal ValorUnitario As String, ByVal TipoDesconto As String, ByVal Desconto As String) As Integer
Public Declare Function Bematech_FI_CancelaItemAnterior Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_CancelaItemGenerico Lib "BemaFi32.dll" (ByVal NumeroItem As String) As Integer
Public Declare Function Bematech_FI_CancelaCupom Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_FechaCupomResumido Lib "BemaFi32.dll" (ByVal FormaPagamento As String, ByVal Mensagem As String) As Integer '
Public Declare Function Bematech_FI_ReducaoZ Lib "BemaFi32.dll" (ByVal Data As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_FechaCupom Lib "BemaFi32.dll" (ByVal FormaPagamento As String, ByVal DiscontoAcrecimo As String, ByVal TipoDescontoAcrecimo As String, ByVal ValorAcrecimoDesconto As String, ByVal ValorPago As String, ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_VendeItemDepartamento Lib "BemaFi32.dll" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal ValorUnitario As String, ByVal Quantidade As String, ByVal Acrescimo As String, ByVal Desconto As String, ByVal IndiceDepartamento As String, ByVal UnidadeMedida As String) As Integer
Public Declare Function Bematech_FI_AumentaDescricaoItem Lib "BemaFi32.dll" (ByVal Descricao As String) As Integer
Public Declare Function Bematech_FI_UsaUnidadeMedida Lib "BemaFi32.dll" (ByVal UnidadeMedida As String) As Integer
Public Declare Function Bematech_FI_AlteraSimboloMoeda Lib "BemaFi32.dll" (ByVal SimboloMoeda As String) As Integer
Public Declare Function Bematech_FI_ProgramaAliquota Lib "BemaFi32.dll" (ByVal Aliquota As String, ByVal ICMS_ISS As Integer) As Integer
Public Declare Function Bematech_FI_ProgramaHorarioVerao Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_NomeiaDepartamento Lib "BemaFi32.dll" (ByVal Indice As Integer, ByVal Departamento As String) As Integer
Public Declare Function Bematech_FI_NomeiaTotalizadorNaoSujeitoIcms Lib "BemaFi32.dll" (ByVal Indice As Integer, ByVal Totalizador As String) As Integer
Public Declare Function Bematech_FI_ProgramaArredondamento Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_ProgramaTruncamento Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_LinhasEntreCupons Lib "BemaFi32.dll" (ByVal Linhas As Integer) As Integer
Public Declare Function Bematech_FI_EspacoEntreLinhas Lib "BemaFi32.dll" (ByVal Dots As Integer) As Integer
Public Declare Function Bematech_FI_RelatorioGerencial Lib "BemaFi32.dll" (ByVal cTexto As String) As Integer
Public Declare Function Bematech_FI_FechaRelatorioGerencial Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_RecebimentoNaoFiscal Lib "BemaFi32.dll" (ByVal IndiceTotalizador As String, ByVal Valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Bematech_FI_AbreComprovanteNaoFiscalVinculado Lib "BemaFi32.dll" (ByVal FormaPagamento As String, ByVal Valor As String, ByVal NumeroCupom As String) As Integer
Public Declare Function Bematech_FI_UsaComprovanteNaoFiscalVinculado Lib "BemaFi32.dll" (ByVal Texto As String) As Integer
Public Declare Function Bematech_FI_FechaComprovanteNaoFiscalVinculado Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_Sangria Lib "BemaFi32.dll" (ByVal Valor As String) As Integer
Public Declare Function Bematech_FI_Suprimento Lib "BemaFi32.dll" (ByVal Valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalData Lib "BemaFi32.dll" (ByVal cDataInicial As String, ByVal cDataFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalReducao Lib "BemaFi32.dll" (ByVal cReducaoInicial As String, ByVal cReducaoFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalSerialData Lib "BemaFi32.dll" (ByVal cDataInicial As String, ByVal cDataFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalSerialReducao Lib "BemaFi32.dll" (ByVal cReducaoInicial As String, ByVal cReducaoFinal As String) As Integer
Public Declare Function Bematech_FI_VersaoFirmware Lib "BemaFi32.dll" (ByVal VersaoFirmware As String) As Integer
Public Declare Function Bematech_FI_CGC_IE Lib "BemaFi32.dll" (ByVal CGC As String, ByVal IE As String) As Integer
Public Declare Function Bematech_FI_GrandeTotal Lib "BemaFi32.dll" (ByVal GrandeTotal As String) As Integer
Public Declare Function Bematech_FI_Cancelamentos Lib "BemaFi32.dll" (ByVal ValorCancelamentos As String) As Integer
Public Declare Function Bematech_FI_Descontos Lib "BemaFi32.dll" (ByVal ValorDescontos As String) As Integer
Public Declare Function Bematech_FI_NumeroOperacoesNaoFiscais Lib "BemaFi32.dll" (ByVal NumeroOperacoes As String) As Integer
Public Declare Function Bematech_FI_NumeroCuponsCancelados Lib "BemaFi32.dll" (ByVal NumeroCancelamentos As String) As Integer
Public Declare Function Bematech_FI_NumeroIntervencoes Lib "BemaFi32.dll" (ByVal NumeroIntervencoes As String) As Integer
Public Declare Function Bematech_FI_NumeroReducoes Lib "BemaFi32.dll" (ByVal NumeroReducoes As String) As Integer
Public Declare Function Bematech_FI_NumeroSubstituicoesProprietario Lib "BemaFi32.dll" (ByVal NumeroSubstituicoes As String) As Integer
Public Declare Function Bematech_FI_UltimoItemVendido Lib "BemaFi32.dll" (ByVal NumeroItem As String) As Integer
Public Declare Function Bematech_FI_ClicheProprietario Lib "BemaFi32.dll" (ByVal Cliche As String) As Integer
Public Declare Function Bematech_FI_NumeroCaixa Lib "BemaFi32.dll" (ByVal NumeroCaixa As String) As Integer
Public Declare Function Bematech_FI_NumeroLoja Lib "BemaFi32.dll" (ByVal NumeroLoja As String) As Integer
Public Declare Function Bematech_FI_SimboloMoeda Lib "BemaFi32.dll" (ByVal SimboloMoeda As String) As Integer
Public Declare Function Bematech_FI_MinutosLigada Lib "BemaFi32.dll" (ByVal Minutos As String) As Integer
Public Declare Function Bematech_FI_MinutosImprimindo Lib "BemaFi32.dll" (ByVal Minutos As String) As Integer
Public Declare Function Bematech_FI_VerificaModoOperacao Lib "BemaFi32.dll" (ByVal Modo As String) As Integer
Public Declare Function Bematech_FI_VerificaEpromConectada Lib "BemaFi32.dll" (ByVal Flag As String) As Integer
Public Declare Function Bematech_FI_FlagsFiscais Lib "BemaFi32.dll" (ByRef Flag As Integer) As Integer
Public Declare Function Bematech_FI_ValorPagoUltimoCupom Lib "BemaFi32.dll" (ByVal ValorCupom As String) As Integer
Public Declare Function Bematech_FI_DataHoraImpressora Lib "BemaFi32.dll" (ByVal Data As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_ContadoresTotalizadoresNaoFiscais Lib "BemaFi32.dll" (ByVal Contadores As String) As Integer
Public Declare Function Bematech_FI_VerificaTotalizadoresNaoFiscais Lib "BemaFi32.dll" (ByVal Totalizadores As String) As Integer
Public Declare Function Bematech_FI_DataHoraReducao Lib "BemaFi32.dll" (ByVal Data As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_DataMovimento Lib "BemaFi32.dll" (ByVal Data As String) As Integer
Public Declare Function Bematech_FI_VerificaTruncamento Lib "BemaFi32.dll" (ByVal Flag As String) As Integer
Public Declare Function Bematech_FI_Acrescimos Lib "BemaFi32.dll" (ByVal ValorAcrescimos As String) As Integer
Public Declare Function Bematech_FI_ContadorBilhetePassagem Lib "BemaFi32.dll" (ByVal ContadorPassagem As String) As Integer
Public Declare Function Bematech_FI_VerificaAliquotasIss Lib "BemaFi32.dll" (ByVal Flag As String) As Integer
Public Declare Function Bematech_FI_VerificaFormasPagamento Lib "BemaFi32.dll" (ByVal Formas As String) As Integer
Public Declare Function Bematech_FI_VerificaRecebimentoNaoFiscal Lib "BemaFi32.dll" (ByVal Recebimentos As String) As Integer
Public Declare Function Bematech_FI_VerificaDepartamentos Lib "BemaFi32.dll" (ByVal Departamentos As String) As Integer
Public Declare Function Bematech_FI_VerificaTipoImpressora Lib "BemaFi32.dll" (ByRef TipoImpressora As Integer) As Integer
Public Declare Function Bematech_FI_VerificaTotalizadoresParciais Lib "BemaFi32.dll" (ByVal cTotalizadores As String) As Integer
Public Declare Function Bematech_FI_RetornoAliquotas Lib "BemaFi32.dll" (ByVal cAliquotas As String) As Integer
Public Declare Function Bematech_FI_VerificaEstadoImpressora Lib "BemaFi32.dll" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Public Declare Function Bematech_FI_DadosUltimaReducao Lib "BemaFi32.dll" (ByVal DadosReducao As String) As Integer
Public Declare Function Bematech_FI_MonitoramentoPapel Lib "BemaFi32.dll" (ByRef Linhas As Integer) As Integer
Public Declare Function Bematech_FI_Autenticacao Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_ProgramaCaracterAutenticacao Lib "BemaFi32.dll" (ByVal Parametros As String) As Integer
Public Declare Function Bematech_FI_AcionaGaveta Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_VerificaEstadoGaveta Lib "BemaFi32.dll" (ByRef EstadoGaveta As Integer) As Integer
Public Declare Function Bematech_FI_ProgramaMoedaSingular Lib "BemaFi32.dll" (ByVal MoedaSingular As String) As Integer
Public Declare Function Bematech_FI_ProgramaMoedaPlural Lib "BemaFi32.dll" (ByVal MoedaPlural As String) As Integer
Public Declare Function Bematech_FI_CancelaImpressaoCheque Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_VerificaStatusCheque Lib "BemaFi32.dll" (ByRef StatusCheque As Integer) As Integer
Public Declare Function Bematech_FI_ImprimeCheque Lib "BemaFi32.dll" (ByVal Banco As String, ByVal Valor As String, ByVal Favorecido As String, ByVal Cidade As String, ByVal Data As String, ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_IncluiCidadeFavorecido Lib "BemaFi32.dll" (ByVal Cidade As String, ByVal Favorecido As String) As Integer
Public Declare Function Bematech_FI_EstornoFormasPagamento Lib "BemaFi32.dll" (ByVal FormaOrigem As String, ByVal FormaDestino As String, ByVal Valor As String) As Integer
Public Declare Function Bematech_FI_ForcaImpactoAgulhas Lib "BemaFi32.dll" (ByVal ForcaImpacto As Integer) As Integer
Public Declare Function Bematech_FI_RetornoImpressora Lib "BemaFi32.dll" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Public Declare Function Bematech_FI_FechaPortaSerial Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_VerificaImpressoraLigada Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_IniciaFechamentoCupom Lib "BemaFi32.dll" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FI_EfetuaFormaPagamento Lib "BemaFi32.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String) As Integer
Public Declare Function Bematech_FI_TerminaFechamentoCupom Lib "BemaFi32.dll" (ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_AbreBilhetePassagem Lib "BemaFi32.dll" (ByVal ImprimeValorFinal As String, ByVal ImprimeEnfatizado As String, ByVal LocalEmbarque As String, ByVal Destino As String, ByVal Linha As String, ByVal Prefixo As String, ByVal Agente As String, ByVal Agencia As String, ByVal Data As String, ByVal Hora As String, ByVal Poltrona As String, ByVal Plataforma As String) As Integer
Public Declare Function Bematech_FI_MapaResumo Lib "BemaFi32.dll" () As Integer

'Funções para Impressora restaurante
Public Declare Function Bematech_FIR_RegistraVenda Lib "BemaFi32.dll" (ByVal Mesa As String, ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_CancelaVenda Lib "BemaFi32.dll" (ByVal Mesa As String, ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_ConferenciaMesa Lib "BemaFi32.dll" (ByVal Mesa As String, ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_AbreConferenciaMesa Lib "BemaFi32.dll" (ByVal Mesa As String) As Integer
Public Declare Function Bematech_FIR_FechaConferenciaMesa Lib "BemaFi32.dll" (ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_TransferenciaMesa Lib "BemaFi32.dll" (ByVal MesaOrigem As String, ByVal MesaDestino As String) As Integer
Public Declare Function Bematech_FIR_AbreCupomRestaurante Lib "BemaFi32.dll" (ByVal Mesa As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_ContaDividida Lib "BemaFi32.dll" (ByVal NumeroCupons As String, ByVal ValorPago As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_FechaCupomContaDividida Lib "BemaFi32.dll" (ByVal NumeroCupons As String, ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal FormasPagamento As String, ByVal ValorFormasPagamento As String, ByVal ValorPagoCliente As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_TransferenciaItem Lib "BemaFi32.dll" (ByVal MesaOrigem As String, ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal MesaDestino As String) As Integer
Public Declare Function Bematech_FIR_RelatorioMesasAbertas Lib "BemaFi32.dll" (ByVal TipoRelatorio As Integer) As Integer
Public Declare Function Bematech_FIR_ImprimeCardapio Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FIR_RelatorioMesasAbertasSerial Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FIR_CardapioPelaSerial Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FIR_RegistroVendaSerial Lib "BemaFi32.dll" (ByVal Mesa As String) As Integer
Public Declare Function Bematech_FIR_VerificaMemoriaLivre Lib "BemaFi32.dll" (ByVal Bytes As String) As Integer
Public Declare Function Bematech_FIR_FechaCupomRestaurante Lib "BemaFi32.dll" (ByVal FormaPagamento As String, ByVal DiscontoAcrecimo As String, ByVal TipoDescontoAcrecimo As String, ByVal ValorAcrecimoDesconto As String, ByVal ValorPago As String, ByVal Mensagem As String) As Integer



Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal LpAplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long

Public Retorno As Integer
Public Funcao As Integer
Public LocalRetorno As String




'Gravar Arquivo Texto na
Sub GravarArquivo(pCod As String, pDesc As String, pQtd As String, pValor As String)
  Dim I
  I = FreeFile
  Open App.Path & "\Pedido.txt" For Append As #I
  
  Print #I, Now() & ";" & pCod & ";" & pDesc & ";" & pQtd & ";" & pValor
  Close #I
End Sub

Public Sub IMP_CUPOM_FISCAL(pLocacaoID As String, pNomeMotel As String)
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  Dim objGeral As busApler.clsGeral
  Dim cTotal As Currency
  'Impr Fiscal
  Dim TipoQuantidade As String
  Dim CasasDecimais As Integer
  Dim TipoDesconto As String
  Dim Aliquota As String
  'Dim Retorno
  Dim iQtd As Long
  'On Error GoTo RotErr:
  On Error Resume Next
  '
  Set objGeral = New busApler.clsGeral
'  '-----------------
'  'Abre cupom Fiscal
'  frmAbreCupomRestaurante.Show vbModal
'  '-----------------
  strSql = "SELECT sum(TAB_PEDIDOCARD.QUANTIDADE) as QUANTIDADE, sum(TAB_PEDIDOCARD.VALOR) as SOMAVALOR, min(CARDAPIO.CODIGO) as CODIGO, min(CARDAPIO.ALIQUOTA) as ALIQUOTA, min(CARDAPIO.DESCRICAO) as DESCRICAO, min(CARDAPIO.VALOR) as VALOR " & _
    "From (PEDIDO PEDIDO INNER JOIN TAB_PEDIDOCARD TAB_PEDIDOCARD ON PEDIDO.PKID = TAB_PEDIDOCARD.PEDIDOID) INNER JOIN CARDAPIO CARDAPIO ON TAB_PEDIDOCARD.CARDAPIOID = CARDAPIO.PKID " & _
    " GROUP BY PEDIDO.ALOCACAOID, CARDAPIO.CODIGO" & _
    " HAVING sum(TAB_PEDIDOCARD.VALOR) > 0 " & _
    " AND PEDIDO.ALOCACAOID = " & pLocacaoID
  
  Set objRs = objGeral.ExecutarSQL(strSql)
  cTotal = 0
  iQtd = 0
  Do While Not objRs.EOF
    'Somar Ttotal
    cTotal = cTotal + (objRs!Valor * objRs!Quantidade)
    'Gravar Arquivo Texto
    GravarArquivo objRs!Codigo, objRs!Descricao, objRs!Quantidade, objRs!SOMAVALOR
    'Impressão Fiscal
    'frmVendaItem.txtCodigo.Text = objRs!Codigo
    'frmVendaItem.txtDescricao.Text = objRs!Descricao
    'frmVendaItem.txtQtde.Text = objRs!Quantidade
    'frmVendaItem.txtValorUnitario.Text = Format(objRs!Valor, "###,##0.00")
    
    'frmVendaItem.Show vbModal
    '------------------
    'IMPRESSORA FISCAL
    'Verifica se a quantidade é inteira ou fracionária
    
    TipoQuantidade = "I" 'inteira
    
    'Verifica se o valor unitário é com 2 ou 3 casas decimais
    CasasDecimais = 2
    
    'Verifica se o desconto é por valor ou por percentual
    TipoDesconto = "$" 'valor
    
    'Aliquota
    'Aliquota = IIf(IsNumeric(objRs!Aliquota), Format(objRs!Aliquota, "00.00"), objRs!Aliquota)
    Aliquota = objRs!Aliquota
    
    Retorno = Bematech_FI_VendeItem(objRs!Codigo, objRs!Descricao, _
              Aliquota, TipoQuantidade, objRs!Quantidade, CasasDecimais, _
              Format(objRs!Valor, "###,##0.00"), TipoDesconto, "0,00")

    'Função que analisa o retorno da impressora
    VerificaRetornoImpressora "", "", "Emissão de Cupom Fiscal"
    
    '
    iQtd = iQtd + 1
    objRs.MoveNext
  Loop
  objRs.Close
  Set objRs = Nothing
  '-----------------
  'Fecha Cupom Fiscal
  'Abre cupom Fiscal
  'frmFechaCupomRestaurante.Show vbModal
  If iQtd <> 0 Then
    Retorno = Bematech_FI_FechaCupomResumido("Dinheiro", "OBRIGADO VOLTE SEMPRE!")
    'Função que analisa o retorno da impressora
    VerificaRetornoImpressora "", "", "Fechamento do Cupom"
  End If
  '-----------------
  'FIM
  Set objGeral = Nothing
  Exit Sub
RotErr:
  MsgBox "O seguinte erro ocorreu: " & Err.Number & " - " & Err.Description
End Sub

Public Sub IMP_CUPOM_FISCAL_VENDA(pVendaID As Long, pNomeMotel As String)
  On Error GoTo trata
  Dim strSql          As String
  Dim objRs           As ADODB.Recordset
  Dim cTotal          As Currency
  'Impr Fiscal
  Dim TipoQuantidade  As String
  Dim CasasDecimais   As Integer
  Dim TipoDesconto    As String
  Dim Aliquota        As String
  Dim Retorno
  Dim iQtd            As Long
  Dim objGeral As busApler.clsGeral
  '
  Set objGeral = New busApler.clsGeral
  '
  '
'  '-----------------
'  'Abre cupom Fiscal
'  frmAbreCupomRestaurante.Show vbModal
'  '-----------------
  strSql = "SELECT sum(TAB_VENDACARD.QUANTIDADE) as QUANTIDADE, sum(TAB_VENDACARD.VALOR) as SOMAVALOR, min(CARDAPIO.CODIGO) as CODIGO, min(CARDAPIO.ALIQUOTA) as ALIQUOTA, min(CARDAPIO.DESCRICAO) as DESCRICAO, min(CARDAPIO.VALOR) as VALOR " & _
    "From (VENDA VENDA INNER JOIN TAB_VENDACARD TAB_VENDACARD ON VENDA.PKID = TAB_VENDACARD.VENDAID) INNER JOIN CARDAPIO CARDAPIO ON TAB_VENDACARD.CARDAPIOID = CARDAPIO.PKID " & _
    " GROUP BY VENDA.PKID, CARDAPIO.CODIGO" & _
    " HAVING sum(TAB_VENDACARD.VALOR) > 0 " & _
    " AND VENDA.PKID = " & pVendaID
  
  Set objRs = objGeral.ExecutarSQL(strSql)
  cTotal = 0
  iQtd = 0
  Do While Not objRs.EOF
    'Somar Ttotal
    cTotal = cTotal + (objRs!Valor * objRs!Quantidade)
    'Gravar Arquivo Texto
    GravarArquivo objRs!Codigo, objRs!Descricao, objRs!Quantidade, objRs!SOMAVALOR
    'Impressão Fiscal
    'frmVendaItem.txtCodigo.Text = objRs!Codigo
    'frmVendaItem.txtDescricao.Text = objRs!Descricao
    'frmVendaItem.txtQtde.Text = objRs!Quantidade
    'frmVendaItem.txtValorUnitario.Text = Format(objRs!Valor, "###,##0.00")
    
    'frmVendaItem.Show vbModal
    '------------------
    'IMPRESSORA FISCAL
    'Verifica se a quantidade é inteira ou fracionária
    
    TipoQuantidade = "I" 'inteira
    
    'Verifica se o valor unitário é com 2 ou 3 casas decimais
    CasasDecimais = 2
    
    'Verifica se o desconto é por valor ou por percentual
    TipoDesconto = "$" 'valor
    
    'Aliquota
    'Aliquota = IIf(IsNumeric(objRs!Aliquota), Format(objRs!Aliquota, "00.00"), objRs!Aliquota)
    Aliquota = objRs!Aliquota
    
    Retorno = Bematech_FI_VendeItem(objRs!Codigo, objRs!Descricao, _
              Aliquota, TipoQuantidade, objRs!Quantidade, CasasDecimais, _
              Format(objRs!Valor, "###,##0.00"), TipoDesconto, "0,00")

    'Função que analisa o retorno da impressora
    VerificaRetornoImpressora "", "", "Emissão de Cupom Fiscal"
    
    '
    iQtd = iQtd + 1
    objRs.MoveNext
  Loop
  objRs.Close
  Set objRs = Nothing
  '-----------------
  'Fecha Cupom Fiscal
  'Abre cupom Fiscal
  'frmFechaCupomRestaurante.Show vbModal
  If iQtd <> 0 Then
    Retorno = Bematech_FI_FechaCupomResumido("Dinheiro", "OBRIGADO VOLTE SEMPRE!")
    'Função que analisa o retorno da impressora
    VerificaRetornoImpressora "", "", "Fechamento do Cupom"
  End If
  '-----------------
  'FIM
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub


'Ler os Valores dos parâmetros nas seções do arquivo ini
Function LeParametrosIni(Secao As String, label As String) As String
   Dim Tst
   Const TamanhoParametro = 80
   Dim ParametroIni As String * TamanhoParametro
   Dim RetornoFuncao
   Dim Contador As Integer
   ParametroIni = ""
     
   RetornoFuncao = GetPrivateProfileString(Secao, label, "-2", ParametroIni, TamanhoParametro, "BemaFI32.ini")
   RetornoFuncao = Mid(ParametroIni, 1, 2)
   If Val(RetornoFuncao) <> -2 Then
       Contador = 1
       Do
           Tst = Mid(ParametroIni, Contador, 1)
           If Asc(Tst) <> 0 Then
               Contador = Contador + 1
           End If
       Loop While ((Asc(Tst) <> 0) And (Contador < Len(ParametroIni)))
       RetornoFuncao = Mid(ParametroIni, 1, Contador)
   End If
   LeParametrosIni = RetornoFuncao
End Function

Public Sub CentralizaJanela(Form As Form)
    Form.Top = (Screen.Height - Form.Height) / 2
    Form.Left = (Screen.Width - Form.Width) / 2
End Sub

Public Function VerificaRetornoImpressora(label As String, RetornoFuncao As String, TituloJanela As String)
    Dim ACK As Integer
    Dim ST1 As Integer
    Dim ST2 As Integer
    Dim RetornaMensagem As Integer
    Dim StringRetorno As String
    Dim ValorRetorno As String
    Dim RetornoStatus As Integer
    Dim Mensagem As String
    
    If Retorno = 0 Then
        MsgBox "Erro de comunicação com a impressora.", vbOKOnly + vbCritical, TituloJanela
        Exit Function
    
    ElseIf Retorno = 1 Then
        RetornoStatus = Bematech_FI_RetornoImpressora(ACK, ST1, ST2)
        ValorRetorno = Str(ACK) & "," & Str(ST1) & "," & Str(ST2)
        
        If label <> "" And RetornoFuncao <> "" Then
            RetornaMensagem = 1
        End If
        
        If ACK = 21 Then
            MsgBox "Status da Impressora: 21" & vbCr & vbLf & "Comando não executado", vbOKOnly + vbInformation, TituloJanela
            Exit Function
        End If
        
        If (ST1 <> 0 Or ST2 <> 0) Then
                If (ST1 >= 128) Then
                    StringRetorno = "Fim de Papel" & vbCr
                    ST1 = ST1 - 128
                End If
                
                If (ST1 >= 64) Then
                    StringRetorno = StringRetorno & "Pouco Papel" & vbCr
                    ST1 = ST1 - 64
                End If
                
                If (ST1 >= 32) Then
                    StringRetorno = StringRetorno & "Erro no relógio" & vbCr
                    ST1 = ST1 - 32
                End If
                
                If (ST1 >= 16) Then
                    StringRetorno = StringRetorno & "Impressora em erro" & vbCr
                    ST1 = ST1 - 16
                End If
                    
                If (ST1 >= 8) Then
                    StringRetorno = StringRetorno & "Primeiro dado do comando não foi Esc" & vbCr
                    ST1 = ST1 - 8
                End If
                
                If (ST1 >= 4) Then
                    StringRetorno = StringRetorno & "Comando inexistente" & vbCr
                    ST1 = ST1 - 4
                End If
                    
                If (ST1 >= 2) Then
                    StringRetorno = StringRetorno & "Cupom fiscal aberto" & vbCr
                    ST1 = ST1 - 2
                End If
                
                If (ST1 >= 1) Then
                    StringRetorno = StringRetorno & "Número de parâmetros inválido no comando" & vbCr
                    ST1 = ST1 - 1
                End If
                    
                If (ST2 >= 128) Then
                    StringRetorno = "Tipo de Parâmetro de comando inválido" & vbCr
                    ST2 = ST2 - 128
                End If
                
                If (ST2 >= 64) Then
                    StringRetorno = StringRetorno & "Memória fiscal lotada" & vbCr
                    ST2 = ST2 - 64
                End If
                
                If (ST2 >= 32) Then
                    StringRetorno = StringRetorno & "Erro na CMOS" & vbCr
                    ST2 = ST2 - 32
                End If
                
                If (ST2 >= 16) Then
                    StringRetorno = StringRetorno & "Alíquota não programada" & vbCr
                    ST2 = ST2 - 16
                End If
                    
                If (ST2 >= 8) Then
                    StringRetorno = StringRetorno & "Capacidade de alíquota programáveis lotada" & vbCr
                    ST2 = ST2 - 8
                End If
                
                If (ST2 >= 4) Then
                    StringRetorno = StringRetorno & "Cancelamento não permitido" & vbCr
                    ST2 = ST2 - 4
                End If
                    
                If (ST2 >= 2) Then
                    StringRetorno = StringRetorno & "CGC/IE do proprietário não programados" & vbCr
                    ST2 = ST2 - 2
                End If
                
                If (ST2 >= 1) Then
                    StringRetorno = StringRetorno & "Comando não executado" & vbCr
                    ST2 = ST2 - 1
                End If
                
                If RetornaMensagem Then
                    Mensagem = "Status da Impressora: " & ValorRetorno & _
                           vbCr & vbLf & StringRetorno & vbCr & vbLf & _
                           label & RetornoFuncao
                Else
                    Mensagem = "Status da Impressora: " & ValorRetorno & _
                       vbCr & vbLf & StringRetorno
                End If
        
                MsgBox Mensagem, vbOKOnly + vbInformation, TituloJanela
                Exit Function
        End If 'fim do ST1 <> 0 and ST2 <> 0
        
        If RetornaMensagem Then
            Mensagem = label & RetornoFuncao
        End If
        
        If Mensagem <> "" Then
            MsgBox Mensagem, vbOKOnly + vbInformation, TituloJanela
        End If
        Exit Function
    ElseIf Retorno = -1 Then
        MsgBox "Erro de execução da função.", vbOKOnly + vbCritical, TituloJanela
        Exit Function
    ElseIf Retorno = -2 Then
        MsgBox "Parâmetro inválido na função.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    ElseIf Retorno = -3 Then
        MsgBox "Alíquota não programada.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    ElseIf Retorno = -4 Then
        MsgBox "O arquivo de inicialização BemaFI32.ini não foi encontrado no diretório default. " + vbCr + "Por favor, copie esse arquivo para o diretório de sistema do Windows." + vbCr + "Se for o Windows 95 ou 98 é o diretório 'System' se for o Windows NT é o diretório 'System32'.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    ElseIf Retorno = -5 Then
        MsgBox "Erro ao abrir a porta de comunicação.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    ElseIf Retorno = -6 Then
        MsgBox "Impressora desligada ou cabo de comunicação desconectado.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    ElseIf Retorno = -7 Then
        MsgBox "Banco não encontrado no arquivo BemaFI32.ini.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    ElseIf Retorno = -8 Then
        MsgBox "Erro ao criar ou gravar no arquivo status.txt ou retorno.txt.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    End If
   
End Function

