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

''''Public Sub InserirConfiguracaoMesa(lngCONFIGURACAOID As Long, _
''''                                   strVALORINICIAL As String, strVALORFINAL As String, strVALORSEPARACAO As String, _
''''                                   strDATAINICIAL As String, _
''''                                   strDATAFINAL As String, _
''''                                   strDATAMASCARA As String, _
''''                                   strHORAINICIAL As String, _
''''                                   strHORAFINAL As String, _
''''                                   strHORAMASCARA As String, _
''''                                   strDURACAOINICIAL As String, _
''''                                   strDURACAOFINAL As String, _
''''                                   strDURACAOMASCARA As String, _
''''                                   strNUMEROINICIAL As String, _
''''                                   strNUMEROFINAL As String, _
''''                                   strRAMALINICIAL As String, _
''''                                   strRAMALFINAL As String, _
''''                                   strREGIAOINICIAL As String, _
''''                                   strREGIAOFINAL As String, _
''''                                   strDESCRICAOINICIAL As String, _
''''                                   strDESCRICAOFINAL As String, _
''''                                   strINTERURBANOINICIAL As String, _
''''                                   strINTERURBANOFINAL As String, _
''''                                   strINTERURBANOSEPARACAO As String, _
''''                                   strCABECALHO As String)
''''
''''
''''
''''  On Error GoTo trata
''''  '
''''  Dim strSql As String
''''  strSql = "INSERT INTO CONFIGURACAOMESA (" & _
''''    "CONFIGURACAOID, VALORINICIAL, VALORFINAL, VALORSEPARACAO, DATAINICIAL, DATAFINAL, DATAMASCARA, HORAINICIAL " & _
''''    ", HORAFINAL, HORAMASCARA, DURACAOINICIAL, DURACAOFINAL, DURACAOMASCARA, NUMEROINICIAL, NUMEROFINAL " & _
''''    ", RAMALINICIAL, RAMALFINAL, REGIAOINICIAL, REGIAOFINAL, DESCRICAOINICIAL, DESCRICAOFINAL, INTERURBANOINICIAL, INTERURBANOFINAL, INTERURBANOSEPARACAO, CABECALHO) VALUES("
''''
''''  strSql = strSql & Formata_Dados(lngCONFIGURACAOID, tpDados_Longo) & _
''''    ", " & Formata_Dados(strVALORINICIAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strVALORFINAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strVALORSEPARACAO, tpDados_Texto) & _
''''    ", " & Formata_Dados(strDATAINICIAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strDATAFINAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strDATAMASCARA, tpDados_Texto) & _
''''    ", " & Formata_Dados(strHORAINICIAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strHORAFINAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strHORAMASCARA, tpDados_Texto) & _
''''    ", " & Formata_Dados(strDURACAOINICIAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strDURACAOFINAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strDURACAOMASCARA, tpDados_Texto) & _
''''    ", " & Formata_Dados(strNUMEROINICIAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strNUMEROFINAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strRAMALINICIAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strRAMALFINAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strREGIAOINICIAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strREGIAOFINAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strDESCRICAOINICIAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strDESCRICAOFINAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strINTERURBANOINICIAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strINTERURBANOFINAL, tpDados_Longo) & _
''''    ", " & Formata_Dados(strINTERURBANOSEPARACAO, tpDados_Texto) & _
''''    ", " & Formata_Dados(strCABECALHO, tpDados_Boolean) & ")"
''''  '
''''  RunSP strSql
''''  '
''''  Exit Sub
''''trata:
''''  Err.Raise Err.Number, _
''''            Err.Source & ".[clsConfiguracao.InserirConfiguracaoMesa]", _
''''            Err.Description
''''End Sub
''''
''''Public Sub AlterarConfiguracaoMesa(lngCONFIGURACAOID As Long, _
''''                                   strVALORINICIAL As String, strVALORFINAL As String, strVALORSEPARACAO As String, _
''''                                   strDATAINICIAL As String, _
''''                                   strDATAFINAL As String, _
''''                                   strDATAMASCARA As String, _
''''                                   strHORAINICIAL As String, _
''''                                   strHORAFINAL As String, _
''''                                   strHORAMASCARA As String, _
''''                                   strDURACAOINICIAL As String, _
''''                                   strDURACAOFINAL As String, _
''''                                   strDURACAOMASCARA As String, _
''''                                   strNUMEROINICIAL As String, _
''''                                   strNUMEROFINAL As String, _
''''                                   strRAMALINICIAL As String, _
''''                                   strRAMALFINAL As String, _
''''                                   strREGIAOINICIAL As String, _
''''                                   strREGIAOFINAL As String, _
''''                                   strDESCRICAOINICIAL As String, _
''''                                   strDESCRICAOFINAL As String, _
''''                                   strINTERURBANOINICIAL As String, _
''''                                   strINTERURBANOFINAL As String, _
''''                                   strINTERURBANOSEPARACAO As String, _
''''                                   strCABECALHO As String)
''''
''''
''''
''''  On Error GoTo trata
''''  '
''''  Dim strSql As String
''''  strSql = "UPDATE CONFIGURACAOMESA SET " & _
''''    "VALORINICIAL = " & Formata_Dados(strVALORINICIAL, tpDados_Longo) & _
''''    ", VALORFINAL = " & Formata_Dados(strVALORFINAL, tpDados_Longo) & _
''''    ", VALORSEPARACAO = " & Formata_Dados(strVALORSEPARACAO, tpDados_Texto) & _
''''    ", DATAINICIAL = " & Formata_Dados(strDATAINICIAL, tpDados_Longo) & _
''''    ", DATAFINAL = " & Formata_Dados(strDATAFINAL, tpDados_Longo) & _
''''    ", DATAMASCARA = " & Formata_Dados(strDATAMASCARA, tpDados_Texto) & _
''''    ", HORAINICIAL = " & Formata_Dados(strHORAINICIAL, tpDados_Longo) & _
''''    ", HORAFINAL = " & Formata_Dados(strHORAFINAL, tpDados_Longo) & _
''''    ", HORAMASCARA = " & Formata_Dados(strHORAMASCARA, tpDados_Texto) & _
''''    ", DURACAOINICIAL = " & Formata_Dados(strDURACAOINICIAL, tpDados_Longo) & _
''''    ", DURACAOFINAL = " & Formata_Dados(strDURACAOFINAL, tpDados_Longo) & _
''''    ", DURACAOMASCARA = " & Formata_Dados(strDURACAOMASCARA, tpDados_Texto) & _
''''    ", NUMEROINICIAL = " & Formata_Dados(strNUMEROINICIAL, tpDados_Longo) & _
''''    ", NUMEROFINAL = " & Formata_Dados(strNUMEROFINAL, tpDados_Longo) & _
''''    ", RAMALINICIAL = " & Formata_Dados(strRAMALINICIAL, tpDados_Longo) & _
''''    ", RAMALFINAL = " & Formata_Dados(strRAMALFINAL, tpDados_Longo) & _
''''    ", REGIAOINICIAL = " & Formata_Dados(strREGIAOINICIAL, tpDados_Longo) & _
''''    ", REGIAOFINAL = " & Formata_Dados(strREGIAOFINAL, tpDados_Longo) & _
''''    ", DESCRICAOINICIAL = " & Formata_Dados(strDESCRICAOINICIAL, tpDados_Longo) & _
''''    ", DESCRICAOFINAL = " & Formata_Dados(strDESCRICAOFINAL, tpDados_Longo) & _
''''    ", INTERURBANOINICIAL = " & Formata_Dados(strINTERURBANOINICIAL, tpDados_Longo) & _
''''    ", INTERURBANOFINAL = " & Formata_Dados(strINTERURBANOFINAL, tpDados_Longo) & _
''''    ", INTERURBANOSEPARACAO = " & Formata_Dados(strINTERURBANOSEPARACAO, tpDados_Texto) & _
''''    ", CABECALHO = " & Formata_Dados(strCABECALHO, tpDados_Boolean)
''''
''''
''''    strSql = strSql & " WHERE CONFIGURACAOID = " & lngCONFIGURACAOID
''''  '
''''  RunSP strSql
''''  '
''''  Exit Sub
''''trata:
''''  Err.Raise Err.Number, _
''''            Err.Source & ".[clsConfiguracao.AlterarConfiguracaoMesa]", _
''''            Err.Description
''''End Sub
''''
''''Public Function ListarConfiguracaoMesa(lngCONFIGURACAOID As Long) As ADODB.Recordset
''''  On Error GoTo trata
''''  '
''''  Dim strSql As String
''''
''''  strSql = "SELECT CONFIGURACAOMESA.* " & _
''''    "FROM CONFIGURACAOMESA " & _
''''    "WHERE CONFIGURACAOMESA.CONFIGURACAOID = " & _
''''      Formata_Dados(lngCONFIGURACAOID, tpDados_Longo)
''''  Set ListarConfiguracaoMesa = RunSPReturnRS(strSql)
''''  '
''''  Exit Function
''''trata:
''''  Err.Raise Err.Number, _
''''            Err.Source & ".[clsConfiguracao.ListarConfiguracaoMesa]", _
''''            Err.Description
''''End Function
''''
''''Public Sub ExcluirConfiguracaoMesa(lngCONFIGURACAOID As Long)
''''  On Error GoTo trata
''''  '
''''  Dim strSql As String
''''  'exclui CONFIGURACAOMESA
''''  strSql = "DELETE FROM CONFIGURACAOMESA WHERE CONFIGURACAOMESA.CONFIGURACAOID = " & _
''''    Formata_Dados(lngCONFIGURACAOID, tpDados_Longo)
''''  RunSP strSql
''''  '
''''  Exit Sub
''''trata:
''''  Err.Raise Err.Number, _
''''            Err.Source & ".[clsConfiguracao.ExcluirConfiguracaoMesa]", _
''''            Err.Description
''''End Sub


Public Sub ExcluirConfiguracao(lngCONFIGURACAOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui CONFIGURACAO
  strSql = "DELETE FROM CONFIGURACAO WHERE CONFIGURACAO.PKID = " & _
    Formata_Dados(lngCONFIGURACAOID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConfiguracao.ExcluirConfiguracao]", _
            Err.Description
End Sub

Public Function VerificaExclusaoConfiguracao(lngCONFIGURACAOID As Long, _
                                             ByRef strDescricaoErro As String) As Boolean
  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  Dim blnRet As Boolean
  strDescricaoErro = ""
  blnRet = True
'''  'LOCAÇÃO
'''  strSql = "SELECT COUNT(*) AS TOTALLANC FROM LOCACAO WHERE CONFIGURACAOID = " & _
'''    Formata_Dados(lngCONFIGURACAOID, tpDados_Longo)
'''  Set objRs = RunSPReturnRS(strSql)
'''  If objRs.EOF Then
'''    blnRet = True
'''  Else
'''    If objRs.Fields("TOTALLANC").Value = 0 Then
'''      blnRet = True
'''    Else
'''      blnRet = False
'''      strDescricaoErro = "Tabela de Locação"
'''    End If
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
'''  If blnRet = True Then
'''    'VENDA
'''    strSql = "SELECT COUNT(*) AS TOTALLANC FROM VENDA WHERE CONFIGURACAOID = " & _
'''      Formata_Dados(lngCONFIGURACAOID, tpDados_Longo)
'''    Set objRs = RunSPReturnRS(strSql)
'''    If objRs.EOF Then
'''      blnRet = True
'''    Else
'''      If objRs.Fields("TOTALLANC").Value = 0 Then
'''        blnRet = True
'''      Else
'''        blnRet = False
'''        strDescricaoErro = "Tabela de Vendas"
'''      End If
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
'''  End If
  VerificaExclusaoConfiguracao = blnRet
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConfiguracao.VerificaExclusaoConfiguracao]", _
            Err.Description
End Function

''''Public Sub AlterarConfiguracaoFechamento(lngCONFIGURACAOID As Long, _
''''                                        intTipo As Integer, _
''''                                        intTipoFechamento As Integer, _
''''                                        blnTrabComFechaCC As Boolean)
''''  On Error GoTo trata
''''  '
''''  Dim strSql As String
''''  strSql = "UPDATE CONFIGURACAO SET " & _
''''    "TIPO = " & Formata_Dados(intTipo, tpDados_Longo) & _
''''    ", TIPOFECHAMENTO = " & Formata_Dados(intTipoFechamento, tpDados_Longo) & _
''''    ", TRABCOMFECHACC = " & Formata_Dados(blnTrabComFechaCC, tpDados_Boolean) & _
''''    " WHERE PKID = " & lngCONFIGURACAOID
''''
''''  RunSP strSql
''''  '
''''  Exit Sub
''''trata:
''''  Err.Raise Err.Number, _
''''            Err.Source & ".[clsConfiguracao.AlterarConfiguracaoFechamento]", _
''''            Err.Description
''''End Sub
Public Sub AlterarConfiguracaoGeral(lngCONFIGURACAOID As Long, _
                                    strCaminho As String, _
                                    strQtdDiasVenda As String)
  On Error GoTo trata
  '
  Dim strSql As String
  strSql = "UPDATE CONFIGURACAO SET " & _
    "CAMINHOIMAGEMCOMPRA = " & Formata_Dados(strCaminho, tpDados_Texto) & _
    ", QTDDIASVENDA = " & Formata_Dados(strQtdDiasVenda, tpDados_Longo) & _
    " WHERE PKID = " & lngCONFIGURACAOID
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConfiguracao.AlterarConfiguracaoGeral]", _
            Err.Description
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
''''  On Error GoTo trata
''''  '
''''  Dim strSql As String
''''  strSql = "UPDATE CONFIGURACAO SET " & _
''''    "TRABSUITEAPTOLIMPO = " & Formata_Dados(blnSuiteAptoLimpo, tpDados_Boolean) & ", TRABCOMLIBERACAO = " & Formata_Dados(blnTrabComLiberacao, tpDados_Boolean) & ", TRABSAIDA = " & Formata_Dados(blnTrabSaida, tpDados_Boolean) & ", TRABDEPASSOCENTRA = " & Formata_Dados(blnTrabDepAssocEntra, tpDados_Boolean) & _
''''    ", DEPOSITOOBRIGATORIO = " & Formata_Dados(blnDepositoObriga, tpDados_Boolean) & ", TRABCOMEXTRAUNIDADE = " & Formata_Dados(blnTrabExtraUnidade, tpDados_Boolean) & ", TRABOCUPANTES = " & Formata_Dados(blnTrabOcupantesveiculo, tpDados_Boolean) & ", LIBERARHRENTRADA = " & Formata_Dados(blnLiberarHrEntrada, tpDados_Boolean) & _
''''    ", TRABCAMPOHOSPEDE = " & Formata_Dados(blnTrabCampoHospede, tpDados_Boolean) & ", TRABCOMENTRADAPES = " & Formata_Dados(blnTrabComEntradaPessoa, tpDados_Boolean) & ", TRABCOMEXEDPESSOA = " & Formata_Dados(blnTrabcomexedPessoa, tpDados_Boolean) & ", TRABCOMTROCO = " & Formata_Dados(blnTrabComTroco, tpDados_Boolean) & _
''''    ", TAMANHONROUNIDADE = " & Formata_Dados(strCompr, tpDados_Texto) & ", DTINICIOENTRADA = " & Formata_Dados(strHoraEntra, tpDados_DataHora) & ", PORTCANCELAENTRADA = " & Formata_Dados(blnPortCancEntra, tpDados_Boolean) & ", RECACESSADADOSFINANC = " & Formata_Dados(blnRecAcessaFinanc, tpDados_Boolean) & _
''''    ", LIBFECHACOMFECHASEMREC = " & Formata_Dados(blnLiberaFechaComFechaSemRec, tpDados_Boolean) & ", PEDIRSENHADESCCORTFECHA = " & Formata_Dados(blnPedirSenhaDescCortFecha, tpDados_Boolean) & ", PEDIRSENHASUPAPOS2FECHA = " & Formata_Dados(blnPedirSenhaSupApos2Fecha, tpDados_Boolean) & _
''''    ", MINTOLERANCIADIARIA = " & Formata_Dados(strMinTolera, tpDados_Longo) & _
''''    ", QTDMINPARARECEB = " & Formata_Dados(strMinReceb, tpDados_Longo) & _
''''    ", QTDMINTRANSFSUITEMENVR = " & Formata_Dados(strMinTransfSuiteMenVr, tpDados_Longo) & _
''''    ", TRABCOMMESATEL = " & Formata_Dados(blnTrabComMesaTel, tpDados_Boolean) & _
''''    ", CAMINHO = " & Formata_Dados(strCaminho, tpDados_Texto) & _
''''    ", QTDDIASPARACOMPENSAR = " & Formata_Dados(strQtdDiasParaCompensar, tpDados_Longo) & _
''''    ", QTDCHEQUESBONS = " & Formata_Dados(strQtdChequesBons, tpDados_Longo) & _
''''    ", PEDIRSENHASUPLIBCHQRECEB = " & Formata_Dados(blnPedirSenhaSupLibChqReceb, tpDados_Boolean) & _
''''    ", TRABCOMCHEQUESBONS = " & Formata_Dados(blnTrabComChequesBons, tpDados_Boolean) & _
''''    ", TRABCOMTAXA = " & Formata_Dados(strTrabComTaxa, tpDados_Texto) & _
''''    ", PERCTAXA = " & Formata_Dados(strPercTaxa, tpDados_Moeda) & _
''''    ", PEDIRSENHASUPDESPTEL = " & Formata_Dados(blnPedirSenhaSupDespesaTelefonia, tpDados_Boolean) & _
''''    ", PEDIRSENHASUPCANCPED = " & Formata_Dados(blnPedirSenhaSupCancPed, tpDados_Boolean) & _
''''    ", PEDIRSENHASUPCLIPROB = " & Formata_Dados(blnPedirSenhaSupCliProb, tpDados_Boolean) & _
''''    ", PEDIRSENHASUPTRANSFUNID = " & Formata_Dados(blnPedirSenhaSupTransfUnid, tpDados_Boolean) & _
''''    ", PEDIRSENHASUPSANGRIA = " & Formata_Dados(blnPedirSenhaSupSangria, tpDados_Boolean) & _
''''    ", TRABPGTOFATURA = " & Formata_Dados(blnTrabComFatura, tpDados_Boolean) & _
''''    ", PEDIRSENHASUPEXCDEP = " & Formata_Dados(blnPedirSenhaSupExcDep, tpDados_Boolean) & _
''''    " WHERE PKID = " & lngCONFIGURACAOID
''''  '
''''  RunSP strSql
''''  '
''''  Exit Sub
''''trata:
''''  Err.Raise Err.Number, _
''''            Err.Source & ".[clsConfiguracao.AlterarConfiguracaoLocacao]", _
''''            Err.Description
''''End Sub
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
  On Error GoTo trata
  '
  Dim strSql As String
  strSql = "UPDATE CONFIGURACAO SET " & _
    "Empresa = " & Formata_Dados(strEmpresa, tpDados_Texto) & _
    ", Cnpj = " & Formata_Dados(strCnpj, tpDados_Texto) & _
    ", InscrMunicipal = " & Formata_Dados(strInscrMunicipal, tpDados_Texto) & _
    ", Endereco = " & Formata_Dados(strEndereco, tpDados_Texto) & _
    ", Bairro = " & Formata_Dados(strBairro, tpDados_Texto) & _
    ", Cidade = " & Formata_Dados(strCidade, tpDados_Texto) & _
    ", Estado = " & Formata_Dados(strEstado, tpDados_Texto) & _
    ", Cep = " & Formata_Dados(strCep, tpDados_Texto) & _
    ", Tel = " & Formata_Dados(strTel, tpDados_Texto) & _
    ", titulo = " & Formata_Dados(strTitulo, tpDados_Texto) & _
    " WHERE PKID = " & lngCONFIGURACAOID
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConfiguracao.AlterarConfiguracaoCadastro]", _
            Err.Description
End Sub

Public Sub InserirConfiguracao()
  On Error GoTo trata
  '
  Dim strSql As String
  strSql = "INSERT INTO CONFIGURACAO (TITULO, Empresa, Endereco, Bairro, Cidade, Estado, Tel, Cep, Avi_Papel, Cnpj, InscrMunicipal, QTDDIASVENDA) " & _
    " VALUES ('TITULO', 'XXX', Null, Null, Null, Null, Null, Null, Null, Null, Null, 1) "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConfiguracao.InserirConfiguracao]", _
            Err.Description
End Sub
''''
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
''''  On Error GoTo trata
''''  '
''''  Dim strSql As String
''''  strSql = "UPDATE CONFIGURACAO SET " & _
''''    "VRMINUTO = " & Formata_Dados(strVrTarifaMin, tpDados_Moeda) & _
''''    ", VRMINUTOINTER = " & Formata_Dados(strVrTarifaMinInter, tpDados_Moeda) & _
''''    ", VRMINUTOINTERNAC = " & Formata_Dados(strVrTarifaMinInternac, tpDados_Moeda) & _
''''    ", DDDLOCAL = " & Formata_Dados(strDDDLocal, tpDados_Texto) & _
''''    ", QTDMINDIR = " & Formata_Dados(strQtdMinDir, tpDados_Longo) & _
''''    ", QTDMINGER = " & Formata_Dados(strQtdMinGer, tpDados_Longo) & _
''''    ", QTDMINREC = " & Formata_Dados(strQtdMinRec, tpDados_Longo) & _
''''    ", QTDMINPORT = " & Formata_Dados(strQtdMinPort, tpDados_Longo) & _
''''    ", QTDMINDIRPED = " & Formata_Dados(strQtdPedMinDir, tpDados_Longo) & _
''''    ", QTDMINGERPED = " & Formata_Dados(strQtdPedMinGer, tpDados_Longo) & _
''''    ", QTDMINRECPED = " & Formata_Dados(strQtdPedMinRec, tpDados_Longo) & _
''''    " WHERE PKID = " & lngCONFIGURACAOID
''''  '
''''  RunSP strSql
''''  '
''''  Exit Sub
''''trata:
''''  Err.Raise Err.Number, _
''''            Err.Source & ".[clsConfiguracao.AlterarConfiguracaoTelEntrPed]", _
''''            Err.Description
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
''''  On Error GoTo trata
''''  '
''''  Dim strSql As String
''''  strSql = "UPDATE CONFIGURACAO SET " & _
''''    "HABSERVDESPDIR = " & Formata_Dados(blnHabServDespDir, tpDados_Boolean) & _
''''    ", HABSERVDESPGER = " & Formata_Dados(blnHabServDespGer, tpDados_Boolean) & _
''''    ", HABSERVDESPREC = " & Formata_Dados(blnHabServDespRec, tpDados_Boolean) & _
''''    ", HABSERVDESPPORT = " & Formata_Dados(blnHabServDespPort, tpDados_Boolean) & _
''''    ", HabDiariaDir = " & Formata_Dados(blnHabDiariaDir, tpDados_Boolean) & _
''''    ", HabDiariaGer = " & Formata_Dados(blnHabDiariaGer, tpDados_Boolean) & _
''''    ", HabDiariaRec = " & Formata_Dados(blnHabDiariaRec, tpDados_Boolean) & _
''''    ", HabDiariaPort = " & Formata_Dados(blnHabDiariaPor, tpDados_Boolean) & _
''''    ", INTERVALOBASE = " & Formata_Dados(strIntervBase, tpDados_Longo) & _
''''    ", INTERVALOOPERADOR = " & Formata_Dados(strIntervOper, tpDados_Longo) & _
''''    ", HABMOVCAIXADIR = " & Formata_Dados(blnHabMovCaixaDir, tpDados_Boolean) & _
''''    ", HABMOVCAIXAGER = " & Formata_Dados(blnHabMovCaixaGer, tpDados_Boolean) & _
''''    ", HABMOVCAIXAREC = " & Formata_Dados(blnHabMovCaixaRec, tpDados_Boolean) & _
''''    ", HABMOVCAIXAPORT = " & Formata_Dados(blnHabMovCaixaPort, tpDados_Boolean) & _
''''    ", VRMOVCAIXA = " & Formata_Dados(strTetoMovCaixa, tpDados_Moeda) & _
''''    " WHERE PKID = " & lngCONFIGURACAOID
''''  '
''''  RunSP strSql
''''  '
''''  Exit Sub
''''trata:
''''  Err.Raise Err.Number, _
''''            Err.Source & ".[clsConfiguracao.AlterarConfiguracaoDespDiaCaixa]", _
''''            Err.Description
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
''''  On Error GoTo trata
''''  '
''''  Dim strSql As String
''''  strSql = "UPDATE CONFIGURACAO SET " & _
''''    "PEDIRSENHAFECHATURNO = " & Formata_Dados(blnPedirSenhaFechaTurno, tpDados_Boolean) & _
''''    ", PEDIRSENHAPEDDIRETORIA = " & Formata_Dados(blnPedirSenhaVdaDiretoria, tpDados_Boolean) & _
''''    ", ESTCADASTRACARDAPIO = " & Formata_Dados(blnEstCadastraCardapio, tpDados_Boolean) & _
''''    ", GERMODIFCONFIG = " & Formata_Dados(blnGerModifConfig, tpDados_Boolean) & _
''''    ", TRABCOMESTORNO = " & Formata_Dados(blnTrabComEstorno, tpDados_Boolean) & _
''''    ", TRABCOMDEPSAN = " & Formata_Dados(blnTrabDepSan, tpDados_Boolean) & _
''''    ", TRABCOMVENDASCOB = " & Formata_Dados(blnTrabComVendasCob, tpDados_Boolean) & _
''''    ", VRINICIALCAIXA = " & Formata_Dados(strVrInicialCaixa, tpDados_Moeda) & _
''''    ", VRISS = " & Formata_Dados(strVrISS, tpDados_Moeda) & _
''''    ", MAXSEQUENCIALIMP = " & Formata_Dados(strSequencial, tpDados_Longo) & _
''''    ", DTRELMOVIMENTO = " & Formata_Dados(strDataInicialRelMov, tpDados_DataHora) & _
''''    ", TRABCOMESTINTER = " & Formata_Dados(blnTrabComEstInter, tpDados_Boolean) & _
''''    ", DIAFECHAFOLHA = " & Formata_Dados(strDiaFechaFolha, tpDados_Longo) & _
''''    ", TRABCOMESTNEGATIVO = " & Formata_Dados(blnTrabComEstNegativo, tpDados_Boolean) & _
''''    ", TRABCOMEXCPENHOR = " & Formata_Dados(blnTrabComExcPenhor, tpDados_Boolean) & _
''''    ", TRABALTVENDA = " & Formata_Dados(blnTrabComAltVda, tpDados_Boolean) & _
''''    " WHERE PKID = " & lngCONFIGURACAOID
''''  '
''''  RunSP strSql
''''  '
''''  Exit Sub
''''trata:
''''  Err.Raise Err.Number, _
''''            Err.Source & ".[clsConfiguracao.AlterarConfiguracaoDiversos]", _
''''            Err.Description
''''End Sub
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
''''  On Error GoTo trata
''''  '
''''  Dim strSql As String
''''  strSql = "UPDATE CONFIGURACAO SET " & _
''''    "IMPVENDASFECHATUR = " & Formata_Dados(blnImprimirVendasFechaTurno, tpDados_Boolean) & _
''''    ", IMPCARTAOFECHATUR = " & Formata_Dados(blnImprimirCartaoFechaTur, tpDados_Boolean) & _
''''    ", IMPPGTOFECHATUR = " & Formata_Dados(blnImprimirPagamentos, tpDados_Boolean) & _
''''    ", IMPCTAFECHATUR = " & Formata_Dados(blnImprimirContas, tpDados_Boolean) & _
''''    ", IMPRIMIRPEDFRIGO = " & Formata_Dados(blnImprimirPedFrigo, tpDados_Boolean) & _
''''    ", IMPRIMIRCOMPROVLOC = " & Formata_Dados(blnImprimirComprovLoc, tpDados_Boolean) & _
''''    ", TRABCOMIMPFISCAL = " & Formata_Dados(blnTrabComImpFiscal, tpDados_Boolean) & _
''''    ", IMPRECEBETURNOSANT = " & Formata_Dados(blnImpRecebeTurnosAnt, tpDados_Boolean) & _
''''    ", MOSTRARCABEC = " & Formata_Dados(blnImprimirCabec, tpDados_Boolean) & _
''''    ", PONTOSIMPBOL = " & Formata_Dados(strPontosBol, tpDados_Longo) & _
''''    ", TIPOIMPFISCAL = " & Formata_Dados(strTipoImpFiscal, tpDados_Texto) & _
''''    ", TRABTABDATA = " & Formata_Dados(blnTrabTabData, tpDados_Boolean) & _
''''    ", IMPTURFECHA = " & Formata_Dados(blnImprimirTurnoFecha, tpDados_Boolean) & _
''''    " WHERE PKID = " & lngCONFIGURACAOID
''''  '
''''  RunSP strSql
''''  '
''''  Exit Sub
''''trata:
''''  Err.Raise Err.Number, _
''''            Err.Source & ".[clsConfiguracao.AlterarConfiguracaoImpressao]", _
''''            Err.Description
''''End Sub
''''Public Function ListarConfiguracaoLocacao(lngCONFIGURACAOID As Long) As ADODB.Recordset
''''  On Error GoTo trata
''''  '
''''  Dim strSql As String
''''
''''  strSql = "SELECT CONFIGURACAO.*, CONFIGURACAOMESA.* " & _
''''    "FROM CONFIGURACAO " & _
''''    " LEFT JOIN CONFIGURACAOMESA ON CONFIGURACAO.PKID = CONFIGURACAOMESA.CONFIGURACAOID " & _
''''    "WHERE CONFIGURACAO.PKID = " & _
''''      Formata_Dados(lngCONFIGURACAOID, tpDados_Longo, tpNulo_NaoAceita)
''''  Set ListarConfiguracaoLocacao = RunSPReturnRS(strSql)
''''  '
''''  Exit Function
''''trata:
''''  Err.Raise Err.Number, _
''''            Err.Source & ".[clsConfiguracao.ListarConfiguracaoLocacao]", _
''''            Err.Description
''''End Function
''''
Public Function ListarConfiguracao() As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT CONFIGURACAO.* " & _
    "FROM CONFIGURACAO "
  Set ListarConfiguracao = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConfiguracao.ListarConfiguracao]", _
            Err.Description
End Function

Public Function ListarConfiguracaoCadastro(lngCONFIGURACAOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT CONFIGURACAO.* " & _
    "FROM CONFIGURACAO " & _
    "WHERE CONFIGURACAO.PKID = " & _
      Formata_Dados(lngCONFIGURACAOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarConfiguracaoCadastro = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConfiguracao.ListarConfiguracaoCadastro]", _
            Err.Description
End Function

''''Public Function ListarConfiguracaoTelEntrPed(lngCONFIGURACAOID As Long) As ADODB.Recordset
''''  On Error GoTo trata
''''  '
''''  Dim strSql As String
''''
''''  strSql = "SELECT CONFIGURACAO.* " & _
''''    "FROM CONFIGURACAO " & _
''''    "WHERE CONFIGURACAO.PKID = " & _
''''      Formata_Dados(lngCONFIGURACAOID, tpDados_Longo, tpNulo_NaoAceita)
''''  Set ListarConfiguracaoTelEntrPed = RunSPReturnRS(strSql)
''''  '
''''  Exit Function
''''trata:
''''  Err.Raise Err.Number, _
''''            Err.Source & ".[clsConfiguracao.ListarConfiguracaoTelEntrPed]", _
''''            Err.Description
''''End Function
''''
''''Public Function ListarConfiguracaoDespDiaCaixa(lngCONFIGURACAOID As Long) As ADODB.Recordset
''''  On Error GoTo trata
''''  '
''''  Dim strSql As String
''''
''''  strSql = "SELECT CONFIGURACAO.* " & _
''''    "FROM CONFIGURACAO " & _
''''    "WHERE CONFIGURACAO.PKID = " & _
''''      Formata_Dados(lngCONFIGURACAOID, tpDados_Longo, tpNulo_NaoAceita)
''''  Set ListarConfiguracaoDespDiaCaixa = RunSPReturnRS(strSql)
''''  '
''''  Exit Function
''''trata:
''''  Err.Raise Err.Number, _
''''            Err.Source & ".[clsConfiguracao.ListarConfiguracaoDespDiaCaixa]", _
''''            Err.Description
''''End Function
''''
''''Public Function ListarConfiguracaoDiversos(lngCONFIGURACAOID As Long) As ADODB.Recordset
''''  On Error GoTo trata
''''  '
''''  Dim strSql As String
''''
''''  strSql = "SELECT CONFIGURACAO.* " & _
''''    "FROM CONFIGURACAO " & _
''''    "WHERE CONFIGURACAO.PKID = " & _
''''      Formata_Dados(lngCONFIGURACAOID, tpDados_Longo, tpNulo_NaoAceita)
''''  Set ListarConfiguracaoDiversos = RunSPReturnRS(strSql)
''''  '
''''  Exit Function
''''trata:
''''  Err.Raise Err.Number, _
''''            Err.Source & ".[clsConfiguracao.ListarConfiguracaoDiversos]", _
''''            Err.Description
''''End Function
''''
''''Public Function ListarConfiguracaoImpressao(lngCONFIGURACAOID As Long) As ADODB.Recordset
''''  On Error GoTo trata
''''  '
''''  Dim strSql As String
''''
''''  strSql = "SELECT CONFIGURACAO.* " & _
''''    "FROM CONFIGURACAO " & _
''''    "WHERE CONFIGURACAO.PKID = " & _
''''      Formata_Dados(lngCONFIGURACAOID, tpDados_Longo, tpNulo_NaoAceita)
''''  Set ListarConfiguracaoImpressao = RunSPReturnRS(strSql)
''''  '
''''  Exit Function
''''trata:
''''  Err.Raise Err.Number, _
''''            Err.Source & ".[clsConfiguracao.ListarConfiguracaoImpressao]", _
''''            Err.Description
''''End Function
Public Function ListarConfiguracaoGeral(lngCONFIGURACAOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT CONFIGURACAO.* " & _
    "FROM CONFIGURACAO " & _
    "WHERE CONFIGURACAO.PKID = " & _
      Formata_Dados(lngCONFIGURACAOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarConfiguracaoGeral = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConfiguracao.ListarConfiguracaoGeral]", _
            Err.Description
End Function


''''Public Function ListarConfiguracaoFechamento(lngCONFIGURACAOID As Long) As ADODB.Recordset
''''  On Error GoTo trata
''''  '
''''  Dim strSql As String
''''
''''  strSql = "SELECT CONFIGURACAO.* " & _
''''    "FROM CONFIGURACAO " & _
''''    "WHERE CONFIGURACAO.PKID = " & _
''''      Formata_Dados(lngCONFIGURACAOID, tpDados_Longo, tpNulo_NaoAceita)
''''  Set ListarConfiguracaoFechamento = RunSPReturnRS(strSql)
''''  '
''''  Exit Function
''''trata:
''''  Err.Raise Err.Number, _
''''            Err.Source & ".[clsConfiguracao.ListarConfiguracaoFechamento]", _
''''            Err.Description
''''End Function
''''
