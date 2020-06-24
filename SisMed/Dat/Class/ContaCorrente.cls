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

Public Sub AlterarFinanceiro(lngFINANCEIROID As Long, _
                             strVRPRESTCONSESPECIE As String, strVRDONORXCONSESPECIE As String, strVRTECRXCONSESPECIE As String, strVRDONOULTRACONSESPECIE As String, _
                             strVRPRESTCONSCARTAO As String, strVRDONORXCONSCARTAO As String, strVRTECRXCONSCARTAO As String, strVRDONOULTRACONSCARTAO As String, strVRPRESTCONSCONVENIO As String, strVRDONORXCONSCONVENIO As String, strVRTECRXCONSCONVENIO As String, strVRDONOULTRACONSCONVENIO As String, _
                             strVRPRESTESPECIE As String, strVRDONORXESPECIE As String, strVRTECRXESPECIE As String, strVRDONOULTRAESPECIE As String, strVRPRESTCARTAO As String, strVRDONORXCARTAO As String, strVRTECRXCARTAO As String, strVRDONOULTRACARTAO As String, strVRPRESTCONVENIO As String, _
                             strVRDONORXCONVENIO As String, strVRTECRXCONVENIO As String, strVRDONOULTRACONVENIO As String, strVRPRESTTOTAL As String, strVRDONORXTOTAL As String, strVRTECRXTOTAL As String, strVRDONOULTRATOTAL As String, strVRTOTAL As String, strVRCASA As String, strVRTOTALGERAL As String, _
                             strVRPRESTCARTAOAREC As String, strVRDONORXCARTAOAREC As String, strVRTECRXCARTAOAREC As String, strVRDONOULTRACARTAOAREC As String, strVRTOTALAREC As String, strDATAAREC As String)
  
  On Error GoTo trata
  '
  Dim strSql As String
      
  strSql = "UPDATE FINANCEIRO SET " & _
    "VRPRESTCONSESPECIE = " & Formata_Dados(strVRPRESTCONSESPECIE, tpDados_Moeda) & _
    "VRDONORXCONSESPECIE = " & Formata_Dados(strVRDONORXCONSESPECIE, tpDados_Moeda) & _
    "VRTECRXCONSESPECIE = " & Formata_Dados(strVRTECRXCONSESPECIE, tpDados_Moeda) & _
    "VRDONOULTRACONSESPECIE = " & Formata_Dados(strVRDONOULTRACONSESPECIE, tpDados_Moeda) & _
    "VRPRESTCONSCARTAO = " & Formata_Dados(strVRPRESTCONSCARTAO, tpDados_Moeda) & _
    "VRDONORXCONSCARTAO = " & Formata_Dados(strVRDONORXCONSCARTAO, tpDados_Moeda) & _
    "VRTECRXCONSCARTAO = " & Formata_Dados(strVRTECRXCONSCARTAO, tpDados_Moeda) & _
    "VRDONOULTRACONSCARTAO = " & Formata_Dados(strVRDONOULTRACONSCARTAO, tpDados_Moeda) & _
    "VRPRESTCONSCONVENIO = " & Formata_Dados(strVRPRESTCONSCONVENIO, tpDados_Moeda) & _
    "VRDONORXCONSCONVENIO = " & Formata_Dados(strVRDONORXCONSCONVENIO, tpDados_Moeda) & _
    "VRTECRXCONSCONVENIO = " & Formata_Dados(strVRTECRXCONSCONVENIO, tpDados_Moeda) & _
    "VRDONOULTRACONSCONVENIO = " & Formata_Dados(strVRDONOULTRACONSCONVENIO, tpDados_Moeda)

  strSql = strSql & _
    "VRPRESTESPECIE = " & Formata_Dados(strVRPRESTESPECIE, tpDados_Moeda) & _
    "VRDONORXESPECIE = " & Formata_Dados(strVRDONORXESPECIE, tpDados_Moeda) & _
    "VRTECRXESPECIE = " & Formata_Dados(strVRTECRXESPECIE, tpDados_Moeda) & _
    "VRDONOULTRAESPECIE = " & Formata_Dados(strVRDONOULTRAESPECIE, tpDados_Moeda) & _
    "VRPRESTCARTAO = " & Formata_Dados(strVRPRESTCARTAO, tpDados_Moeda) & _
    "VRDONORXCARTAO = " & Formata_Dados(strVRDONORXCARTAO, tpDados_Moeda) & _
    "VRTECRXCARTAO = " & Formata_Dados(strVRTECRXCARTAO, tpDados_Moeda) & _
    "VRDONOULTRACARTAO = " & Formata_Dados(strVRDONOULTRACARTAO, tpDados_Moeda) & _
    "VRPRESTCONVENIO = " & Formata_Dados(strVRPRESTCONVENIO, tpDados_Moeda) & _
    "VRDONORXCONVENIO = " & Formata_Dados(strVRDONORXCONVENIO, tpDados_Moeda) & _
    "VRTECRXCONVENIO = " & Formata_Dados(strVRTECRXCONVENIO, tpDados_Moeda) & _
    "VRDONOULTRACONVENIO = " & Formata_Dados(strVRDONOULTRACONVENIO, tpDados_Moeda) & _
    "VRPRESTTOTAL = " & Formata_Dados(strVRPRESTTOTAL, tpDados_Moeda) & _
    "VRDONORXTOTAL = " & Formata_Dados(strVRDONORXTOTAL, tpDados_Moeda)

  strSql = strSql & _
    "VRTECRXTOTAL = " & Formata_Dados(strVRTECRXTOTAL, tpDados_Moeda) & _
    "VRDONOULTRATOTAL = " & Formata_Dados(strVRDONOULTRATOTAL, tpDados_Moeda) & _
    "VRTOTAL = " & Formata_Dados(strVRTOTAL, tpDados_Moeda) & _
    "VRCASA = " & Formata_Dados(strVRCASA, tpDados_Moeda) & _
    "VRTOTALGERAL = " & Formata_Dados(strVRTOTALGERAL, tpDados_Moeda) & _
    "VRPRESTCARTAOAREC = " & Formata_Dados(strVRPRESTCARTAOAREC, tpDados_Moeda) & _
    "VRDONORXCARTAOAREC = " & Formata_Dados(strVRDONORXCARTAOAREC, tpDados_Moeda) & _
    "VRTECRXCARTAOAREC = " & Formata_Dados(strVRTECRXCARTAOAREC, tpDados_Moeda) & _
    "VRDONOULTRACARTAOAREC = " & Formata_Dados(strVRDONOULTRACARTAOAREC, tpDados_Moeda) & _
    "VRTOTALAREC = " & Formata_Dados(strVRTOTALAREC, tpDados_Moeda) & _
    "DATAAREC = " & Formata_Dados(strDATAAREC, tpDados_DataHora) & _
    " WHERE PKID = " & Formata_Dados(lngFINANCEIROID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContaCorrente.AlterarFinanceiro]", _
            Err.Description
End Sub

Public Sub InserirFinanceiro(lngGRID As Long, _
                             strVRPRESTCONSESPECIE As String, strVRDONORXCONSESPECIE As String, strVRTECRXCONSESPECIE As String, strVRDONOULTRACONSESPECIE As String, _
                             strVRPRESTCONSCARTAO As String, strVRDONORXCONSCARTAO As String, strVRTECRXCONSCARTAO As String, strVRDONOULTRACONSCARTAO As String, strVRPRESTCONSCONVENIO As String, strVRDONORXCONSCONVENIO As String, strVRTECRXCONSCONVENIO As String, strVRDONOULTRACONSCONVENIO As String, _
                             strVRPRESTESPECIE As String, strVRDONORXESPECIE As String, strVRTECRXESPECIE As String, strVRDONOULTRAESPECIE As String, strVRPRESTCARTAO As String, strVRDONORXCARTAO As String, strVRTECRXCARTAO As String, strVRDONOULTRACARTAO As String, strVRPRESTCONVENIO As String, _
                             strVRDONORXCONVENIO As String, strVRTECRXCONVENIO As String, strVRDONOULTRACONVENIO As String, strVRPRESTTOTAL As String, strVRDONORXTOTAL As String, strVRTECRXTOTAL As String, strVRDONOULTRATOTAL As String, strVRTOTAL As String, strVRCASA As String, strVRTOTALGERAL As String, _
                             strVRPRESTCARTAOAREC As String, strVRDONORXCARTAOAREC As String, strVRTECRXCARTAOAREC As String, strVRDONOULTRACARTAOAREC As String, strVRTOTALAREC As String, strDATAAREC As String)
  
  On Error GoTo trata
  '
  Dim strSql As String
      
  strSql = "INSERT INTO FINANCEIRO (VRPRESTCONSESPECIE, VRDONORXCONSESPECIE, VRTECRXCONSESPECIE, VRDONOULTRACONSESPECIE, VRPRESTCONSCARTAO, VRDONORXCONSCARTAO, VRTECRXCONSCARTAO, VRDONOULTRACONSCARTAO, VRPRESTCONSCONVENIO, VRDONORXCONSCONVENIO, VRTECRXCONSCONVENIO, VRDONOULTRACONSCONVENIO " & _
            "VRPRESTESPECIE, VRDONORXESPECIE, VRTECRXESPECIE, VRDONOULTRAESPECIE, VRPRESTCARTAO, VRDONORXCARTAO, VRTECRXCARTAO, VRDONOULTRACARTAO, VRPRESTCONVENIO, VRDONORXCONVENIO, VRTECRXCONVENIO, VRDONOULTRACONVENIO, VRPRESTTOTAL, VRDONORXTOTAL, " & _
            "VRTECRXTOTAL, VRDONOULTRATOTAL, VRTOTAL, VRCASA, VRTOTALGERAL, VRPRESTCARTAOAREC, VRDONORXCARTAOAREC, VRTECRXCARTAOAREC, VRDONOULTRACARTAOAREC, VRTOTALAREC, DATAAREC, GRID) VALUES (" & _
    "  " & Formata_Dados(strVRPRESTCONSESPECIE, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRDONORXCONSESPECIE, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRTECRXCONSESPECIE, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRDONOULTRACONSESPECIE, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRPRESTCONSCARTAO, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRDONORXCONSCARTAO, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRTECRXCONSCARTAO, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRDONOULTRACONSCARTAO, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRPRESTCONSCONVENIO, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRDONORXCONSCONVENIO, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRTECRXCONSCONVENIO, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRDONOULTRACONSCONVENIO, tpDados_Moeda)

  strSql = strSql & _
    " , " & Formata_Dados(strVRPRESTESPECIE, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRDONORXESPECIE, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRTECRXESPECIE, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRDONOULTRAESPECIE, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRPRESTCARTAO, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRDONORXCARTAO, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRTECRXCARTAO, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRDONOULTRACARTAO, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRPRESTCONVENIO, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRDONORXCONVENIO, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRTECRXCONVENIO, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRDONOULTRACONVENIO, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRPRESTTOTAL, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRDONORXTOTAL, tpDados_Moeda)

  strSql = strSql & _
    " , " & Formata_Dados(strVRTECRXTOTAL, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRDONOULTRATOTAL, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRTOTAL, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRCASA, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRTOTALGERAL, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRPRESTCARTAOAREC, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRDONORXCARTAOAREC, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRTECRXCARTAOAREC, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRDONOULTRACARTAOAREC, tpDados_Moeda) & _
    " , " & Formata_Dados(strVRTOTALAREC, tpDados_Moeda) & _
    " , " & Formata_Dados(strDATAAREC, tpDados_DataHora) & _
    " , " & Formata_Dados(lngGRID, tpDados_Longo) & _
    ");"
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContaCorrente.InserirFinanceiro]", _
            Err.Description
End Sub

Public Function SelecionarUltCCDaGR(lngGRID As Long, _
                                    strStatusLancamento As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT MAX(CONTACORRENTE.PKID) AS MAXPKID " & _
    " FROM CONTACORRENTE " & _
    " WHERE CONTACORRENTE."
    
  Select Case strStatusLancamento & ""
  Case "RC"
    strSql = strSql & "GRID"
  Case Else
    strSql = strSql & "" 'Para forçar erro
  End Select
    
    
  strSql = strSql & " = " & Formata_Dados(lngGRID, tpDados_Longo)
  
  Set SelecionarUltCCDaGR = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContaCorrente.SelecionarUltCCDaGR]", _
            Err.Description
End Function

Public Function SelecionarPagamentos(strStatusLanc As String, _
                                     lngGR As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CONTACORRENTE.*, ISNULL(CARTAODEBITO.INDCONVENIO, 'N') AS INDCONVENIO, CARTAODEBITO.NOME AS NOME_CARTAO_DEBITO, " & _
    " case when STATUSCC = 'CC' then 'Cartão de Crédito' " & _
    " when STATUSCC = 'CD' then 'Cartao de Debito' " & _
    " when STATUSCC = 'ES' then 'Especie' " & _
    " when STATUSCC = 'PH' then 'Penhor' " & _
    " when STATUSCC = 'CH' then 'Cheque' " & _
    " when STATUSCC = 'FT' then 'Fatura' end AS DESC_STATUSCC " & _
    " FROM CONTACORRENTE " & _
    " LEFT JOIN CARTAODEBITO ON CARTAODEBITO.PKID = CONTACORRENTE.CARTAODEBITOID "
  Select Case strStatusLanc
  Case "RC"
    strSql = strSql & " WHERE CONTACORRENTE.GRID = "
  End Select
  strSql = strSql & Formata_Dados(lngGR, tpDados_Longo)
  
  Set SelecionarPagamentos = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContaCorrente.SelecionarPagamentos]", _
            Err.Description
End Function

Public Function SelecionarContaCorrente(lngContaCorrenteId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CONTACORRENTE.*, PRONTUARIO.NOME AS NOME_FUNC, CARTAO.NOME AS NOME_CARTAO, BANCO.NOME AS NOME_BANCO, CARTAODEBITO.NOME AS NOME_CARTAODEBITO, " & _
    " case INDDEBITOCREDITO when 'D' then 'Débito' else 'Crédito' end AS DESC_INDDEBCRED, " & _
    " case STATUSCC when 'CC' then 'Cartão de Crédito' " & _
    " when 'CD' then 'Cartão de Débito' " & _
    " when 'ES' then 'Espécie' " & _
    " when 'PH' then 'Penhor' " & _
    " when 'CH' then 'Cheque' " & _
    " when 'FT' then 'Fatura' end AS DESC_STATUSCC " & _
    " FROM CONTACORRENTE " & _
    " LEFT JOIN PRONTUARIO ON PRONTUARIO.PKID = CONTACORRENTE.FUNCIONARIOID " & _
    " LEFT JOIN CARTAO ON CARTAO.PKID = CONTACORRENTE.CARTAOID " & _
    " LEFT JOIN BANCO ON BANCO.PKID = CONTACORRENTE.BANCOID " & _
    " LEFT JOIN CARTAODEBITO ON CARTAODEBITO.PKID = CONTACORRENTE.CARTAODEBITOID " & _
    " WHERE CONTACORRENTE.PKID = " & _
    Formata_Dados(lngContaCorrenteId, tpDados_Longo, tpNulo_NaoAceita)
  
  Set SelecionarContaCorrente = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContaCorrente.SelecionarContaCorrente]", _
            Err.Description
End Function

Public Sub ExcluirContaCorrente(strStatusLanc As String, _
                                lngContaCorrenteId As Long)
  On Error GoTo trata
  '
  Dim strSql          As String
  strSql = "DELETE FROM CONTACORRENTE WHERE PKID = " & _
    Formata_Dados(lngContaCorrenteId, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContaCorrente.ExcluirContaCorrente]", _
            Err.Description
End Sub

Public Sub ExcluirContaCorrenteToda(strStatusLanc As String, _
                                    lngGRID As Long)
  On Error GoTo trata
  '
  Dim strSql          As String
  strSql = "DELETE CONTACORRENTE WHERE "
  Select Case strStatusLanc & ""
  Case "RC"
    strSql = strSql & "GRID"
  Case Else
    strSql = strSql & "" 'Para forçar erro
  End Select
    
  
  strSql = strSql & " = " & _
    Formata_Dados(lngGRID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContaCorrente.ExcluirContaCorrente]", _
            Err.Description
End Sub

Public Function InserirCC(lngGRID As Long, _
                          lngTURNOCCID As Long, _
                          strDtHora As String, _
                          strValor As String, _
                          strIndDebitoCredito As String, _
                          strStatusCC As String, _
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
  '
  Dim strSql As String
      
  strSql = "INSERT INTO CONTACORRENTE ("
  Select Case strStatusLancamento & ""
  Case "RC"
    strSql = strSql & "GRID"
  Case Else
    strSql = strSql & "" 'Para forçar erro
  End Select
    
  strSql = strSql & ", TURNOCCID, DTHORACC, VALOR, INDDEBITOCREDITO, STATUSCC, STATUSLANCAMENTO, CARTAOID, BANCOID, FUNCIONARIOID, RESPONSAVEL, AGENCIA, CONTA, NROCHEQUE, CPF, CLIENTE, DESCOBJETO, DOCUMENTOPENHOR, LOTE, CARTAODEBITOID, VRGORJETA, VRTROCO, NROPARCELAS, DTPRIMEIRAPARCELA) VALUES (" & _
    Formata_Dados(lngGRID, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(lngTURNOCCID, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strDtHora, tpDados_DataHora, tpNulo_Aceita) & _
    ", " & Formata_Dados(strValor, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(strIndDebitoCredito, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strStatusCC, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strStatusLancamento, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strCartaoId, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strBancoId, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strFuncionarioId, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strResponsavel, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strAgencia, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strConta, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strNroCheque, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strCPF, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strCliente, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strObjeto, tpDados_Texto, tpNulo_Aceita) & ", " & Formata_Dados(strDocumentoPenhor, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strLote, tpDados_Texto, tpNulo_Aceita) & ", " & Formata_Dados(strCartaoDebId, tpDados_Longo, tpNulo_Aceita) & ", " & Formata_Dados(strVrGorjeta, tpDados_Moeda, tpNulo_Aceita) & ", " & Formata_Dados(strVrCalcTroco, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(strNroParcelas, tpDados_Longo) & _
    ", " & Formata_Dados(strDtPrimeiraParc, tpDados_DataHora) & _
    ");"
  RunSP strSql
  InserirCC = 0
  'InserirCC = ExecSPReturnInteger("sp_conta_corrente_inserir", "")
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContaCorrente.InserirCC]", _
            Err.Description
End Function

Public Sub AlterarCCValor(lngCCId As Long, _
                          strValor As String)
  On Error GoTo trata
  '
  Dim strSql As String
      
  strSql = "UPDATE CONTACORRENTE SET " & _
    "VALOR = " & Formata_Dados(strValor, tpDados_Moeda, tpNulo_Aceita) & _
    " WHERE PKID = " & Formata_Dados(lngCCId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContaCorrente.AlterarCCValor]", _
            Err.Description
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
  '
  Dim strSql As String
      
  strSql = "UPDATE CONTACORRENTE SET " & _
    "VALOR = " & Formata_Dados(strValor, tpDados_Moeda, tpNulo_Aceita) & _
    ", STATUSCC = " & Formata_Dados(strStatusCC, tpDados_Texto) & _
    ", CARTAOID = " & Formata_Dados(strCartaoId, tpDados_Longo, tpNulo_Aceita) & _
    ", BANCOID = " & Formata_Dados(strBancoId, tpDados_Longo, tpNulo_Aceita) & _
    ", FUNCIONARIOID = " & Formata_Dados(strFuncionarioId, tpDados_Longo, tpNulo_Aceita) & _
    ", AGENCIA = " & Formata_Dados(strAgencia, tpDados_Texto, tpNulo_Aceita) & _
    ", CONTA = " & Formata_Dados(strConta, tpDados_Texto, tpNulo_Aceita) & _
    ", NROCHEQUE = " & Formata_Dados(strNroCheque, tpDados_Texto, tpNulo_Aceita) & _
    ", CPF = " & Formata_Dados(strCPF, tpDados_Texto, tpNulo_Aceita) & _
    ", CLIENTE = " & Formata_Dados(strCliente, tpDados_Texto, tpNulo_Aceita) & _
    ", DESCOBJETO = " & Formata_Dados(strObjeto, tpDados_Texto, tpNulo_Aceita) & _
    ", DOCUMENTOPENHOR = " & Formata_Dados(strDocumentoPenhor, tpDados_Texto, tpNulo_Aceita) & _
    ", LOTE = " & Formata_Dados(strLote, tpDados_Texto, tpNulo_Aceita) & _
    ", CARTAODEBITOID = " & Formata_Dados(strCartaoDebId, tpDados_Longo, tpNulo_Aceita) & _
    ", VRGORJETA = " & Formata_Dados(strVrGorjeta, tpDados_Moeda, tpNulo_Aceita) & _
    ", VRTROCO = " & Formata_Dados(strVrCalcTroco, tpDados_Moeda, tpNulo_Aceita) & _
    ", NROPARCELAS = " & Formata_Dados(strNroParcelas, tpDados_Longo) & _
    ", DTPRIMEIRAPARCELA = " & Formata_Dados(strDtPrimeiraParc, tpDados_DataHora) & _
    " WHERE PKID = " & Formata_Dados(lngCCId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsContaCorrente.AlterarCC]", _
            Err.Description
End Sub


