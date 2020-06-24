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

Public Function SelecionarPagamentos(strStatusLanc As String, _
                                     lngLOCDESPVDAEXT As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objCC As datApler.clsContaCorrente
  Set objCC = New datApler.clsContaCorrente
  
  'Selecionar pagamentos da Conta Corrente
  Set SelecionarPagamentos = objCC.SelecionarPagamentos(strStatusLanc, _
                                                        lngLOCDESPVDAEXT)
  '
  Set objCC = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarContaCorrente(lngContaCorrenteId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objCC As datApler.clsContaCorrente
  Set objCC = New datApler.clsContaCorrente
  
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
  Dim objContaCorrente  As datApler.clsContaCorrente
  Dim blnChequeTerc     As Boolean
  Dim objRs             As ADODB.Recordset
  Set objContaCorrente = New datApler.clsContaCorrente
  
  blnChequeTerc = False
  'Verificar se strStatusLanc = DESPESA
  If strStatusLanc = "DE" Then
    Set objRs = objContaCorrente.SelecionarContaCorrente(lngContaCorrenteId)
    If Not objRs.EOF Then
      If objRs.Fields("LOCACAOID").Value & "" <> "" Or _
         objRs.Fields("VENDAID").Value & "" <> "" Or _
         objRs.Fields("EXTRAID").Value & "" <> "" Then
         'É um cheque de terceiros
         blnChequeTerc = True
      End If
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  
  'Excluir ContaCorrente
  objContaCorrente.ExcluirContaCorrente strStatusLanc, _
                                        lngContaCorrenteId, _
                                        blnChequeTerc
  '
  Set objContaCorrente = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ExcluirContaCorrenteToda(strStatusLanc As String, _
                                    lngLocDespVdaExtId As Long)

  On Error GoTo trata
  Dim objContaCorrente  As datApler.clsContaCorrente
  Dim objRs             As ADODB.Recordset
  Set objContaCorrente = New datApler.clsContaCorrente
  
  'Excluir Toda ContaCorrente
  objContaCorrente.ExcluirContaCorrenteToda strStatusLanc, _
                                            lngLocDespVdaExtId
  '
  Set objContaCorrente = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function InserirCC(lngLocDespVdaExtId As Long, _
                          lngTURNOCCID As Long, strDtHora As String, strValor As String, strIndDebitoCredito As String, strStatusCC As String, _
                          strStatusLancamento As String, _
                          strCartaoId As String, _
                          strBancoId As String, _
                          strGarcomId As String, _
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
  Dim clsCC   As datApler.clsContaCorrente
  Dim objRs   As ADODB.Recordset
  '
  Set clsCC = New datApler.clsContaCorrente
  InserirCC = 0
  'Inserir Conta Corrente
  
  InserirCC = clsCC.InserirCC(lngLocDespVdaExtId, _
                              lngTURNOCCID, _
                              strDtHora, _
                              strValor, _
                              strIndDebitoCredito, _
                              strStatusCC, _
                              strStatusLancamento, _
                              strCartaoId, _
                              strBancoId, _
                              strGarcomId, _
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
  
  Set objRs = clsCC.SelecionarUltCCDaLoc(lngLocDespVdaExtId, _
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



Public Sub AssociaCCDespesa(strDespesaId As String, _
                            strCCId As String)
  On Error GoTo trata
  Dim clsCC As datApler.clsContaCorrente
  '
  Set clsCC = New datApler.clsContaCorrente
  
  'Associar Despesa a Conta Corrente
  clsCC.AssociaCCDespesa strDespesaId, _
                         strCCId
  '
  Set clsCC = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarCCValor(lngCCId As Long, _
                          strValor As String)
  On Error GoTo trata
  Dim clsCC As datApler.clsContaCorrente
  '
  Set clsCC = New datApler.clsContaCorrente
  
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
                     strGarcomId As String, _
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
  Dim clsCC As datApler.clsContaCorrente
  '
  Set clsCC = New datApler.clsContaCorrente
  
  'Alterar Conta Corrente
  
  clsCC.AlterarCC strStatusCC, _
                  lngCCId, _
                  strValor, _
                  strCartaoId, _
                  strBancoId, _
                  strGarcomId, _
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


