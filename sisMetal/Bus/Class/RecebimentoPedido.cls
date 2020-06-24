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

'''Public Function SelecionarPagamentos(strStatusLanc As String, _
'''                                     lngLOCDESPVDAEXT As Long) As ADODB.Recordset
'''  On Error GoTo trata
'''  Dim objCC As datSisMetal.clsRecebimentoPedido
'''  Set objCC = New datSisMetal.clsRecebimentoPedido
'''
'''  'Selecionar pagamentos da Conta Corrente
'''  Set SelecionarPagamentos = objCC.SelecionarPagamentos(strStatusLanc, _
'''                                                        lngLOCDESPVDAEXT)
'''  '
'''  Set objCC = Nothing
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Function
'''
Public Function SelecionarContaCorrente(lngContaCorrentePedId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objCC As datSisMetal.clsRecebimentoPedido
  Set objCC = New datSisMetal.clsRecebimentoPedido

  'Listar Conta Corrente
  Set SelecionarContaCorrente = objCC.SelecionarContaCorrente(lngContaCorrentePedId)
  '
  Set objCC = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirContaCorrente(lngContaCorrentePedId As Long)

  On Error GoTo trata
  Dim objContaCorrente  As datSisMetal.clsRecebimentoPedido
  Set objContaCorrente = New datSisMetal.clsRecebimentoPedido
  'Excluir ContaCorrente
  objContaCorrente.ExcluirContaCorrente lngContaCorrentePedId
  '
  Set objContaCorrente = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'''Public Sub ExcluirContaCorrenteToda(strStatusLanc As String, _
'''                                    lngLocDespVdaExtId As Long)
'''
'''  On Error GoTo trata
'''  Dim objContaCorrente  As datSisMetal.clsRecebimentoPedido
'''  Dim objRs             As ADODB.Recordset
'''  Set objContaCorrente = New datSisMetal.clsRecebimentoPedido
'''
'''  'Excluir Toda ContaCorrente
'''  objContaCorrente.ExcluirContaCorrenteToda strStatusLanc, _
'''                                            lngLocDespVdaExtId
'''  '
'''  Set objContaCorrente = Nothing
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Sub
'''
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
  Dim objRecebimentoPedido   As datSisMetal.clsRecebimentoPedido
'''  Dim objRs   As ADODB.Recordset
  '
  Set objRecebimentoPedido = New datSisMetal.clsRecebimentoPedido
  InserirCC = 0
  'Inserir Conta Corrente

  InserirCC = objRecebimentoPedido.InserirCC(lngPEDIDOVENDAID, _
                                              strDtHora, _
                                              strValor, _
                                              strIndDebitoCredito, _
                                              strStatusCC, _
                                              strCartaoId, _
                                              strBancoId, _
                                              strResponsavel, _
                                              strAgencia, _
                                              strConta, _
                                              strNroCheque, _
                                              strCPF, _
                                              strCliente, _
                                              strLote, _
                                              strCartaoDebId, _
                                              strVrCalcTroco, _
                                              strNroParcelas, _
                                              strDtPrimeiraParc)

'''  Set objRs = objRecebimentoPedido.SelecionarUltCCDaLoc(lngLocDespVdaExtId, _
'''                                         strStatusLancamento)
'''  If Not objRs.EOF Then
'''    If IsNumeric(objRs.Fields("MAXPKID").Value) Then
'''      InserirCC = objRs.Fields("MAXPKID").Value
'''    End If
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
  '
  Set objRecebimentoPedido = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function



'''Public Sub AssociaCCDespesa(strDespesaId As String, _
'''                            strCCId As String)
'''  On Error GoTo trata
'''  Dim objRecebimentoPedido As datSisMetal.clsRecebimentoPedido
'''  '
'''  Set objRecebimentoPedido = New datSisMetal.clsRecebimentoPedido
'''
'''  'Associar Despesa a Conta Corrente
'''  objRecebimentoPedido.AssociaCCDespesa strDespesaId, _
'''                         strCCId
'''  '
'''  Set objRecebimentoPedido = Nothing
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Sub
'''
'''Public Sub AlterarCCValor(lngCCId As Long, _
'''                          strValor As String)
'''  On Error GoTo trata
'''  Dim objRecebimentoPedido As datSisMetal.clsRecebimentoPedido
'''  '
'''  Set objRecebimentoPedido = New datSisMetal.clsRecebimentoPedido
'''
'''  'Alterar Conta Corrente
'''
'''  objRecebimentoPedido.AlterarCCValor lngCCId, _
'''                       strValor
'''  '
'''  Set objRecebimentoPedido = Nothing
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
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
  Dim objRecebimentoPedido As datSisMetal.clsRecebimentoPedido
  '
  Set objRecebimentoPedido = New datSisMetal.clsRecebimentoPedido

  'Alterar Conta Corrente

  objRecebimentoPedido.AlterarCC strStatusCC, _
                  lngCCId, _
                  strValor, _
                  strCartaoId, _
                  strBancoId, _
                  strAgencia, _
                  strConta, _
                  strNroCheque, _
                  strCPF, _
                  strCliente, _
                  strLote, _
                  strCartaoDebId, _
                  strVrCalcTroco, _
                  strNroParcelas, _
                  strDtPrimeiraParc

  '
  Set objRecebimentoPedido = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


