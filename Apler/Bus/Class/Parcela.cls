VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParcela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarParcela(lngParcelaId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objParcela As datApler.clsParcela
  Set objParcela = New datApler.clsParcela
  
  'Listar Parcela
  Set SelecionarParcela = objParcela.SelecionarParcela(lngParcelaId)
  '
  Set objParcela = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirParcela(lngParcelaId As Long)

  On Error GoTo trata
  Dim objParcela As datApler.clsParcela
  Set objParcela = New datApler.clsParcela
  
  'Excluir Parcela
  objParcela.ExcluirParcela lngParcelaId
  '
  Set objParcela = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirParcela(lngContaCorrenteId As Long, _
                          strParcela As Long, _
                          strDtVencimento As String, _
                          strDtPagamento As String, _
                          strVrParcela As String, _
                          strVrPago As String)
  On Error GoTo trata
  Dim clsParcela As datApler.clsParcela
  '
  Set clsParcela = New datApler.clsParcela
  
  'Inserir Parcela
  
  clsParcela.InserirParcela lngContaCorrenteId, _
                            strParcela, _
                            strDtVencimento, _
                            strDtPagamento, _
                            strVrParcela, _
                            strVrPago
  '
  Set clsParcela = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarParcela(lngParcelaId As Long, _
                          strDtVencimento As String, _
                          strDtPagamento As String, _
                          strVrParcela As String, _
                          strVrPago As String)
  On Error GoTo trata
  Dim clsParcela As datApler.clsParcela
  '
  Set clsParcela = New datApler.clsParcela
  
  'Alterar Parcela
  
  clsParcela.AlterarParcela lngParcelaId, _
                            strDtVencimento, _
                            strDtPagamento, _
                            strVrParcela, _
                            strVrPago
  '
  Set clsParcela = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ExcluirParcelasDaCC(lngCCId As Long)

  On Error GoTo trata
  Dim objParcela As datApler.clsParcela
  Set objParcela = New datApler.clsParcela
  
  'Excluir Parcela
  objParcela.ExcluirParcelasDaCC lngCCId
  '
  Set objParcela = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub CadastrarParcelas(lngContaCorrenteId As Long, _
                             strParcelas As Long, _
                             strDtPrimeiroVencimento As String, _
                             strVrTotalParcelas As String)
  On Error GoTo trata
  Dim objParcela As datApler.clsParcela
  Dim intI As Integer
  Dim curVrParcela As Currency
  Dim dtaDtVencimento As Date
  '
  Set objParcela = New datApler.clsParcela
  'Obter varlo de cada parcela
  curVrParcela = CCur(strVrTotalParcelas) / CLng(strParcelas)
  For intI = 1 To CInt(strParcelas)
    'Para cada parcela
    'Obter Data do Vencimento
    dtaDtVencimento = DateAdd("M", intI - 1, CDate(Right(strDtPrimeiroVencimento, 4) & "/" & Mid(strDtPrimeiroVencimento, 4, 2) & "/" & Left(strDtPrimeiroVencimento, 2)))
    'Inserir Parcela
    objParcela.InserirParcela lngContaCorrenteId, _
                              CLng(intI), _
                              Format(dtaDtVencimento, "DD/MM/YYYY"), _
                              "", _
                              Format(curVrParcela, "###,##0.00"), _
                              ""
  Next
  '
  Set objParcela = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


