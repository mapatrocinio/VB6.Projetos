VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFaixaSubProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarFaixaSubProdutoPeloPkid(lngFaixaSubProdutoId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objFaixaSubProduto As datApler.clsFaixaSubProduto
  Set objFaixaSubProduto = New datApler.clsFaixaSubProduto
  
  'Listar FaixaSubProduto
  Set SelecionarFaixaSubProdutoPeloPkid = objFaixaSubProduto.SelecionarFaixaSubProdutoPeloPkid(lngFaixaSubProdutoId)
  '
  Set objFaixaSubProduto = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirFaixaSubProduto(lngFaixaSubProdutoId As Long)
  On Error GoTo trata
  Dim objFaixaSubProduto As datApler.clsFaixaSubProduto
  Set objFaixaSubProduto = New datApler.clsFaixaSubProduto
  
  'Excluir FaixaSubProduto
  objFaixaSubProduto.ExcluirFaixaSubProduto lngFaixaSubProdutoId
  '
  Set objFaixaSubProduto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirFaixaSubProduto(lngSUBPRODUTOID As Long, _
                                  strDescricao As String, _
                                  strDatInicial As String, _
                                  strDataFinal As String, _
                                  strFxInicial As String, _
                                  strFxFinal As String, _
                                  strValor As String, _
                                  strPrCusto As String)
  On Error GoTo trata
  Dim objFaixaSubProduto     As datApler.clsFaixaSubProduto
  '
  Set objFaixaSubProduto = New datApler.clsFaixaSubProduto
  'Inserir FaixaSubProduto
  objFaixaSubProduto.InserirFaixaSubProduto lngSUBPRODUTOID, _
                                            strDescricao, _
                                            strDatInicial, _
                                            strDataFinal, _
                                            strFxInicial, _
                                            strFxFinal, _
                                            strValor, _
                                            strPrCusto
  '
  Set objFaixaSubProduto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarFaixaSubProduto(lngPKID As Long, _
                                  strDescricao As String, _
                                  strDatInicial As String, _
                                  strDataFinal As String, _
                                  strDataDesativacao As String, _
                                  strFxInicial As String, _
                                  strFxFinal As String, _
                                  strValor As String, _
                                  strPrCusto As String, _
                                  strStatus As String)

  On Error GoTo trata
  Dim objFaixaSubProduto     As datApler.clsFaixaSubProduto
  '
  Set objFaixaSubProduto = New datApler.clsFaixaSubProduto
  'Atualizar FaixaSubProduto
  objFaixaSubProduto.AlterarFaixaSubProduto lngPKID, _
                                            strDescricao, _
                                            strDatInicial, _
                                            strDataFinal, _
                                            strDataDesativacao, _
                                            strFxInicial, _
                                            strFxFinal, _
                                            strValor, _
                                            strPrCusto, _
                                            strStatus
  '
  Set objFaixaSubProduto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

