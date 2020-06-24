VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItemPedidoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirItemPedidoVenda(lngPEDIDOVENDAID As Long, _
                                 lngPRODUTOID As Long, _
                                 strQuantidade As String, _
                                 lngFUNCIONARIOID As Long)
  On Error GoTo trata
  Dim objItemPedidoVenda    As datSisMetal.clsItemPedidoVenda
  Dim objInsumo             As datSisMetal.clsInsumo
  Dim objFuncionario        As datSisMetal.clsFuncionario
  Dim objRs                 As ADODB.Recordset
'''  Dim curComprimento  As Currency
  Dim strValorInstalacao      As String
  Dim strValorFrete           As String
  Dim curValor                As Currency
  Dim curValorComissVend      As Currency
  Dim curPercComiss           As Currency
  Dim blnComissaoVendedor     As Boolean
  '
  curValor = 0
  curValorComissVend = 0
  Set objInsumo = New datSisMetal.clsInsumo
  Set objRs = objInsumo.SelecionarProdutoPeloPkid(lngPRODUTOID)
  If Not objRs.EOF Then
    curValor = IIf(IsNull(objRs.Fields("PRECO").Value), 0, objRs.Fields("PRECO").Value) * CLng(strQuantidade)
    blnComissaoVendedor = IIf(IsNull(objRs.Fields("COMISSAO_VENDEDOR").Value), False, objRs.Fields("COMISSAO_VENDEDOR").Value)
  End If
  objRs.Close
  Set objRs = Nothing
  Set objInsumo = Nothing
  '
  If blnComissaoVendedor = True Then
    'Calcula valor da comissão para o vendedor
    curPercComiss = 0
    Set objFuncionario = New datSisMetal.clsFuncionario
    Set objRs = objFuncionario.SelecionarFuncionarioPeloPkid(lngFUNCIONARIOID)
    If Not objRs.EOF Then
      curPercComiss = IIf(IsNull(objRs.Fields("PERC_VENDA").Value), 0, objRs.Fields("PERC_VENDA").Value)
    End If
    objRs.Close
    Set objRs = Nothing
    Set objFuncionario = Nothing
    If curPercComiss <> 0 Then
      'Existe comissão cadastrada para vendedor, realizar cálculo
      curValorComissVend = curValor * curPercComiss / 100
    End If
  End If
  Set objItemPedidoVenda = New datSisMetal.clsItemPedidoVenda
  'Inserir ItemPedidoVenda
  'Selecionar PESO_VARA da linha
  'Format(curComprimento, "###,##0.000")
  objItemPedidoVenda.InserirItemPedidoVenda lngPEDIDOVENDAID, _
                                            lngPRODUTOID, _
                                            strQuantidade, _
                                            Format(curValor, "###,##0.00"), _
                                            strValorInstalacao, _
                                            strValorFrete, _
                                            Format(curValorComissVend, "###,##0.00")

  '
  Set objItemPedidoVenda = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarItemPedidoVenda(lngPEDIDOVENDAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objItemPedidoVenda As datSisMetal.clsItemPedidoVenda
  Set objItemPedidoVenda = New datSisMetal.clsItemPedidoVenda
  
  'Listar ItemPedidoVenda
  Set ListarItemPedidoVenda = objItemPedidoVenda.ListarItemPedidoVenda(lngPEDIDOVENDAID)
  '
  Set objItemPedidoVenda = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarItemPedidoVendaPeloPkid(lngItemPedidoVendaId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objItemPedidoVenda As datSisMetal.clsItemPedidoVenda
  Set objItemPedidoVenda = New datSisMetal.clsItemPedidoVenda
  
  'Listar ItemPedidoVenda
  Set SelecionarItemPedidoVendaPeloPkid = objItemPedidoVenda.SelecionarItemPedidoVendaPeloPkid(lngItemPedidoVendaId)
  '
  Set objItemPedidoVenda = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirItemPedidoVenda(lngItemPedidoVendaId As Long)
  On Error GoTo trata
  Dim objItemPedidoVenda As datSisMetal.clsItemPedidoVenda
  Set objItemPedidoVenda = New datSisMetal.clsItemPedidoVenda
  
  'Excluir ItemPedidoVenda
  objItemPedidoVenda.ExcluirItemPedidoVenda lngItemPedidoVendaId
  '
  Set objItemPedidoVenda = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarItemPedidoVenda(lngPKID As Long, _
                                  lngPRODUTOID As Long, _
                                  strQuantidade As String, _
                                  lngFUNCIONARIOID As Long)

  On Error GoTo trata
  Dim objItemPedidoVenda        As datSisMetal.clsItemPedidoVenda
  Dim strValorInstalacao        As String
  Dim strValorFrete             As String
  '
  Dim objInsumo                 As datSisMetal.clsInsumo
  Dim objFuncionario        As datSisMetal.clsFuncionario
  Dim objRs                     As ADODB.Recordset
  Dim curValor                  As Currency
  Dim curValorComissVend      As Currency
  Dim curPercComiss           As Currency
  Dim blnComissaoVendedor     As Boolean
  '
  curValor = 0
  curValorComissVend = 0
  Set objInsumo = New datSisMetal.clsInsumo
  Set objRs = objInsumo.SelecionarProdutoPeloPkid(lngPRODUTOID)
  If Not objRs.EOF Then
    curValor = IIf(IsNull(objRs.Fields("PRECO").Value), 0, objRs.Fields("PRECO").Value) * CLng(strQuantidade)
    blnComissaoVendedor = IIf(IsNull(objRs.Fields("COMISSAO_VENDEDOR").Value), False, objRs.Fields("COMISSAO_VENDEDOR").Value)
  End If
  objRs.Close
  Set objRs = Nothing
  Set objInsumo = Nothing
  '
  If blnComissaoVendedor = True Then
    'Calcula valor da comissão para o vendedor
    curPercComiss = 0
    Set objFuncionario = New datSisMetal.clsFuncionario
    Set objRs = objFuncionario.SelecionarFuncionarioPeloPkid(lngFUNCIONARIOID)
    If Not objRs.EOF Then
      curPercComiss = IIf(IsNull(objRs.Fields("PERC_VENDA").Value), 0, objRs.Fields("PERC_VENDA").Value)
    End If
    objRs.Close
    Set objRs = Nothing
    Set objFuncionario = Nothing
    If curPercComiss <> 0 Then
      'Existe comissão cadastrada para vendedor, realizar cálculo
      curValorComissVend = curValor * curPercComiss / 100
    End If
  End If
  Set objItemPedidoVenda = New datSisMetal.clsItemPedidoVenda
  'Atualizar ItemPedidoVenda
  objItemPedidoVenda.AlterarItemPedidoVenda lngPKID, _
                                  strQuantidade, _
                                  Format(curValor, "###,##0.00"), _
                                  strValorInstalacao, _
                                  strValorFrete, _
                                  Format(curValorComissVend, "###,##0.00")
  '
  Set objItemPedidoVenda = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub



