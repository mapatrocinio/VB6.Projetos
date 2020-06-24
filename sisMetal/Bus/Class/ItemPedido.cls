VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItemPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirItemPedidoItem(lngITEM_PEDIDOID As Long, _
                                 lngPEDIDOID As Long, _
                                 lngLINHAID As Long, _
                                 strPeso As String, _
                                 strPesoAnod As String, _
                                 strPesoFab As String)
  On Error GoTo trata
  Dim objItemPedido   As datSisMetal.clsItemPedido
  Dim objLinhaPerfil  As datSisMetal.clsLinhaPerfil
  Dim objRs           As ADODB.Recordset
  Dim curComprimento  As Currency
  '
  curComprimento = 0
  Set objLinhaPerfil = New datSisMetal.clsLinhaPerfil
  Set objRs = objLinhaPerfil.SelecionarLinha(lngLINHAID)
  If Not objRs.EOF Then
    curComprimento = objRs.Fields("VALOR_VARA").Value
  End If
  objRs.Close
  Set objRs = Nothing
  Set objLinhaPerfil = Nothing
  '
  Set objItemPedido = New datSisMetal.clsItemPedido
  'Inserir ItemPedido
  'Selecionar PESO_VARA da linha
  
  objItemPedido.InserirItemPedidoItem lngITEM_PEDIDOID, _
                                      lngPEDIDOID, _
                                      lngLINHAID, _
                                      strPeso, _
                                      Format(curComprimento, "###,##0.000"), _
                                      strPesoAnod, _
                                      strPesoFab
  '
  Set objItemPedido = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function SelecionarItemPedidoPeloPkid(lngItemPedidoId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objItemPedido As datSisMetal.clsItemPedido
  Set objItemPedido = New datSisMetal.clsItemPedido
  
  'Listar ItemPedido
  Set SelecionarItemPedidoPeloPkid = objItemPedido.SelecionarItemPedidoPeloPkid(lngItemPedidoId)
  '
  Set objItemPedido = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirItemPedido(lngItemPedidoId As Long)
  On Error GoTo trata
  Dim objItemPedido As datSisMetal.clsItemPedido
  Set objItemPedido = New datSisMetal.clsItemPedido
  
  'Excluir ItemPedido
  objItemPedido.ExcluirItemPedido lngItemPedidoId
  '
  Set objItemPedido = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirItemPedido(lngPEDIDOID As Long, _
                             lngLINHAID As Long, _
                             strPeso As String, _
                             strComprimentoVara As String)
  On Error GoTo trata
  Dim objItemPedido       As datSisMetal.clsItemPedido
  Dim curQuantidade       As Currency
  '
  curQuantidade = CalculoQuantidadePedido(lngLINHAID, _
                                          strPeso)
  '
  Set objItemPedido = New datSisMetal.clsItemPedido
  'Inserir ItemPedido
  objItemPedido.InserirItemPedido lngPEDIDOID, _
                                  lngLINHAID, _
                                  Format(IIf(curQuantidade = 0, "", curQuantidade), "###,###") & "", _
                                  strPeso, _
                                  strComprimentoVara
  '
  Set objItemPedido = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function CalculoQuantidadePedido(lngLINHAID As Long, _
                                        strPeso As String) As Currency

  On Error GoTo trata
  Dim objItemPedido       As datSisMetal.clsItemPedido
  'Obter peso da linha
  Set objItemPedido = New datSisMetal.clsItemPedido
  CalculoQuantidadePedido = objItemPedido.CalculoQuantidadePedido(lngLINHAID, _
                                                                  strPeso)
  Set objItemPedido = Nothing
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CalculoPesoPedido(lngLINHAID As Long, _
                                  strQuantidade As String) As Currency

  On Error GoTo trata
  Dim objItemPedido       As datSisMetal.clsItemPedido
  'Obter peso da linha
  Set objItemPedido = New datSisMetal.clsItemPedido
  CalculoPesoPedido = objItemPedido.CalculoPesoPedido(lngLINHAID, _
                                                      strQuantidade)
  Set objItemPedido = Nothing
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Sub AlterarItemPedido(lngPKID As Long, _
                             lngLINHAID As Long, _
                             strPeso As String, _
                             strComprimentoVara As String)

  On Error GoTo trata
  Dim objItemPedido       As datSisMetal.clsItemPedido
  Dim curQuantidade       As Currency
  '
  curQuantidade = CalculoQuantidadePedido(lngLINHAID, _
                                          strPeso)
  Set objItemPedido = New datSisMetal.clsItemPedido
  'Atualizar ItemPedido
  objItemPedido.AlterarItemPedido lngPKID, _
                                  lngLINHAID, _
                                  Format(IIf(curQuantidade = 0, "", curQuantidade), "###,###") & "", _
                                  strPeso, _
                                  strComprimentoVara
  '
  Set objItemPedido = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

