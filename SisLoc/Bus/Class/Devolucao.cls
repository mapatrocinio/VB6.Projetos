VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDevolucao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirItemDeVolucao(lngITEMDEVOLUCAOID As Long)

  On Error GoTo trata
  Dim objDevolucao As datSisLoc.clsDevolucao
  '
  Set objDevolucao = New datSisLoc.clsDevolucao
  'Alterar Estoque Pela Devolucao
  objDevolucao.ExlcuirItemDevolucao lngITEMDEVOLUCAOID
  '
  Set objDevolucao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirItemDeVolucao(lngITEMNFID As Long, _
                                lngDEVOLUCAOID As Long, _
                                strQuantidade As String, _
                                strQuantidadeNF As String, _
                                strQuantidadeAvaria As String, _
                                strValorIndenizacaoest As String)

  On Error GoTo trata
  Dim objDevolucao As datSisLoc.clsDevolucao
  '
  Set objDevolucao = New datSisLoc.clsDevolucao
  'Alterar Estoque Pela Devolucao
  objDevolucao.InserirItemDeVolucao lngITEMNFID, _
                                    lngDEVOLUCAOID, _
                                    strQuantidade, _
                                    strQuantidadeNF, _
                                    strQuantidadeAvaria, _
                                    strValorIndenizacaoest
  '
  Set objDevolucao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarEstoquePelaDevolucao(ByVal lngESTOQUEID As Long, _
                                       ByVal strQuantidade As String, _
                                       ByVal strStatus As String)

  On Error GoTo trata
  Dim objDevolucao As datSisLoc.clsDevolucao
  '
  Set objDevolucao = New datSisLoc.clsDevolucao
  'Alterar Estoque Pela Devolucao
  objDevolucao.AlterarEstoquePelaDevolucao lngESTOQUEID, _
                                           strQuantidade, _
                                           strStatus
  '
  Set objDevolucao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirDevolucao(ByRef lngDEVOLUCAOID As Long, _
                            ByRef lngContratoId As Long, _
                            ByRef lngOBRAID As Long, _
                            ByRef strSequencial As String, _
                            strData As String, _
                            strNFCliente As String, _
                            strDataEmissao As String, _
                            strDataDevolucao As String)

  On Error GoTo trata
  Dim objDevolucao As datSisLoc.clsDevolucao
  '
  Set objDevolucao = New datSisLoc.clsDevolucao
  'Inserir Devolucao
  strSequencial = objDevolucao.CapturaSequencialDevolucao(lngContratoId)
  '
  objDevolucao.InserirDevolucao lngDEVOLUCAOID, _
                                lngContratoId, _
                                lngOBRAID, _
                                strSequencial, _
                                strData, _
                                strNFCliente, _
                                strDataEmissao, _
                                strDataDevolucao
  '
  Set objDevolucao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarDevolucao(lngDEVOLUCAOID As Long, _
                            strData As String, _
                            strNFCliente As String, _
                            strDataEmissao As String, _
                            strDataDevolucao As String)
                         
  On Error GoTo trata
  Dim objDevolucao As datSisLoc.clsDevolucao
  Set objDevolucao = New datSisLoc.clsDevolucao
  
  'Inserir Devolucao
  objDevolucao.AlterarDevolucao lngDEVOLUCAOID, _
                                strData, _
                                strNFCliente, _
                                strDataEmissao, _
                                strDataDevolucao

  '
  Set objDevolucao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarDevolucaoPeloSeq(lngContratoId As Long, _
                                       lngOBRAID As Long, _
                                       strSequencial As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objDevolucao As datSisLoc.clsDevolucao
  Set objDevolucao = New datSisLoc.clsDevolucao
  
  'Listar Devolucao
  Set ListarDevolucaoPeloSeq = objDevolucao.ListarDevolucaoPeloSeq(lngContratoId, _
                                                                   lngOBRAID, _
                                                                   strSequencial)
  '
  Set objDevolucao = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ListarDevolucao(lngDEVOLUCAOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objDevolucao As datSisLoc.clsDevolucao
  Set objDevolucao = New datSisLoc.clsDevolucao
  
  'Listar Devolucao
  Set ListarDevolucao = objDevolucao.ListarDevolucao(lngDEVOLUCAOID)
  '
  Set objDevolucao = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirDevolucao(lngDEVOLUCAOID As Long)

  On Error GoTo trata
  Dim objDevolucao As datSisLoc.clsDevolucao
  Set objDevolucao = New datSisLoc.clsDevolucao
  
  'Excluir Devolucao
  objDevolucao.ExlcuirDevolucao lngDEVOLUCAOID
  '
  Set objDevolucao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

