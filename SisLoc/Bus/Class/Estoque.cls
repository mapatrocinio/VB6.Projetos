VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ListarEstoquePeloCodigo(strCodigo As String) As ADODB.Recordset
  On Error GoTo trata
  Dim clsEst As datSisLoc.clsEstoque
  Set clsEst = New datSisLoc.clsEstoque
  
  'Listar Estoque pelo código
  Set ListarEstoquePeloCodigo = clsEst.ListarEstoquePeloCodigo(strCodigo)
  '
  Set clsEst = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CapturaEstoquePeloCodigo(strCodigo As String) As ADODB.Recordset
  On Error GoTo trata
  Dim clsEst As datSisLoc.clsEstoque
  Set clsEst = New datSisLoc.clsEstoque
  
  'Listar Estoque pelo código
  Set CapturaEstoquePeloCodigo = clsEst.CapturaEstoquePeloCodigo(strCodigo)
  '
  Set clsEst = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub InserirEstoque(lngESTOQUEID As Long, _
                          strCodigo As String, _
                          strDescricao As String, _
                          strUnidadeId As String, _
                          strQuantidade As String, _
                          strValor As String, _
                          strValorIndenizacao As String, _
                          strPeso As String, _
                          strAltura As String, _
                          strLargura As String, _
                          strQuantidadeMin As String)
                         
  On Error GoTo trata
  Dim clsEst      As datSisLoc.clsEstoque
  '
  Set clsEst = New datSisLoc.clsEstoque
  'Inserir Estoque
  clsEst.InserirEstoque lngESTOQUEID, _
                        strCodigo, _
                        strDescricao, _
                        strUnidadeId, _
                        strQuantidade, _
                        strValor, _
                        strValorIndenizacao, _
                        strPeso, _
                        strAltura, _
                        strLargura, _
                        strQuantidadeMin
  '
  Set clsEst = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarQuantidadeEstoque(lngESTOQUEID As Long, _
                                    lngQtdRealEstoque As Long)
                         
  On Error GoTo trata
  Dim clsEst        As datSisLoc.clsEstoque
  '
  Set clsEst = New datSisLoc.clsEstoque
  
  'Alterar Estoque e Tab_GrupoEstInter
  clsEst.AlterarQuantidadeEstoque lngESTOQUEID, _
                                  lngQtdRealEstoque

  '
  Set clsEst = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarEstoque(lngESTOQUEID As Long, _
                          strCodigo As String, _
                          strDescricao As String, _
                          strUnidadeId As String, _
                          strQuantidade As String, _
                          strValor As String, _
                          strValorIndenizacao As String, _
                          strPeso As String, _
                          strAltura As String, _
                          strLargura As String, _
                          strQuantidadeMin As String)
                         
  On Error GoTo trata
  Dim clsEst        As datSisLoc.clsEstoque
  '
  Set clsEst = New datSisLoc.clsEstoque
  
  'Alterar Estoque
  clsEst.AlterarEstoque lngESTOQUEID, _
                        strCodigo, _
                        strDescricao, _
                        strUnidadeId, _
                        strQuantidade, _
                        strValor, _
                        strValorIndenizacao, _
                        strPeso, _
                        strAltura, _
                        strLargura, _
                        strQuantidadeMin

  '
  Set clsEst = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarEstoque(lngESTOQUEID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsEst As datSisLoc.clsEstoque
  Set clsEst = New datSisLoc.clsEstoque
  
  'Listar Estoque
  Set ListarEstoque = clsEst.ListarEstoque(lngESTOQUEID)
  '
  Set clsEst = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirESTOQUE(ByVal lngESTOQUEID As Long)
  '
  Dim clsEst    As datSisLoc.clsEstoque
  '
  On Error GoTo trata
  Set clsEst = New datSisLoc.clsEstoque
  '
  clsEst.ExcluirESTOQUE lngESTOQUEID
  '
  Set clsEst = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

