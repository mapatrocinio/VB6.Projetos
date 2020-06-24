VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItemNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarItemNFContrato(lngContratoId As Long, _
                                         lngObraId As Long, _
                                         lngESTOQUEID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objItemNF As datSisLoc.clsItemNF
  Set objItemNF = New datSisLoc.clsItemNF
  
  'Listar Item NF
  Set SelecionarItemNFContrato = objItemNF.SelecionarItemNFContrato(lngContratoId, _
                                                                    lngObraId, _
                                                                    lngESTOQUEID)
  '
  Set objItemNF = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub AlterarEstoquePeloRetItemNF(ByVal lngESTOQUEID As Long, _
                                       ByVal strQuantidade As String)
           
  On Error GoTo trata
  Dim objItemNF     As datSisLoc.clsItemNF
  '
  Set objItemNF = New datSisLoc.clsItemNF
  'Inserir ITEMNF
  objItemNF.AlterarEstoquePeloRetItemNF lngESTOQUEID, _
                                        strQuantidade
  '
  Set objItemNF = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
                                    
Public Sub AlterarEstoquePeloItemNF(ByVal lngESTOQUEID As Long, _
                                    ByVal strQuantidade As String)
           
  On Error GoTo trata
  Dim objItemNF     As datSisLoc.clsItemNF
  '
  Set objItemNF = New datSisLoc.clsItemNF
  'Inserir ITEMNF
  objItemNF.AlterarEstoquePeloItemNF lngESTOQUEID, _
                                     strQuantidade
  '
  Set objItemNF = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
                                    
                                    
Public Sub InserirITEMNF(lngNFID As Long, _
                         lngESTOQUEID As Long, _
                         strQuantidade As String, _
                         strValor As String, _
                         strValorEstoque As String, _
                         strLargura As String, _
                         strAltura As String, _
                         strValorLanc As String)
  On Error GoTo trata
  Dim objItemNF     As datSisLoc.clsItemNF
  '
  Set objItemNF = New datSisLoc.clsItemNF
  'Inserir ITEMNF
  objItemNF.InserirITEMNF lngNFID, _
                          lngESTOQUEID, _
                          strQuantidade, _
                          strValor, _
                          strValorEstoque, _
                          strLargura, _
                          strAltura, _
                          strValorLanc
  '
  Set objItemNF = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub



Public Sub ExcluirITEMNF(lngITEMNFID As Long)
  On Error GoTo trata
  Dim objItemNF As datSisLoc.clsItemNF
  Set objItemNF = New datSisLoc.clsItemNF
  
  'Excluir ITEMNF
  objItemNF.ExcluirITEMNF lngITEMNFID
  '
  Set objItemNF = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
'''
'''Public Function SelecionarNFTotalProc(lngNFID As Long) As ADODB.Recordset
'''  On Error GoTo trata
'''  Dim objItemNF As datSisLoc.clsItemNF
'''  Set objItemNF = New datSisLoc.clsItemNF
'''
'''  'Listar NF
'''  Set SelecionarNFTotalProc = objItemNF.SelecionarNFTotalProc(lngNFID)
'''  '
'''  Set objItemNF = Nothing
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Function
'''
