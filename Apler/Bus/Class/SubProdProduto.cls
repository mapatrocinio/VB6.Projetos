VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubProdProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarSubProdProdutoPeloPkid(lngSubProdProdutoId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objSubProdProduto As datApler.clsSubProdProduto
  Set objSubProdProduto = New datApler.clsSubProdProduto
  
  'Listar SubProdProduto
  Set SelecionarSubProdProdutoPeloPkid = objSubProdProduto.SelecionarSubProdProdutoPeloPkid(lngSubProdProdutoId)
  '
  Set objSubProdProduto = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirSubProdProduto(lngSubProdProdutoId As Long)
  On Error GoTo trata
  Dim objSubProdProduto As datApler.clsSubProdProduto
  Set objSubProdProduto = New datApler.clsSubProdProduto
  
  'Excluir SubProdProduto
  objSubProdProduto.ExcluirSubProdProduto lngSubProdProdutoId
  '
  Set objSubProdProduto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirSubProdProduto(lngPRODUTOID As Long, _
                                 strDescricao As String)
  On Error GoTo trata
  Dim objSubProdProduto     As datApler.clsSubProdProduto
  '
  Set objSubProdProduto = New datApler.clsSubProdProduto
  'Inserir SubProdProduto
  objSubProdProduto.InserirSubProdProduto lngPRODUTOID, _
                                          strDescricao
  '
  Set objSubProdProduto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarSubProdProduto(lngPKID As Long, _
                                 strDescricao As String, _
                                 strDataDesativacao As String, _
                                 strStatus As String)

  On Error GoTo trata
  Dim objSubProdProduto     As datApler.clsSubProdProduto
  '
  Set objSubProdProduto = New datApler.clsSubProdProduto
  'Atualizar SubProdProduto
  objSubProdProduto.AlterarSubProdProduto lngPKID, _
                                          strDescricao, _
                                          strDataDesativacao, _
                                          strStatus
  '
  Set objSubProdProduto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

