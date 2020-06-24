VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProdutoConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarProdutoConvenioPeloPkid(lngProdutoConvenioId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objProdutoConvenio As datApler.clsProdutoConvenio
  Set objProdutoConvenio = New datApler.clsProdutoConvenio
  
  'Listar ProdutoConvenio
  Set SelecionarProdutoConvenioPeloPkid = objProdutoConvenio.SelecionarProdutoConvenioPeloPkid(lngProdutoConvenioId)
  '
  Set objProdutoConvenio = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirProdutoConvenio(lngProdutoConvenioId As Long)
  On Error GoTo trata
  Dim objProdutoConvenio As datApler.clsProdutoConvenio
  Set objProdutoConvenio = New datApler.clsProdutoConvenio
  
  'Excluir ProdutoConvenio
  objProdutoConvenio.ExcluirProdutoConvenio lngProdutoConvenioId
  '
  Set objProdutoConvenio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirProdutoConvenio(lngCONVENIOID As Long, _
                                  strDescricao As String, _
                                  strPercComissao As String, _
                                  strPercVitalicio As String, _
                                  strDataDesativacao As String)
  On Error GoTo trata
  Dim objProdutoConvenio     As datApler.clsProdutoConvenio
  '
  Set objProdutoConvenio = New datApler.clsProdutoConvenio
  'Inserir ProdutoConvenio
  objProdutoConvenio.InserirProdutoConvenio lngCONVENIOID, _
                                            strDescricao, _
                                            strPercComissao, _
                                            strPercVitalicio, _
                                            strDataDesativacao
  '
  Set objProdutoConvenio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarProdutoConvenio(lngPKID As Long, _
                                  strDescricao As String, _
                                  strPercComissao As String, _
                                  strPercVitalicio As String, _
                                  strDataDesativacao As String, _
                                  strStatus As String)

  On Error GoTo trata
  Dim objProdutoConvenio     As datApler.clsProdutoConvenio
  '
  Set objProdutoConvenio = New datApler.clsProdutoConvenio
  'Atualizar ProdutoConvenio
  objProdutoConvenio.AlterarProdutoConvenio lngPKID, _
                                            strDescricao, _
                                            strPercComissao, _
                                            strPercVitalicio, _
                                            strDataDesativacao, _
                                            strStatus
  '
  Set objProdutoConvenio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

