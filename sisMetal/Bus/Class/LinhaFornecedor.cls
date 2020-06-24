VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLinhaFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarLinhaFornecedorPeloPkid(lngLinhaFornecedorId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objLinhaFornecedor As datSisMetal.clsLinhaFornecedor
  Set objLinhaFornecedor = New datSisMetal.clsLinhaFornecedor
  
  'Listar LinhaFornecedor
  Set SelecionarLinhaFornecedorPeloPkid = objLinhaFornecedor.SelecionarLinhaFornecedorPeloPkid(lngLinhaFornecedorId)
  '
  Set objLinhaFornecedor = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirLinhaFornecedor(lngLinhaFornecedorId As Long)
  On Error GoTo trata
  Dim objLinhaFornecedor As datSisMetal.clsLinhaFornecedor
  Set objLinhaFornecedor = New datSisMetal.clsLinhaFornecedor
  
  'Excluir LinhaFornecedor
  objLinhaFornecedor.ExcluirLinhaFornecedor lngLinhaFornecedorId
  '
  Set objLinhaFornecedor = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirLinhaFornecedor(lngLINHAID As Long, _
                                  lngFORNECEDORID As Long, _
                                  strCodigo As String, _
                                  strPeso As String)
  On Error GoTo trata
  Dim objLinhaFornecedor     As datSisMetal.clsLinhaFornecedor
  '
  Set objLinhaFornecedor = New datSisMetal.clsLinhaFornecedor
  'Inserir LinhaFornecedor
  objLinhaFornecedor.InserirLinhaFornecedor lngLINHAID, _
                                            lngFORNECEDORID, _
                                            strCodigo, _
                                            strPeso
  '
  Set objLinhaFornecedor = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarLinhaFornecedor(lngPKID As Long, _
                                  strCodigo As String, _
                                  strPeso As String)

  On Error GoTo trata
  Dim objLinhaFornecedor     As datSisMetal.clsLinhaFornecedor
  '
  Set objLinhaFornecedor = New datSisMetal.clsLinhaFornecedor
  'Atualizar LinhaFornecedor
  objLinhaFornecedor.AlterarLinhaFornecedor lngPKID, _
                                            strCodigo, _
                                            strPeso
  '
  Set objLinhaFornecedor = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

