VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarLoja(intTipoLoja As Integer, _
                               strNome As String, _
                               lngLOJAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objLoja As datSisMetal.clsLoja
  Set objLoja = New datSisMetal.clsLoja
  
  'Listar Ficha Cliente
  Set SelecionarLoja = objLoja.SelecionarLoja(intTipoLoja, _
                                              strNome, _
                                              lngLOJAID)
  '
  Set objLoja = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarFornecedorPeloNome(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objLoja As datSisMetal.clsLoja
  Set objLoja = New datSisMetal.clsLoja

  'Listar Loja
  Set SelecionarFornecedorPeloNome = objLoja.SelecionarFornecedorPeloNome(strNome)
  '
  Set objLoja = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarLojaPeloNome(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objLoja As datSisMetal.clsLoja
  Set objLoja = New datSisMetal.clsLoja

  'Listar Loja
  Set SelecionarLojaPeloNome = objLoja.SelecionarLojaPeloNome(strNome)
  '
  Set objLoja = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarLojaPeloPkid(lngLOJAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objLoja As datSisMetal.clsLoja
  Set objLoja = New datSisMetal.clsLoja

  'Listar Loja
  Set SelecionarLojaPeloPkid = objLoja.SelecionarLojaPeloPkid(lngLOJAID)
  '
  Set objLoja = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirLoja(lngLOJAID As Long)
  On Error GoTo trata
  Dim objLoja As datSisMetal.clsLoja
  Set objLoja = New datSisMetal.clsLoja

  'Excluir Loja
  objLoja.ExcluirLoja lngLOJAID
  '
  Set objLoja = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarFornecedor(lngPKID As Long, _
                             strValorKg As String)

  On Error GoTo trata
  Dim objLoja     As datSisMetal.clsLoja
  '
  Set objLoja = New datSisMetal.clsLoja
  'Alterar Fornecedor
  objLoja.AlterarFornecedor lngPKID, _
                            strValorKg
  Set objLoja = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirFornecedor(ByRef lngPKID As Long, _
                             strValorKg As String)

  On Error GoTo trata
  Dim objLoja     As datSisMetal.clsLoja
  '
  Set objLoja = New datSisMetal.clsLoja
  'Inserir Fornecedor
  objLoja.InserirFornecedor lngPKID, _
                            strValorKg
  Set objLoja = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirFilial(ByRef lngPKID As Long)

  On Error GoTo trata
  Dim objLoja     As datSisMetal.clsLoja
  '
  Set objLoja = New datSisMetal.clsLoja
  'Inserir Filial
  objLoja.InserirFilial lngPKID
  Set objLoja = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub InserirEmpresa(ByRef lngPKID As Long)

  On Error GoTo trata
  Dim objLoja     As datSisMetal.clsLoja
  '
  Set objLoja = New datSisMetal.clsLoja
  'Inserir Empresa
  objLoja.InserirEmpresa lngPKID
  Set objLoja = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirAnodizadora(ByRef lngPKID As Long)

  On Error GoTo trata
  Dim objLoja     As datSisMetal.clsLoja
  '
  Set objLoja = New datSisMetal.clsLoja
  'Inserir Anodizadora
  objLoja.InserirAnodizadora lngPKID
  Set objLoja = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub InserirFabrica(ByRef lngPKID As Long)

  On Error GoTo trata
  Dim objLoja     As datSisMetal.clsLoja
  '
  Set objLoja = New datSisMetal.clsLoja
  'Inserir Fabrica
  objLoja.InserirFabrica lngPKID
  Set objLoja = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

    
Public Sub InserirLoja(ByRef lngPKID As Long, _
                           strNome As String, _
                           strNomeFantasia As String, _
                           strCnpj As String, _
                           strInscrEstd As String, _
                           strInscrMunic As String, _
                           strTel1 As String, strTel2 As String, strTel3 As String, strFax As String, _
                           strEmail As String, _
                           strContato As String, _
                           strTelContato As String, _
                           strEndRua As String, _
                           strEndNumero As String, _
                           strEndCompl As String, _
                           strEndEstado As String, _
                           strEndCep As String, _
                           strEndBairro As String, _
                           strEndCidade As String, _
                           strStatus As String)
  On Error GoTo trata
  Dim objLoja     As datSisMetal.clsLoja
  '
  Set objLoja = New datSisMetal.clsLoja
  'Inserir Loja
  objLoja.InserirLoja lngPKID, _
                              strNome, _
                              strNomeFantasia, _
                              strCnpj, _
                              strInscrEstd, _
                              strInscrMunic, _
                              strTel1, strTel2, strTel3, strFax, _
                              strEmail, _
                              strContato, _
                              strTelContato, _
                              strEndRua, _
                              strEndNumero, _
                              strEndCompl, _
                              strEndEstado, _
                              strEndCep, _
                              strEndBairro, _
                              strEndCidade, _
                              strStatus
  Set objLoja = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarLoja(lngPKID As Long, _
                           strNome As String, _
                           strNomeFantasia As String, _
                           strCnpj As String, _
                           strInscrEstd As String, _
                           strInscrMunic As String, _
                           strTel1 As String, strTel2 As String, strTel3 As String, strFax As String, _
                           strEmail As String, _
                           strContato As String, _
                           strTelContato As String, _
                           strEndRua As String, _
                           strEndNumero As String, _
                           strEndCompl As String, _
                           strEndEstado As String, _
                           strEndCep As String, _
                           strEndBairro As String, _
                           strEndCidade As String, _
                           strStatus As String)

  On Error GoTo trata
  Dim objLoja     As datSisMetal.clsLoja
  '
  Set objLoja = New datSisMetal.clsLoja
  'Atualizar Loja
  objLoja.AlterarLoja lngPKID, _
                              strNome, _
                              strNomeFantasia, _
                              strCnpj, _
                              strInscrEstd, _
                              strInscrMunic, _
                              strTel1, strTel2, strTel3, strFax, _
                              strEmail, _
                              strContato, _
                              strTelContato, _
                              strEndRua, _
                              strEndNumero, _
                              strEndCompl, _
                              strEndEstado, _
                              strEndCep, _
                              strEndBairro, _
                              strEndCidade, _
                              strStatus
  '
  Set objLoja = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

