VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function CapturaProduto(ByVal strCodigo As String) As ADODB.Recordset
  '
  On Error GoTo trata
    
  Dim objInsumo As datSisMetal.clsInsumo
  '
  Set objInsumo = New datSisMetal.clsInsumo
  
  Set CapturaProduto = objInsumo.CapturaProduto(strCodigo)
  '
  Set objInsumo = Nothing
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarProdutoPeloCodigo(strCodigo As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objInsumo As datSisMetal.clsInsumo
  Set objInsumo = New datSisMetal.clsInsumo

  'Listar Insumo
  Set SelecionarProdutoPeloCodigo = objInsumo.SelecionarProdutoPeloCodigo(strCodigo)
  '
  Set objInsumo = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarInsumoPeloCodigo(strCodigo As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objInsumo As datSisMetal.clsInsumo
  Set objInsumo = New datSisMetal.clsInsumo

  'Listar Insumo
  Set SelecionarInsumoPeloCodigo = objInsumo.SelecionarInsumoPeloCodigo(strCodigo)
  '
  Set objInsumo = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarInsumoPeloPkid(lngINSUMOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objInsumo As datSisMetal.clsInsumo
  Set objInsumo = New datSisMetal.clsInsumo

  'Listar Insumo
  Set SelecionarInsumoPeloPkid = objInsumo.SelecionarInsumoPeloPkid(lngINSUMOID)
  '
  Set objInsumo = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function SelecionarProdutoPeloPkid(lngINSUMOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objInsumo As datSisMetal.clsInsumo
  Set objInsumo = New datSisMetal.clsInsumo

  'Listar Insumo
  Set SelecionarProdutoPeloPkid = objInsumo.SelecionarProdutoPeloPkid(lngINSUMOID)
  '
  Set objInsumo = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirInsumo(lngINSUMOID As Long)
  On Error GoTo trata
  Dim objInsumo As datSisMetal.clsInsumo
  Set objInsumo = New datSisMetal.clsInsumo

  'Excluir Insumo
  objInsumo.ExcluirInsumo lngINSUMOID
  '
  Set objInsumo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarAcessorio(lngPKID As Long, _
                            lngGRUPOID As Long, _
                            lngEMBALAGEMID As Long, _
                            strNome As String, _
                            strValor As String, _
                            strQtdMinima As String, _
                            strQtdEstoque As String)

  On Error GoTo trata
  Dim objInsumo     As datSisMetal.clsInsumo
  '
  Set objInsumo = New datSisMetal.clsInsumo
  'Alterar Acessorio
  objInsumo.AlterarAcessorio lngPKID, _
                             lngGRUPOID, _
                             lngEMBALAGEMID, _
                             strNome, _
                             strValor, _
                             strQtdMinima, _
                             strQtdEstoque
  Set objInsumo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub AlterarProduto(lngPKID As Long, _
                          lngGRUPOPRODUTOID As Long, _
                          lngEMBALAGEMPRODUTOID As Long, _
                          lngFORNECEDORID As Long, _
                          lngFAMILIAID As Long, _
                          lngIPIID As Long, _
                          lngICMSID As Long, _
                          strNomeProduto As String, _
                          strNomeAbrevProduto As String, _
                          strValorProduto As String, _
                          strPeso As String, _
                          strAltura As String, _
                          strLargura As String, _
                          strModRef As String, _
                          strTabela As String, _
                          strFinancVenda As String, _
                          strEstoqueMinimo As String, _
                          strMargemEstoque As String, _
                          strSaldoEstoque As String, _
                          strCustoProduto As String, _
                          strMargemAjuste As String, _
                          strPrecoVenda As String, _
                          strTAM As String, _
                          strPAD As String, _
                          strSOB As String, strComissaoVendedor As String)

  On Error GoTo trata
  Dim objInsumo     As datSisMetal.clsInsumo
  '
  Set objInsumo = New datSisMetal.clsInsumo
  'Alterar Acessorio
  objInsumo.AlterarProduto lngPKID, _
                          lngGRUPOPRODUTOID, _
                          lngEMBALAGEMPRODUTOID, _
                          lngFORNECEDORID, _
                          lngFAMILIAID, _
                          lngIPIID, _
                          lngICMSID, _
                          strNomeProduto, _
                          strNomeAbrevProduto, _
                          strValorProduto, _
                          strPeso, _
                          strAltura, _
                          strLargura, _
                          strModRef, _
                          strTabela, _
                          strFinancVenda, _
                          strEstoqueMinimo, _
                          strMargemEstoque, _
                          strSaldoEstoque, _
                          strCustoProduto, _
                          strMargemAjuste, _
                          strPrecoVenda, _
                          strTAM, _
                          strPAD, _
                          strSOB, strComissaoVendedor
  Set objInsumo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarPerfil(lngPKID As Long, _
                         lngLINHAID As Long, _
                         lngCORID As Long, _
                         strPesoMinimo As String, _
                         strPesoEstoque As String)

  On Error GoTo trata
  Dim objInsumo     As datSisMetal.clsInsumo
  '
  Set objInsumo = New datSisMetal.clsInsumo
  'Alterar Perfil
  objInsumo.AlterarPerfil lngPKID, _
                          lngLINHAID, _
                          lngCORID, _
                          strPesoMinimo, _
                          strPesoEstoque
  Set objInsumo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirAcessorio(lngPKID As Long, _
                            lngGRUPOID As Long, _
                            lngEMBALAGEMID As Long, _
                            strNome As String, _
                            strValor As String, _
                            strQtdMinima As String, _
                            strQtdEstoque As String)

  On Error GoTo trata
  Dim objInsumo     As datSisMetal.clsInsumo
  '
  Set objInsumo = New datSisMetal.clsInsumo
  'Inserir Acessorio
  objInsumo.InserirAcessorio lngPKID, _
                             lngGRUPOID, _
                             lngEMBALAGEMID, _
                             strNome, _
                             strValor, _
                             strQtdMinima, _
                             strQtdEstoque
  Set objInsumo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub InserirProduto(lngPKID As Long, _
                          lngGRUPOPRODUTOID As Long, _
                          lngEMBALAGEMPRODUTOID As Long, _
                          lngFORNECEDORID As Long, _
                          lngFAMILIAID As Long, _
                          lngIPIID As Long, _
                          lngICMSID As Long, _
                          strNomeProduto As String, _
                          strNomeAbrevProduto As String, _
                          strValorProduto As String, _
                          strPeso As String, _
                          strAltura As String, _
                          strLargura As String, _
                          strModRef As String, _
                          strTabela As String, _
                          strFinancVenda As String, _
                          strEstoqueMinimo As String, _
                          strMargemEstoque As String, _
                          strSaldoEstoque As String, _
                          strCustoProduto As String, _
                          strMargemAjuste As String, _
                          strPrecoVenda As String, _
                          strTAM As String, _
                          strPAD As String, _
                          strSOB As String, strComissaoVendedor As String)

  On Error GoTo trata
  Dim objInsumo     As datSisMetal.clsInsumo
  '
  Set objInsumo = New datSisMetal.clsInsumo
  'Inserir Acessorio
  objInsumo.InserirProduto lngPKID, _
                          lngGRUPOPRODUTOID, _
                          lngEMBALAGEMPRODUTOID, _
                          lngFORNECEDORID, _
                          lngFAMILIAID, _
                          lngIPIID, _
                          lngICMSID, _
                          strNomeProduto, _
                          strNomeAbrevProduto, _
                          strValorProduto, _
                          strPeso, _
                          strAltura, _
                          strLargura, _
                          strModRef, _
                          strTabela, _
                          strFinancVenda, _
                          strEstoqueMinimo, _
                          strMargemEstoque, _
                          strSaldoEstoque, _
                          strCustoProduto, _
                          strMargemAjuste, _
                          strPrecoVenda, _
                          strTAM, _
                          strPAD, _
                          strSOB, strComissaoVendedor
  Set objInsumo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirPerfil(lngPKID As Long, _
                         lngLINHAID As Long, _
                         lngCORID As Long, _
                         strPesoMinimo As String, _
                         strPesoEstoque As String)

  On Error GoTo trata
  Dim objInsumo     As datSisMetal.clsInsumo
  '
  Set objInsumo = New datSisMetal.clsInsumo
  'Inserir Perfil
  objInsumo.InserirPerfil lngPKID, _
                          lngLINHAID, _
                          lngCORID, _
                          strPesoMinimo, _
                          strPesoEstoque
  Set objInsumo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirInsumo(ByRef lngPKID As Long, _
                         strCodigo As String, _
                         strStatus As String)
  On Error GoTo trata
  Dim objInsumo     As datSisMetal.clsInsumo
  '
  Set objInsumo = New datSisMetal.clsInsumo
  'Inserir Insumo
  objInsumo.InserirInsumo lngPKID, _
                          strCodigo, _
                          strStatus
  Set objInsumo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarInsumo(lngPKID As Long, _
                         strCodigo As String, _
                         strStatus As String)

  On Error GoTo trata
  Dim objInsumo     As datSisMetal.clsInsumo
  '
  Set objInsumo = New datSisMetal.clsInsumo
  'Atualizar Insumo
  objInsumo.AlterarInsumo lngPKID, _
                          strCodigo, _
                          strStatus
  '
  Set objInsumo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

