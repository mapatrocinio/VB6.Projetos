VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParceiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Sub DesassociarUsuarioAoParceiro(lngCONTROLEACESSOID As Long, _
                                        lngPARCEIROID As Long)
  On Error GoTo trata
  '
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim objParceiro   As datSisContas.clsParceiro
  '
  Set objParceiro = New datSisContas.clsParceiro
  '
  objParceiro.ExlcuirTab_ControlParc lngCONTROLEACESSOID, _
                                     lngPARCEIROID
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AssociarUsuarioAoParceiro(lngCONTROLEACESSOID As Long, _
                                     lngPARCEIROID As Long)
  On Error GoTo trata
  '
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim objParceiro As datSisContas.clsParceiro
  '
  Set objParceiro = New datSisContas.clsParceiro
  '
  objParceiro.InserirTab_ControlParc lngCONTROLEACESSOID, _
                                     lngPARCEIROID
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'''
'''Public Function ListarTAB_GRUPOESTESTINTERPelaDescCod(strDescGrupoEstoque As String, _
'''                                                      strCodigo As String) As ADODB.Recordset
'''  On Error GoTo trata
'''  Dim clsEstInter As datSisContas.clsEstoqueIntermediario
'''  Set clsEstInter = New datSisContas.clsEstoqueIntermediario
'''
'''  'Listar TAB_GRUPOESTESTINTER Pela descrição e pelo código
'''  Set ListarTAB_GRUPOESTESTINTERPelaDescCod = clsEstInter.ListarTAB_GRUPOESTESTINTERPelaDescCod(strDescGrupoEstoque, _
'''                                                                                                strCodigo)
'''  '
'''  Set clsEstInter = Nothing
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Function
'''
'''Public Sub AlterarTAB_GRUPOEST(lngTAB_GRUPOEST As Long, _
'''                               lngQtdAnterior As Long, _
'''                               lngQtdAtual As Long, _
'''                               lngESTOQUEID As Long, _
'''                               lngQTDESTOQUE As Long, _
'''                               strTipo As String, _
'''                               strOpcao As String)
'''  On Error GoTo trata
'''  Dim clsEstInter As datSisContas.clsEstoqueIntermediario
'''  Dim lngQtdRealEstoque As Long
'''  '
'''  Set clsEstInter = New datSisContas.clsEstoqueIntermediario
'''
'''  'Alterar tab_grupoest
'''  clsEstInter.AlterarTAB_GRUPOEST lngTAB_GRUPOEST, _
'''                                  lngQtdAtual, _
'''                                  strTipo
'''  '
'''  'Alterar ESTOQUE
'''  If strOpcao = "R" Then 'Retornar ao estoque
'''    lngQtdRealEstoque = lngQTDESTOQUE + (lngQtdAnterior - lngQtdAtual)
'''    clsEstInter.AlterarEstoque lngESTOQUEID, _
'''                               lngQtdRealEstoque
'''  End If
'''  '
'''  Set clsEstInter = Nothing
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Sub
'''
'''Public Sub ExcluirESTOQUEINTERMEDIARIO(strTipo As String, _
'''                                       lngTAB_GRUPOEST As Long, _
'''                                       lngESTOQUEINTERMEDIARIOID As Long, _
'''                                       blnTAB_GRUPO As Boolean)
'''
'''  Dim clsEstInter    As datSisContas.clsEstoqueIntermediario
'''  '
'''  On Error GoTo trata
'''  Set clsEstInter = New datSisContas.clsEstoqueIntermediario
'''  '
'''  clsEstInter.ExcluirESTOQUEINTERMEDIARIO strTipo, _
'''                                          lngTAB_GRUPOEST, _
'''                                          lngESTOQUEINTERMEDIARIOID, _
'''                                          blnTAB_GRUPO
'''  '
'''  Set clsEstInter = Nothing
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Sub
'''
'''Public Sub InserirEstoqueIntermediario(strCodigo As String, _
'''                                       strDescricao As String, _
'''                                       strUnidade As String, _
'''                                       lngESTOQUEID As Long)
'''  On Error GoTo trata
'''  Dim clsEstInter As datSisContas.clsEstoqueIntermediario
'''  Set clsEstInter = New datSisContas.clsEstoqueIntermediario
'''
'''  'Inserir Estoque Itermediário
'''  clsEstInter.InserirEstoqueIntermediario strCodigo, _
'''                                          strDescricao, _
'''                                          strUnidade, _
'''                                          lngESTOQUEID
'''  '
'''  Set clsEstInter = Nothing
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Sub
'''
'''Public Function ListarEstoqueIntermediarioPeloCodigo(strCodigo As String) As ADODB.Recordset
'''  On Error GoTo trata
'''  Dim clsEstInter As datSisContas.clsEstoqueIntermediario
'''  Set clsEstInter = New datSisContas.clsEstoqueIntermediario
'''
'''  'Listar Estoque Intermediario Pelo Codigo
'''  Set ListarEstoqueIntermediarioPeloCodigo = clsEstInter.ListarEstoqueIntermediarioPeloCodigo(strCodigo)
'''  '
'''  Set clsEstInter = Nothing
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Function
'''
'''Public Sub DesassociarUnidadeAoGrpEstoque(lngAPARTAMENTOID As Long, lngGRUPOESTOQUEID As Long)
'''  On Error GoTo trata
'''  '
'''  Dim strSql    As String
'''  Dim objRs     As ADODB.Recordset
'''  Dim clsEstInter As datSisContas.clsEstoqueIntermediario
'''  '
'''  Set clsEstInter = New datSisContas.clsEstoqueIntermediario
'''  '
'''  clsEstInter.ExlcuirTab_GrupoEstApto lngAPARTAMENTOID, lngGRUPOESTOQUEID
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Sub
'''
'''
'''Public Sub AssociarUnidadeAoGrpEstoque(lngAPARTAMENTOID As Long, lngGRUPOESTOQUEID As Long)
'''  On Error GoTo trata
'''  '
'''  Dim strSql    As String
'''  Dim objRs     As ADODB.Recordset
'''  Dim clsEstInter As datSisContas.clsEstoqueIntermediario
'''  '
'''  Set clsEstInter = New datSisContas.clsEstoqueIntermediario
'''  '
'''  clsEstInter.InserirTab_GrupoEstApto lngAPARTAMENTOID, lngGRUPOESTOQUEID
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Sub
'''
Public Sub ExcluirParceiro(lngPARCEIROID As Long)

  On Error GoTo trata
  Dim objParceiro As datSisContas.clsParceiro
  Set objParceiro = New datSisContas.clsParceiro

  'Inserir Parceiro
  objParceiro.ExlcuirParceiro lngPARCEIROID
  '
  Set objParceiro = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function VerificaParceiro(lngPARCEIROID As Long) As Boolean
  On Error GoTo trata
  Dim objParceiro As datSisContas.clsParceiro
  Set objParceiro = New datSisContas.clsParceiro

  'Verificar
  VerificaParceiro = objParceiro.VerificaParceiro(lngPARCEIROID)
  '
  Set objParceiro = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub AlterarParceiro(strDescricao As String, _
                           lngPARCEIROID As Long)
  On Error GoTo trata
  Dim objParceiro As datSisContas.clsParceiro
  Set objParceiro = New datSisContas.clsParceiro

  'Alterar Grupo Estoque
  objParceiro.AlterarParceiro strDescricao, lngPARCEIROID
  '
  Set objParceiro = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirParceiro(strDescricao As String)
  On Error GoTo trata
  Dim objParceiro As datSisContas.clsParceiro
  Set objParceiro = New datSisContas.clsParceiro

  'Inserir Grupo Estoque
  objParceiro.InserirParceiro strDescricao
  '
  Set objParceiro = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarParceiroPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objParceiro As datSisContas.clsParceiro
  Set objParceiro = New datSisContas.clsParceiro

  'Listar Parceiro pela descrição
  Set ListarParceiroPelaDesc = objParceiro.ListarParceiroPelaDesc(strDescricao)
  '
  Set objParceiro = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ListarParceiro(lngPARCEIROID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objParceiro As datSisContas.clsParceiro
  Set objParceiro = New datSisContas.clsParceiro

  'Listar Parceiro
  Set ListarParceiro = objParceiro.ListarParceiro(lngPARCEIROID)
  '
  Set objParceiro = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

'''Public Function ValidacamposGrupoEstoque(strDescricao As String, _
'''                                         strTipo As String, _
'''                                         blnPadraoVendas As Boolean, _
'''                                         lngGRUPOESTOQUEID As Long, _
'''                                         strMsgErro As String) As Boolean
'''  On Error GoTo trata
'''  Dim clsEstInter As datSisContas.clsEstoqueIntermediario
'''  Set clsEstInter = New datSisContas.clsEstoqueIntermediario
'''
'''  'Listar Pedido
'''  ValidacamposGrupoEstoque = clsEstInter.ValidacamposGrupoEstoque(strDescricao, strTipo, blnPadraoVendas, lngGRUPOESTOQUEID, strMsgErro)
'''  '
'''  Set clsEstInter = Nothing
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Function
'''
