VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEntradaMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirTAB_ENTRADAMATERIAL(ByVal lngTAB_ENTRADAMATERIALID As Long, _
                                      ByVal lngPRODUTOID As Long, _
                                      ByVal lngQUANTIDADENOVA As Long)
  Dim clsEntMat    As datSisMetal.clsEntradaMaterial
  '
  On Error GoTo trata
  Set clsEntMat = New datSisMetal.clsEntradaMaterial
  '
  clsEntMat.ExcluirTAB_ENTRADAMATERIAL lngTAB_ENTRADAMATERIALID, _
                                       lngPRODUTOID, _
                                       lngQUANTIDADENOVA
  '
  Set clsEntMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirEntradaMaterial(strDescricao As String, _
                                  strUsuario As String, _
                                  lngDOCUMENTOENTRADAID As Long, _
                                  lngFORNECEDORID As Long, _
                                  strDataAquisicao As String, _
                                  strNumero As String, _
                                  ByRef lngENTRADAMATERIALID, _
                                  ByRef strData As String, _
                                  ByRef strCodigo As String)
  On Error GoTo trata
  
  Dim objEntMat   As datSisMetal.clsEntradaMaterial
  Dim strSql      As String
  Dim objRs       As ADODB.Recordset
  '
  strData = Format(Now, "DD/MM/YYYY hh:mm")
  '
  Set objEntMat = New datSisMetal.clsEntradaMaterial
  strCodigo = objEntMat.CapturaCodigoEntradaMaterial
  'Inserir Entrada de material
  objEntMat.InserirEntradaMaterial strData, _
                                   strDescricao, _
                                   strCodigo, _
                                   strUsuario, _
                                   lngDOCUMENTOENTRADAID, _
                                   lngFORNECEDORID, _
                                   strDataAquisicao, _
                                   strNumero
  '
  'Após inserir a entrada de material, Retornar PKID pelo código
  Set objRs = objEntMat.CapturaEntradaMaterial(strCodigo)
  If Not objRs.EOF Then
    lngENTRADAMATERIALID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set objEntMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarDocumentoEntradaPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  Dim clsEntMat As datSisMetal.clsEntradaMaterial
  Set clsEntMat = New datSisMetal.clsEntradaMaterial
  
  'Listar Documento de entrada
  Set ListarDocumentoEntradaPelaDesc = clsEntMat.ListarDocumentoEntradaPelaDesc(strDescricao)
  '
  Set clsEntMat = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function ValidarExclusaoEntradaMaterial(lngENTRADAMATERIALID As Long) As String
  On Error GoTo trata
  Dim strSql      As String
  Dim strRetorno  As String
  Dim objRs       As ADODB.Recordset
  Dim clsGer      As datSisMetal.clsGeral
  '
  Set clsGer = New datSisMetal.clsGeral
  strRetorno = ""
  'Valida entrada de material
  'TAB_ENTRADAMATERIAL
  strSql = "Select count(*) AS TOTAL from TAB_ENTRADAMATERIAL WHERE ENTRADAMATERIALID = " & lngENTRADAMATERIALID
  Set objRs = clsGer.ExecutarSQL(strSql)
  '
  If Not objRs.EOF Then
    If objRs.Fields("TOTAL").Value > 0 Then
      strRetorno = "Itens de entrada de material"
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  '
  ValidarExclusaoEntradaMaterial = strRetorno
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub InserirDocumentoEntrada(strNome As String)

  On Error GoTo trata
  Dim clsEntMat  As datSisMetal.clsEntradaMaterial
  '
  Set clsEntMat = New datSisMetal.clsEntradaMaterial
  
  'Inserir Documento de entrada
  clsEntMat.InserirDocumentoEntrada strNome
  '
  Set clsEntMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarEntradaMaterial(lngENTRADAMATERIALID As Long, _
                                  strDescricao As String, _
                                  lngDOCUMENTOENTRADAID As Long, _
                                  lngFORNECEDORID As Long, _
                                  strNumero As String, _
                                  strDataAquisicao As String)
                         
  On Error GoTo trata
  Dim clsEntMat As datSisMetal.clsEntradaMaterial
  Set clsEntMat = New datSisMetal.clsEntradaMaterial
  
  'Inserir Entrada de Material
  clsEntMat.AlterarEntradaMaterial lngENTRADAMATERIALID, _
                                   strDescricao, _
                                   lngDOCUMENTOENTRADAID, _
                                   lngFORNECEDORID, _
                                   strNumero, _
                                   strDataAquisicao

  '
  Set clsEntMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarProdutoPelaEntradaMaterial(ByVal lngPRODUTOID As Long, _
                                             ByVal lngQuantidade As Long, _
                                             ByVal strQuantidade As String)
  '
  On Error GoTo trata
    
  Dim objEntMat As datSisMetal.clsEntradaMaterial
  '
  Set objEntMat = New datSisMetal.clsEntradaMaterial
  '
  objEntMat.AlterarProdutoPelaEntradaMaterial lngPRODUTOID, _
                                              lngQuantidade, _
                                              strQuantidade
  '
  Set objEntMat = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirTAB_ENTRADAMATERIAL(ByVal lngENTRADAMATERIALID As Long, _
                                      ByVal lngPRODUTOID As Long, _
                                      ByVal strQuantidade As String, _
                                      ByVal strDataValidade As String)
  '
  On Error GoTo trata
    
  Dim clsEntMat As datSisMetal.clsEntradaMaterial
  '
  Set clsEntMat = New datSisMetal.clsEntradaMaterial
  '
  clsEntMat.InserirTAB_ENTRADAMATERIAL lngENTRADAMATERIALID, _
                                       lngPRODUTOID, _
                                       strQuantidade, _
                                       strDataValidade
  '
  Set clsEntMat = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ValidarExclusaoTab_EntradaMaterial(strMsgErro As String, _
                                              strCodigo As String, _
                                              strQuantidadeASerBaixada As String, _
                                              lngNOVAQUANTIDADE As Long)
  On Error GoTo trata
  Dim strSql          As String
  Dim objRs           As ADODB.Recordset
  Dim objProduto      As datSisMetal.clsInsumo
  Dim lngQuantidade   As Long
  '
  Set objProduto = New datSisMetal.clsInsumo
  '
  strMsgErro = ""
  'Pegar item do estoque
  'Verifica se estoque está cadastrado
  Set objRs = objProduto.SelecionarProdutoPeloCodigo(strCodigo)

  If objRs.EOF Then
    strMsgErro = "Código do produto não cadastrado"
  Else
    lngQuantidade = IIf(Not IsNumeric(objRs.Fields("SALDOESTOQUE").Value), 0, objRs.Fields("SALDOESTOQUE").Value)
  End If
  '
  objRs.Close
  Set objRs = Nothing
  '
  If strMsgErro = "" Then
    'Validar e Calcular a nova quantidade
    lngNOVAQUANTIDADE = lngQuantidade - IIf(Not IsNumeric(strQuantidadeASerBaixada), 0, CLng(strQuantidadeASerBaixada))
    If lngNOVAQUANTIDADE < 0 Then
      strMsgErro = "Produto não possui estoque suficiente para ser estornado"
    End If
  End If
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ValidarInclusaoProduto(lngPRODUTOID As Long, _
                                  lngENTRADAMATERIALID As Long, _
                                  strMsgErro As String, _
                                  strCodigo As String, _
                                  lngQuantidade As Long)
  On Error GoTo trata
  Dim strSql      As String
  Dim objRs       As ADODB.Recordset
  Dim objGeral    As datSisMetal.clsGeral
  Dim objProduto  As datSisMetal.clsInsumo
  '
  Set objGeral = New datSisMetal.clsGeral
  Set objProduto = New datSisMetal.clsInsumo
  '
  strMsgErro = ""
  'Pegar item do estoque
  'Verifica se estoque está cadastrado
  Set objRs = objProduto.SelecionarProdutoPeloCodigo(strCodigo)

  If objRs.EOF Then
    strMsgErro = "Código do produto não cadastrado"
  Else
    lngPRODUTOID = objRs.Fields("INSUMOID").Value
    lngQuantidade = IIf(Not IsNumeric(objRs.Fields("SALDOESTOQUE").Value), 0, objRs.Fields("SALDOESTOQUE").Value)
  End If
  '
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarDocumentoEntrada(lngDOCUMENTOENTRADAID As Long, _
                                   strNome As String)
                         
  On Error GoTo trata
  Dim clsEntMat As datSisMetal.clsEntradaMaterial
  Set clsEntMat = New datSisMetal.clsEntradaMaterial
  
  'Inserir Documento de Entrada
  clsEntMat.AlterarDocumentoEntrada lngDOCUMENTOENTRADAID, _
                                    strNome

  '
  Set clsEntMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarEntradaMaterial(lngENTRADAMATERIALID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsEntMat As datSisMetal.clsEntradaMaterial
  Set clsEntMat = New datSisMetal.clsEntradaMaterial
  
  'Listar entrada de material
  Set ListarEntradaMaterial = clsEntMat.ListarEntradaMaterial(lngENTRADAMATERIALID)
  '
  Set clsEntMat = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ListarDocumentoEntrada(lngDOCUMENTOENTRADAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsEntMat As datSisMetal.clsEntradaMaterial
  Set clsEntMat = New datSisMetal.clsEntradaMaterial
  
  'Listar Documento de entrada
  Set ListarDocumentoEntrada = clsEntMat.ListarDocumentoEntrada(lngDOCUMENTOENTRADAID)
  '
  Set clsEntMat = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirDOCUMENTOENTRADA(ByVal lngDOCUMENTOENTRADAID As Long)
  Dim clsEntMat    As datSisMetal.clsEntradaMaterial
  '
  On Error GoTo trata
  Set clsEntMat = New datSisMetal.clsEntradaMaterial
  '
  clsEntMat.ExcluirDOCUMENTOENTRADA lngDOCUMENTOENTRADAID
  '
  Set clsEntMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ExcluirENTRADAMATERIAL(ByVal lngENTRADAMATERIALID As Long)
  Dim clsEntMat    As datSisMetal.clsEntradaMaterial
  '
  On Error GoTo trata
  Set clsEntMat = New datSisMetal.clsEntradaMaterial
  '
  clsEntMat.ExcluirENTRADAMATERIAL lngENTRADAMATERIALID
  '
  Set clsEntMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


