VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSaidaMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirTAB_SAIDAMATERIAL(ByVal lngTAB_SAIDAMATERIALID As Long, _
                                    ByVal lngPRODUTOID As Long, _
                                    ByVal lngQUANTIDADENOVA As Long)
  Dim clsSaiMat    As datSisMetal.clsSaidaMaterial
  '
  On Error GoTo trata
  Set clsSaiMat = New datSisMetal.clsSaidaMaterial
  '
  clsSaiMat.ExcluirTAB_SAIDAMATERIAL lngTAB_SAIDAMATERIALID, _
                                     lngPRODUTOID, _
                                     lngQUANTIDADENOVA
  '
  Set clsSaiMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirSaidaMaterial(strDescricao As String, _
                                strUsuario As String, _
                                lngDOCUMENTOSAIDAID As Long, _
                                lngFILIALID As Long, _
                                strDataAquisicao As String, _
                                ByRef lngSAIDAMATERIALID, _
                                ByRef strData As String, _
                                ByRef strCodigo As String, _
                                strDataTransacao As String, _
                                strMotivo As String)
  On Error GoTo trata
  
  Dim objEntMat   As datSisMetal.clsSaidaMaterial
  Dim strSql      As String
  Dim objRs       As ADODB.Recordset
  '
  strData = Format(Now, "DD/MM/YYYY hh:mm")
  '
  Set objEntMat = New datSisMetal.clsSaidaMaterial
  strCodigo = objEntMat.CapturaCodigoSaidaMaterial
  'Inserir Saida de material
  objEntMat.InserirSaidaMaterial strData, _
                                 strDescricao, _
                                 strCodigo, _
                                 strUsuario, _
                                 lngDOCUMENTOSAIDAID, _
                                 lngFILIALID, _
                                 strDataAquisicao, _
                                 strDataTransacao, _
                                 strMotivo
  '
  'Após inserir a entrada de material, Retornar PKID pelo código
  Set objRs = objEntMat.CapturaSaidaMaterial(strCodigo)
  If Not objRs.EOF Then
    lngSAIDAMATERIALID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set objEntMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarDocumentoSaidaPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  Dim clsSaiMat As datSisMetal.clsSaidaMaterial
  Set clsSaiMat = New datSisMetal.clsSaidaMaterial
  
  'Listar Documento de entrada
  Set ListarDocumentoSaidaPelaDesc = clsSaiMat.ListarDocumentoSaidaPelaDesc(strDescricao)
  '
  Set clsSaiMat = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function ValidarExclusaoSaidaMaterial(lngSAIDAMATERIALID As Long) As String
  On Error GoTo trata
  Dim strSql      As String
  Dim strRetorno  As String
  Dim objRs       As ADODB.Recordset
  Dim clsGer      As datSisMetal.clsGeral
  '
  Set clsGer = New datSisMetal.clsGeral
  strRetorno = ""
  'Valida entrada de material
  'TAB_SAIDAMATERIAL
  strSql = "Select count(*) AS TOTAL from TAB_SAIDAMATERIAL WHERE SAIDAMATERIALID = " & lngSAIDAMATERIALID
  Set objRs = clsGer.ExecutarSQL(strSql)
  '
  If Not objRs.EOF Then
    If objRs.Fields("TOTAL").Value > 0 Then
      strRetorno = "Itens de saída de material"
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  '
  ValidarExclusaoSaidaMaterial = strRetorno
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub InserirDocumentoSaida(strNome As String)

  On Error GoTo trata
  Dim clsSaiMat  As datSisMetal.clsSaidaMaterial
  '
  Set clsSaiMat = New datSisMetal.clsSaidaMaterial
  
  'Inserir Documento de entrada
  clsSaiMat.InserirDocumentoSaida strNome
  '
  Set clsSaiMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarSaidaMaterial(lngSAIDAMATERIALID As Long, _
                                strDescricao As String, _
                                lngDOCUMENTOSAIDAID As Long, _
                                lngFILIALID As Long, _
                                strDataAquisicao As String, _
                                strDataTransacao As String, _
                                strMotivo As String)
                         
  On Error GoTo trata
  Dim clsSaiMat As datSisMetal.clsSaidaMaterial
  Set clsSaiMat = New datSisMetal.clsSaidaMaterial
  
  'Inserir Saida de Material
  clsSaiMat.AlterarSaidaMaterial lngSAIDAMATERIALID, _
                                 strDescricao, _
                                 lngDOCUMENTOSAIDAID, _
                                 lngFILIALID, _
                                strDataAquisicao, _
                                strDataTransacao, _
                                strMotivo

  '
  Set clsSaiMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarProdutoPelaSaidaMaterial(ByVal lngPRODUTOID As Long, _
                                           ByVal lngQuantidade As Long, _
                                           ByVal strQuantidade As String, _
                                           ByVal strDocumento As String)
  '
  On Error GoTo trata
    
  Dim objEntMat As datSisMetal.clsSaidaMaterial
  '
  Set objEntMat = New datSisMetal.clsSaidaMaterial
  '
  objEntMat.AlterarProdutoPelaSaidaMaterial lngPRODUTOID, _
                                            lngQuantidade, _
                                            strQuantidade, _
                                            strDocumento
  '
  Set objEntMat = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirTAB_SAIDAMATERIAL(ByVal lngSAIDAMATERIALID As Long, _
                                    ByVal lngPRODUTOID As Long, _
                                    ByVal strQuantidade As String)
  '
  On Error GoTo trata
    
  Dim clsSaiMat As datSisMetal.clsSaidaMaterial
  '
  Set clsSaiMat = New datSisMetal.clsSaidaMaterial
  '
  clsSaiMat.InserirTAB_SAIDAMATERIAL lngSAIDAMATERIALID, _
                                     lngPRODUTOID, _
                                     strQuantidade
  '
  Set clsSaiMat = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ValidarExclusaoTab_SaidaMaterial(strMsgErro As String, _
                                            strCodigo As String, _
                                            strQuantidadeASerBaixada As String, _
                                            lngNOVAQUANTIDADE As Long, _
                                            strDocumento As String)
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
    If strDocumento = "TRANSFERÊNCIA DE ENTRADA" Then
      lngNOVAQUANTIDADE = lngQuantidade - IIf(Not IsNumeric(strQuantidadeASerBaixada), 0, CLng(strQuantidadeASerBaixada))
    Else
      lngNOVAQUANTIDADE = lngQuantidade + IIf(Not IsNumeric(strQuantidadeASerBaixada), 0, CLng(strQuantidadeASerBaixada))
    End If
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
                                  lngSAIDAMATERIALID As Long, _
                                  strMsgErro As String, _
                                  strCodigo As String, _
                                  lngQuantidade As Long, _
                                  strQuantidadeASerBaixada As String, _
                                  strDocumento As String)
  On Error GoTo trata
  Dim strSql              As String
  Dim objRs               As ADODB.Recordset
  Dim objGeral            As datSisMetal.clsGeral
  Dim objProduto          As datSisMetal.clsInsumo
  Dim lngNOVAQUANTIDADE   As Long
  '
  Set objGeral = New datSisMetal.clsGeral
  Set objProduto = New datSisMetal.clsInsumo
  '
  strMsgErro = ""
  'Pegar item do estoque
  'Verifica se estoque está cadastrado
  '''Set objRs = objProduto.SelecionarProdutoPeloCodigo(strCodigo)
  strSql = "SELECT PRODUTO.*, INSUMO.CODIGO " & _
    " FROM PRODUTO INNER JOIN INSUMO ON INSUMO.PKID = PRODUTO.INSUMOID " & _
    "WHERE (NOME LIKE '%" & strCodigo & "%' " & _
        " OR CODIGO LIKE '%" & strCodigo & "%') " & _
   " ORDER BY PRODUTO.NOME;"
  Set objRs = objGeral.ExecutarSQL(strSql)
  
  If objRs.EOF Then
    strMsgErro = "Código do produto não cadastrado"
  Else
    If objRs.RecordCount > 1 Then
      strMsgErro = "Existem mais de um produto cadastrado para este código/descrição"
    Else
      lngPRODUTOID = objRs.Fields("INSUMOID").Value
      lngQuantidade = IIf(Not IsNumeric(objRs.Fields("SALDOESTOQUE").Value), 0, objRs.Fields("SALDOESTOQUE").Value)
    End If
  End If
  '
  If strMsgErro = "" Then
    'Validar e Calcular a nova quantidade
    If strDocumento = "TRANSFERÊNCIA DE ENTRADA" Then
      lngNOVAQUANTIDADE = lngQuantidade + IIf(Not IsNumeric(strQuantidadeASerBaixada), 0, CLng(strQuantidadeASerBaixada))
    Else
      lngNOVAQUANTIDADE = lngQuantidade - IIf(Not IsNumeric(strQuantidadeASerBaixada), 0, CLng(strQuantidadeASerBaixada))
    End If
    
    If lngNOVAQUANTIDADE < 0 Then
      strMsgErro = "Produto não possui estoque suficiente para incluir sua saída"
    End If
  End If
  '
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarDocumentoSaida(lngDOCUMENTOSAIDAID As Long, _
                                   strNome As String)
                         
  On Error GoTo trata
  Dim clsSaiMat As datSisMetal.clsSaidaMaterial
  Set clsSaiMat = New datSisMetal.clsSaidaMaterial
  
  'Inserir Documento de Saida
  clsSaiMat.AlterarDocumentoSaida lngDOCUMENTOSAIDAID, _
                                    strNome

  '
  Set clsSaiMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarSaidaMaterial(lngSAIDAMATERIALID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsSaiMat As datSisMetal.clsSaidaMaterial
  Set clsSaiMat = New datSisMetal.clsSaidaMaterial
  
  'Listar entrada de material
  Set ListarSaidaMaterial = clsSaiMat.ListarSaidaMaterial(lngSAIDAMATERIALID)
  '
  Set clsSaiMat = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ListarDocumentoSaida(lngDOCUMENTOSAIDAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsSaiMat As datSisMetal.clsSaidaMaterial
  Set clsSaiMat = New datSisMetal.clsSaidaMaterial
  
  'Listar Documento de entrada
  Set ListarDocumentoSaida = clsSaiMat.ListarDocumentoSaida(lngDOCUMENTOSAIDAID)
  '
  Set clsSaiMat = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirDOCUMENTOSAIDA(ByVal lngDOCUMENTOSAIDAID As Long)
  Dim clsSaiMat    As datSisMetal.clsSaidaMaterial
  '
  On Error GoTo trata
  Set clsSaiMat = New datSisMetal.clsSaidaMaterial
  '
  clsSaiMat.ExcluirDOCUMENTOSAIDA lngDOCUMENTOSAIDAID
  '
  Set clsSaiMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ExcluirSAIDAMATERIAL(ByVal lngSAIDAMATERIALID As Long)
  Dim clsSaiMat    As datSisMetal.clsSaidaMaterial
  '
  On Error GoTo trata
  Set clsSaiMat = New datSisMetal.clsSaidaMaterial
  '
  clsSaiMat.ExcluirSAIDAMATERIAL lngSAIDAMATERIALID
  '
  Set clsSaiMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


