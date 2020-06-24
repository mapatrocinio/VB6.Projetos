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
  On Error GoTo trata
  '
  Dim strSql As String
  'Atualizar Estoque
  strSql = "UPDATE PRODUTO SET SALDOESTOQUE = " & Formata_Dados(lngQUANTIDADENOVA, tpDados_Longo, tpNulo_NaoAceita) & _
    " WHERE INSUMOID = " & Formata_Dados(lngPRODUTOID, tpDados_Longo, tpNulo_NaoAceita)
    
  RunSP strSql
  'exclui TAB_ENTRADAMATERIAL
  strSql = "DELETE FROM TAB_ENTRADAMATERIAL WHERE PKID = " & _
    Formata_Dados(lngTAB_ENTRADAMATERIALID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.ExcluirTAB_ENTRADAMATERIAL]", _
            Err.Description
End Sub

Public Function CapturaCodigoEntradaMaterial() As Long
  '
  Dim objRs   As ADODB.Recordset
  Dim lngRet  As Long
  
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "SELECT MAX(CODIGO) AS COD FROM ENTRADAMATERIAL"
  
  Set objRs = RunSPReturnRS(strSql)
  If objRs.EOF Then
    lngRet = 1
  ElseIf Not IsNumeric(objRs.Fields("COD").Value) Then
    lngRet = 1
  Else
    lngRet = CLng(objRs.Fields("COD").Value) + 1
  End If
  '
  objRs.Close
  Set objRs = Nothing
  '
  CapturaCodigoEntradaMaterial = lngRet
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.CapturaCodigoEntradaMaterial]", _
            Err.Description
End Function

Public Function CapturaEntradaMaterial(ByVal strCodigo As String) As ADODB.Recordset
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "SELECT * FROM ENTRADAMATERIAL WHERE CODIGO = " & Formata_Dados(strCodigo, tpDados_Longo, tpNulo_NaoAceita)
  Set CapturaEntradaMaterial = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.CapturaEntradaMaterial]", _
            Err.Description
End Function

Public Sub InserirEntradaMaterial(strData As String, _
                                  strDescricao As String, _
                                  strCodigo As String, _
                                  strUsuario As String, _
                                  lngDOCUMENTOENTRADAID As Long, _
                                  lngFORNECEDORID As Long, _
                                  strDataAquisicao As String, _
                                  strNumero As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO ENTRADAMATERIAL (DATA, DESCRICAO, CODIGO, USUARIO, DOCUMENTOENTRADAID, FORNECEDORID, DATAAQUISICAO, NUMERO) VALUES (" & _
    Formata_Dados(strData, tpDados_DataHora, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strCodigo, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strUsuario, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngDOCUMENTOENTRADAID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngFORNECEDORID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strDataAquisicao, tpDados_DataHora, tpNulo_Aceita) & ", " & _
    Formata_Dados(strNumero, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.InserirEntradaMaterial]", _
            Err.Description
End Sub

Public Function ListarDocumentoEntradaPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT DOCUMENTOENTRADA.* " & _
    "FROM DOCUMENTOENTRADA " & _
    "WHERE DOCUMENTOENTRADA.NOME = " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarDocumentoEntradaPelaDesc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.ListarDocumentoEntradaPelaDesc]", _
            Err.Description
End Function

Public Sub InserirDocumentoEntrada(strNome As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO DOCUMENTOENTRADA (NOME) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.InserirDocumentoEntrada]", _
            Err.Description
End Sub
                         
Public Sub AlterarProdutoPelaEntradaMaterial(ByVal lngPRODUTOID As Long, _
                                             ByVal lngQUANTIDADE As Long, _
                                             ByVal strQuantidade As String)
  '
  On Error GoTo trata
  Dim lngQUANTIDADENOVA As Long
  '
  lngQUANTIDADENOVA = lngQUANTIDADE + CLng(strQuantidade)
  
  Dim strSql  As String
  '
  strSql = "UPDATE PRODUTO SET SALDOESTOQUE = " & Formata_Dados(lngQUANTIDADENOVA, tpDados_Longo, tpNulo_NaoAceita) & " " & _
    "WHERE INSUMOID = " & Formata_Dados(lngPRODUTOID, tpDados_Longo)
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.AlterarProdutoPelaEntradaMaterial]", _
            Err.Description
End Sub
                         
Public Sub InserirTAB_ENTRADAMATERIAL(ByVal lngENTRADAMATERIALID As Long, _
                                      ByVal lngPRODUTOID As Long, _
                                      ByVal strQuantidade As String, _
                                      ByVal strDataValidade As String)
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "INSERT INTO TAB_ENTRADAMATERIAL (ENTRADAMATERIALID, PRODUTOID, QUANTIDADE, DATAVALIDADE) VALUES (" & _
    Formata_Dados(lngENTRADAMATERIALID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngPRODUTOID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strQuantidade, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strDataValidade, tpDados_DataHora) & ");"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.InserirTAB_ENTRADAMATERIAL]", _
            Err.Description
End Sub
                         
Public Sub AlterarEntradaMaterial(lngENTRADAMATERIALID As Long, _
                                  strDescricao As String, _
                                  lngDOCUMENTOENTRADAID As Long, _
                                  lngFORNECEDORID As Long, _
                                  strNumero As String, _
                                  strDataAquisicao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE ENTRADAMATERIAL SET DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "DOCUMENTOENTRADAID = " & Formata_Dados(lngDOCUMENTOENTRADAID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    "FORNECEDORID = " & Formata_Dados(lngFORNECEDORID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    "DATAAQUISICAO = " & Formata_Dados(strDataAquisicao, tpDados_DataHora) & ", " & _
    "NUMERO = " & Formata_Dados(strNumero, tpDados_Texto, tpNulo_Aceita) & " " & _
    "WHERE PKID = " & lngENTRADAMATERIALID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.AlterarEntradaMaterial]", _
            Err.Description
End Sub
Public Sub AlterarDocumentoEntrada(lngDOCUMENTOENTRADAID As Long, _
                                   strNome As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE DOCUMENTOENTRADA SET NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngDOCUMENTOENTRADAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.AlterarDocumentoEntrada]", _
            Err.Description
End Sub

Public Function ListarEntradaMaterial(lngENTRADAMATERIALID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ENTRADAMATERIAL.*, DOCUMENTOENTRADA.NOME AS DOCENTRADA, LOJA.NOME AS NOME_FORNECEDOR " & _
    "FROM ENTRADAMATERIAL LEFT JOIN DOCUMENTOENTRADA ON DOCUMENTOENTRADA.PKID = ENTRADAMATERIAL.DOCUMENTOENTRADAID " & _
    " LEFT JOIN LOJA ON LOJA.PKID = ENTRADAMATERIAL.FORNECEDORID " & _
    "WHERE ENTRADAMATERIAL.PKID = " & _
    Formata_Dados(lngENTRADAMATERIALID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarEntradaMaterial = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.ListarEntradaMaterial]", _
            Err.Description
End Function

Public Function ListarDocumentoEntrada(lngDOCUMENTOENTRADAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT DOCUMENTOENTRADA.* " & _
    "FROM DOCUMENTOENTRADA " & _
    "WHERE DOCUMENTOENTRADA.PKID = " & _
    Formata_Dados(lngDOCUMENTOENTRADAID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarDocumentoEntrada = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.ListarDocumentoEntrada]", _
            Err.Description
End Function

Public Sub ExcluirDOCUMENTOENTRADA(lngDOCUMENTOENTRADAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui DOCUMENTO DE ENTRADA
  strSql = "DELETE FROM DOCUMENTOENTRADA WHERE DOCUMENTOENTRADA.PKID = " & _
    Formata_Dados(lngDOCUMENTOENTRADAID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.ExcluirDOCUMENTOENTRADA]", _
            Err.Description
End Sub

Public Sub ExcluirENTRADAMATERIAL(lngENTRADAMATERIALID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui ENTRADA DE MATERIAL
  strSql = "DELETE FROM ENTRADAMATERIAL WHERE ENTRADAMATERIAL.PKID = " & _
    Formata_Dados(lngENTRADAMATERIALID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.ExcluirENTRADAMATERIAL]", _
            Err.Description
End Sub

