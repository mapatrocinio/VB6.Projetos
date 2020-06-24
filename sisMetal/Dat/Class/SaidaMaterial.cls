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
  On Error GoTo trata
  '
  Dim strSql As String
  'Atualizar Estoque
  strSql = "UPDATE PRODUTO SET SALDOESTOQUE = " & Formata_Dados(lngQUANTIDADENOVA, tpDados_Longo, tpNulo_NaoAceita) & _
    " WHERE INSUMOID = " & Formata_Dados(lngPRODUTOID, tpDados_Longo, tpNulo_NaoAceita)
    
  RunSP strSql
  'exclui TAB_SAIDAMATERIAL
  strSql = "DELETE FROM TAB_SAIDAMATERIAL WHERE PKID = " & _
    Formata_Dados(lngTAB_SAIDAMATERIALID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSaidaMaterial.ExcluirTAB_SAIDAMATERIAL]", _
            Err.Description
End Sub

Public Function CapturaCodigoSaidaMaterial() As Long
  '
  Dim objRs   As ADODB.Recordset
  Dim lngRet  As Long
  
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "SELECT MAX(CODIGO) AS COD FROM SAIDAMATERIAL"
  
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
  CapturaCodigoSaidaMaterial = lngRet
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSaidaMaterial.CapturaCodigoSaidaMaterial]", _
            Err.Description
End Function

Public Function CapturaSaidaMaterial(ByVal strCodigo As String) As ADODB.Recordset
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "SELECT * FROM SAIDAMATERIAL WHERE CODIGO = " & Formata_Dados(strCodigo, tpDados_Longo, tpNulo_NaoAceita)
  Set CapturaSaidaMaterial = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSaidaMaterial.CapturaSaidaMaterial]", _
            Err.Description
End Function

Public Sub InserirSaidaMaterial(strData As String, _
                                strDescricao As String, _
                                strCodigo As String, _
                                strUsuario As String, _
                                lngDOCUMENTOSAIDAID As Long, _
                                lngFILIALID As Long, _
                                strDataAquisicao As String, _
                                strDataTransacao As String, _
                                strMotivo As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO SAIDAMATERIAL (DATA, DESCRICAO, CODIGO, USUARIO, DOCUMENTOSAIDAID, FILIALID, DATAAQUISICAO, DATA_TRANSACAO, MOTIVO) VALUES (" & _
    Formata_Dados(strData, tpDados_DataHora, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strCodigo, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strUsuario, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngDOCUMENTOSAIDAID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(IIf(lngFILIALID = 0, "", lngFILIALID), tpDados_Longo) & ", " & _
    Formata_Dados(strDataAquisicao, tpDados_DataHora, tpNulo_Aceita) & ", " & _
    Formata_Dados(strDataTransacao, tpDados_DataHora, tpNulo_Aceita) & ", " & _
    Formata_Dados(strMotivo, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSaidaMaterial.InserirSaidaMaterial]", _
            Err.Description
End Sub

Public Function ListarDocumentoSaidaPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT DOCUMENTOSAIDA.* " & _
    "FROM DOCUMENTOSAIDA " & _
    "WHERE DOCUMENTOSAIDA.NOME = " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarDocumentoSaidaPelaDesc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSaidaMaterial.ListarDocumentoSaidaPelaDesc]", _
            Err.Description
End Function

Public Sub InserirDocumentoSaida(strNome As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO DOCUMENTOSAIDA (NOME) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSaidaMaterial.InserirDocumentoSaida]", _
            Err.Description
End Sub
                         
Public Sub AlterarProdutoPelaSaidaMaterial(ByVal lngPRODUTOID As Long, _
                                           ByVal lngQUANTIDADE As Long, _
                                           ByVal strQuantidade As String, _
                                           ByVal strDocumento As String)
  '
  On Error GoTo trata
  Dim lngQUANTIDADENOVA As Long
  '
  If strDocumento = "TRANSFERÊNCIA DE ENTRADA" Then
    lngQUANTIDADENOVA = lngQUANTIDADE + CLng(strQuantidade)
  Else
    lngQUANTIDADENOVA = lngQUANTIDADE - CLng(strQuantidade)
  End If
  Dim strSql  As String
  '
  strSql = "UPDATE PRODUTO SET SALDOESTOQUE = " & Formata_Dados(lngQUANTIDADENOVA, tpDados_Longo, tpNulo_NaoAceita) & " " & _
    "WHERE INSUMOID = " & Formata_Dados(lngPRODUTOID, tpDados_Longo)
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSaidaMaterial.AlterarProdutoPelaSaidaMaterial]", _
            Err.Description
End Sub
                         
Public Sub InserirTAB_SAIDAMATERIAL(ByVal lngSAIDAMATERIALID As Long, _
                                    ByVal lngPRODUTOID As Long, _
                                    ByVal strQuantidade As String)
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "INSERT INTO TAB_SAIDAMATERIAL (SAIDAMATERIALID, PRODUTOID, QUANTIDADE) VALUES (" & _
    Formata_Dados(lngSAIDAMATERIALID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngPRODUTOID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strQuantidade, tpDados_Longo, tpNulo_NaoAceita) & ");"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSaidaMaterial.InserirTAB_SAIDAMATERIAL]", _
            Err.Description
End Sub
                         
Public Sub AlterarSaidaMaterial(lngSAIDAMATERIALID As Long, _
                                strDescricao As String, _
                                lngDOCUMENTOSAIDAID As Long, _
                                lngFILIALID As Long, _
                                strDataAquisicao As String, _
                                strDataTransacao As String, _
                                strMotivo As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE SAIDAMATERIAL SET DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "DOCUMENTOSAIDAID = " & Formata_Dados(lngDOCUMENTOSAIDAID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    "FILIALID = " & Formata_Dados(lngFILIALID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    "DATAAQUISICAO = " & Formata_Dados(strDataAquisicao, tpDados_DataHora) & ", " & _
    "DATA_TRANSACAO = " & Formata_Dados(strDataTransacao, tpDados_DataHora) & ", " & _
    "MOTIVO = " & Formata_Dados(strMotivo, tpDados_Texto) & " " & _
    "WHERE PKID = " & lngSAIDAMATERIALID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSaidaMaterial.AlterarSaidaMaterial]", _
            Err.Description
End Sub
Public Sub AlterarDocumentoSaida(lngDOCUMENTOSAIDAID As Long, _
                                 strNome As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE DOCUMENTOSAIDA SET NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngDOCUMENTOSAIDAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSaidaMaterial.AlterarDocumentoSaida]", _
            Err.Description
End Sub

Public Function ListarSaidaMaterial(lngSAIDAMATERIALID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT SAIDAMATERIAL.*, DOCUMENTOSAIDA.NOME AS DOCSAIDA, LOJA.NOME AS NOME_FILIAL " & _
    "FROM SAIDAMATERIAL LEFT JOIN DOCUMENTOSAIDA ON DOCUMENTOSAIDA.PKID = SAIDAMATERIAL.DOCUMENTOSAIDAID " & _
    " LEFT JOIN LOJA ON LOJA.PKID = SAIDAMATERIAL.FILIALID " & _
    "WHERE SAIDAMATERIAL.PKID = " & _
    Formata_Dados(lngSAIDAMATERIALID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarSaidaMaterial = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSaidaMaterial.ListarSaidaMaterial]", _
            Err.Description
End Function

Public Function ListarDocumentoSaida(lngDOCUMENTOSAIDAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT DOCUMENTOSAIDA.* " & _
    "FROM DOCUMENTOSAIDA " & _
    "WHERE DOCUMENTOSAIDA.PKID = " & _
    Formata_Dados(lngDOCUMENTOSAIDAID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarDocumentoSaida = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSaidaMaterial.ListarDocumentoSaida]", _
            Err.Description
End Function

Public Sub ExcluirDOCUMENTOSAIDA(lngDOCUMENTOSAIDAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui DOCUMENTO DE SAIDA
  strSql = "DELETE FROM DOCUMENTOSAIDA WHERE DOCUMENTOSAIDA.PKID = " & _
    Formata_Dados(lngDOCUMENTOSAIDAID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSaidaMaterial.ExcluirDOCUMENTOSAIDA]", _
            Err.Description
End Sub

Public Sub ExcluirSAIDAMATERIAL(lngSAIDAMATERIALID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui SAIDA DE MATERIAL
  strSql = "DELETE FROM SAIDAMATERIAL WHERE SAIDAMATERIAL.PKID = " & _
    Formata_Dados(lngSAIDAMATERIALID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSaidaMaterial.ExcluirSAIDAMATERIAL]", _
            Err.Description
End Sub

