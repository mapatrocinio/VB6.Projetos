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


Public Sub ExcluirITEMENTRADA(ByVal lngITEMENTRADAID As Long, _
                                      ByVal lngESTOQUEID As Long, _
                                      ByVal lngQUANTIDADENOVA As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'Atualizar Estoque
  strSql = "UPDATE ESTOQUE SET QUANTIDADE = ISNULL(QUANTIDADE,0) - " & Formata_Dados(lngQUANTIDADENOVA, tpDados_Longo, tpNulo_NaoAceita) & _
    " WHERE PKID = " & Formata_Dados(lngESTOQUEID, tpDados_Longo, tpNulo_NaoAceita)
    
  RunSP strSql
  'exclui ITEMENTRADA
  strSql = "DELETE FROM ITEMENTRADA WHERE PKID = " & _
    Formata_Dados(lngITEMENTRADAID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.ExcluirITEMENTRADA]", _
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
                                  lngDOCUMENTOID As Long, _
                                  strFornecedor As String, _
                                  strDataAquisicao As String, _
                                  strNumero As String, _
                                  ByRef lngENTRADAMATERIALID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO ENTRADAMATERIAL (DATA, DESCRICAO, CODIGO, USUARIO, DOCUMENTOID, FORNECEDOR, DATAAQUISICAO, NUMERO) VALUES (" & _
    Formata_Dados(strData, tpDados_DataHora, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strCodigo, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strUsuario, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngDOCUMENTOID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strFornecedor, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strDataAquisicao, tpDados_DataHora, tpNulo_Aceita) & ", " & _
    Formata_Dados(strNumero, tpDados_Texto, tpNulo_Aceita) & ");"
    
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngENTRADAMATERIALID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.InserirEntradaMaterial]", _
            Err.Description
End Sub


Public Sub InserirDocumentoEntrada(strNome As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO DOCUMENTO (NOME) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.InserirDocumentoEntrada]", _
            Err.Description
End Sub
                         
Public Sub AlterarEstoquePelaEntradaMaterial(ByVal lngESTOQUEID As Long, _
                                             ByVal strCodigo As String, _
                                             ByVal strQuantidade As String)
  '
  On Error GoTo trata
  '
  Dim strSql  As String
  '
  strSql = "UPDATE ESTOQUE SET QUANTIDADE = ISNULL(QUANTIDADE,0) + " & Formata_Dados(strQuantidade, tpDados_Longo) & " " & _
    "WHERE PKID = " & lngESTOQUEID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.AlterarEstoquePelaEntradaMaterial]", _
            Err.Description
End Sub
                         
Public Sub InserirITEMENTRADA(ByVal lngENTRADAMATERIALID As Long, _
                                      ByVal lngESTOQUEID As Long, _
                                      ByVal strQuantidade As String)
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "INSERT INTO ITEMENTRADA (ENTRADAMATERIALID, ESTOQUEID, QUANTIDADE) VALUES (" & _
    Formata_Dados(lngENTRADAMATERIALID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngESTOQUEID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strQuantidade, tpDados_Longo, tpNulo_NaoAceita) & ");"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntradaMaterial.InserirITEMENTRADA]", _
            Err.Description
End Sub
                         
Public Sub AlterarEntradaMaterial(lngENTRADAMATERIALID As Long, _
                                  strDescricao As String, _
                                  lngDOCUMENTOID As Long, _
                                  strFornecedor As String, _
                                  strNumero As String, _
                                  strDataAquisicao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE ENTRADAMATERIAL SET DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "DOCUMENTOID = " & Formata_Dados(lngDOCUMENTOID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    "FORNECEDOR = " & Formata_Dados(strFornecedor, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
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

Public Function ListarEntradaMaterial(lngENTRADAMATERIALID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ENTRADAMATERIAL.*, DOCUMENTO.NOME AS DOCENTRADA " & _
    "FROM ENTRADAMATERIAL LEFT JOIN DOCUMENTO ON DOCUMENTO.PKID = ENTRADAMATERIAL.DOCUMENTOID " & _
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

