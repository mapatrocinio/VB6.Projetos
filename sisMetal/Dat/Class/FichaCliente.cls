VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFichaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirFichaCliente(lngFichaClienteId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM FICHACLIENTE " & _
    "WHERE FICHACLIENTE.PKID = " & _
    Formata_Dados(lngFichaClienteId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFichaCliente.ExcluirFichaCliente]", _
            Err.Description
End Function

Public Function SelecionarFichaClientePeloPkid(lngFichaClienteId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT FICHACLIENTE.*, TIPODOCUMENTO.DESCRICAO AS DESC_TIPODOCUMENTO " & _
    "FROM FICHACLIENTE LEFT JOIN TIPODOCUMENTO ON TIPODOCUMENTO.PKID = FICHACLIENTE.TIPODOCUMENTOID " & _
    "WHERE FICHACLIENTE.PKID = " & _
    Formata_Dados(lngFichaClienteId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarFichaClientePeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFichaCliente.SelecionarFichaClientePeloPkid]", _
            Err.Description
End Function


Public Sub AlterarFichaCliente(lngPKID As Long, _
                               lngTIPODOCUMENTOID As Long, _
                               strSobreNome As String, _
                               strNome As String, _
                               strEndereco As String, _
                               strNumero As String, _
                               strComplemento As String, _
                               strBairro As String, _
                               strCidade As String, _
                               strEstado As String, _
                               strCep As String, _
                               strPais As String, _
                               strTel1 As String, _
                               strTel2 As String, _
                               strTel3 As String, _
                               strDtNascimento As String, _
                               strSexo As String, _
                               strNroDocumento As String, _
                               strPercDesc As String, _
                               strEmail As String, _
                               strObservacao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE FICHACLIENTE SET " & _
    "TIPODOCUMENTOID = " & Formata_Dados(lngTIPODOCUMENTOID, tpDados_Longo, tpNulo_Aceita) & ", " & _
    "SOBRENOME = " & Formata_Dados(strSobreNome, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "ENDERECO = " & Formata_Dados(strEndereco, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "NUMERO = " & Formata_Dados(strNumero, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "COMPLEMENTO = " & Formata_Dados(strComplemento, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "BAIRRO = " & Formata_Dados(strBairro, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "CIDADE = " & Formata_Dados(strCidade, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "ESTADO = " & Formata_Dados(strEstado, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "CEP = " & Formata_Dados(strCep, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "PAIS = " & Formata_Dados(strPais, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "TEL1 = " & Formata_Dados(strTel1, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "TEL2 = " & Formata_Dados(strTel2, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "TEL3 = " & Formata_Dados(strTel3, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "DTNASCIMENTO = " & Formata_Dados(strDtNascimento, tpDados_DataHora, tpNulo_Aceita) & ", " & _
    "SEXO = " & Formata_Dados(strSexo, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "NRODOCUMENTO = " & Formata_Dados(strNroDocumento, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "PERCDESC = " & Formata_Dados(strPercDesc, tpDados_Moeda) & ", " & _
    "EMAIL = " & Formata_Dados(strEmail, tpDados_Texto) & ", " & _
    "OBSERVACAO = " & Formata_Dados(strObservacao, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFichaCliente.AlterarFichaCliente]", _
            Err.Description
End Sub

Public Function InserirFichaCliente(lngTIPODOCUMENTOID As Long, _
                                    strSobreNome As String, _
                                    strNome As String, _
                                    strEndereco As String, _
                                    strNumero As String, _
                                    strComplemento As String, _
                                    strBairro As String, _
                                    strCidade As String, _
                                    strEstado As String, _
                                    strCep As String, _
                                    strPais As String, _
                                    strTel1 As String, _
                                    strTel2 As String, _
                                    strTel3 As String, _
                                    strDtNascimento As String, _
                                    strSexo As String, _
                                    strNroDocumento As String, _
                                    strPercDesc As String, _
                                    strEmail As String, _
                                    strObservacao As String) As Long

  On Error GoTo trata
  '
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  '
  InserirFichaCliente = 0
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & " INSERT INTO FICHACLIENTE (TIPODOCUMENTOID, SOBRENOME, NOME, ENDERECO, NUMERO, COMPLEMENTO, BAIRRO, CIDADE, ESTADO, CEP, PAIS, TEL1, TEL2, TEL3, DTNASCIMENTO, SEXO, NRODOCUMENTO, PERCDESC, EMAIL, OBSERVACAO) VALUES (" & _
    Formata_Dados(lngTIPODOCUMENTOID, tpDados_Longo, tpNulo_Aceita) & ", " & _
    Formata_Dados(strSobreNome, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strEndereco, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strNumero, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strComplemento, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strBairro, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strCidade, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strEstado, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strCep, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strPais, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strTel1, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strTel2, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strTel3, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strDtNascimento, tpDados_DataHora, tpNulo_Aceita) & ", " & _
    Formata_Dados(strSexo, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strNroDocumento, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strPercDesc, tpDados_Moeda) & ", " & _
    Formata_Dados(strEmail, tpDados_Texto) & ", " & _
    Formata_Dados(strObservacao, tpDados_Texto) & ")"
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    InserirFichaCliente = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFichaCliente.InserirFichaCliente]", _
            Err.Description
End Function


Public Function SelecionarFichaClientePeloSobreNome(strSobreNome As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT FICHACLIENTE.*, TIPODOCUMENTO.DESCRICAO, " & _
    "QTD_LOC = ISNULL((SELECT COUNT(*) FROM TAB_FICHACLIELOC WHERE FICHACLIENTEID = FICHACLIENTE.PKID), 0) " & _
    "FROM FICHACLIENTE LEFT JOIN TIPODOCUMENTO ON TIPODOCUMENTO.PKID = FICHACLIENTE.TIPODOCUMENTOID " & _
    "WHERE FICHACLIENTE.SOBRENOME = " & _
    Formata_Dados(strSobreNome, tpDados_Texto, tpNulo_Aceita)
  Set SelecionarFichaClientePeloSobreNome = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFichaCliente.SelecionarFichaClientePeloSobreNome]", _
            Err.Description
End Function

Public Function SelecionarFichaCliente(lngTIPODOCUMENTOID As Long, _
                                       strNroDocumento As String, _
                                       lngFichaClienteId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT FICHACLIENTE.*, TIPODOCUMENTO.DESCRICAO " & _
    "FROM FICHACLIENTE LEFT JOIN TIPODOCUMENTO ON TIPODOCUMENTO.PKID = FICHACLIENTE.TIPODOCUMENTOID " & _
    "WHERE FICHACLIENTE.TIPODOCUMENTOID = " & _
      Formata_Dados(lngTIPODOCUMENTOID, tpDados_Longo, tpNulo_Aceita) & _
    " AND FICHACLIENTE.NRODOCUMENTO = " & _
      Formata_Dados(strNroDocumento, tpDados_Texto, tpNulo_Aceita) & _
    " AND FICHACLIENTE.PKID <> " & Formata_Dados(lngFichaClienteId, tpDados_Longo)
  Set SelecionarFichaCliente = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFichaCliente.SelecionarFichaCliente]", _
            Err.Description
End Function
Public Function SelecionarFichaClientePeloNome(strNome As String, _
                                              strSobreNome As String, _
                                              strDataNascimento As String, _
                                              lngFichaClienteId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT FICHACLIENTE.*, TIPODOCUMENTO.DESCRICAO AS DESC_TIPODOCUMENTO " & _
    "FROM FICHACLIENTE LEFT JOIN TIPODOCUMENTO ON TIPODOCUMENTO.PKID = FICHACLIENTE.TIPODOCUMENTOID " & _
    "WHERE FICHACLIENTE.NOME = " & _
      Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & _
      " AND FICHACLIENTE.SOBRENOME = " & _
        Formata_Dados(strSobreNome, tpDados_Texto, tpNulo_Aceita) & _
      " AND FICHACLIENTE.DTNASCIMENTO = " & _
        Formata_Dados(strDataNascimento, tpDados_DataHora, tpNulo_Aceita) & _
      " AND FICHACLIENTE.PKID <> " & Formata_Dados(lngFichaClienteId, tpDados_Longo)
      

  Set SelecionarFichaClientePeloNome = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFichaCliente.SelecionarFichaClientePeloNome]", _
            Err.Description
End Function

