VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Public Sub InserirEmpresa(strNome As String, _
                          strTel As String, _
                          strTel1 As String, _
                          strTel2 As String, _
                          strContato As String, _
                          strObs As String, _
                          strTipoEmpresaId As String, _
                          strPercentualAg As String, _
                          strEndRua As String, _
                          strEndNumero As String, _
                          strEndComplemento As String, _
                          strEndCep As String, _
                          strEndBairro As String, _
                          strEndCidade As String, _
                          strEndEstado As String, _
                          strEndPais As String, _
                          strCobRua As String, _
                          strCobNumero As String, _
                          strCobComplemento As String, _
                          strCobCep As String, _
                          strCobBairro As String, _
                          strCobCidade As String, _
                          strCobEstado As String, strCobPais As String, _
                          strCGC As String, strInscrEstadual As String, strInscrMunicipal As String, strPercDescDiaria As String, strCredito As String, _
                          strNomeFantasia As String, strNroContrato As String, strDtInicioContrato As String, strDtFimContrato As String, blnEndCobIgualEndCorr As Boolean, lngParceiroId As Long)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO EMPRESA (NOME, TEL, TEL1, TEL2, CONTATO, OBSERVACAO, TIPOEMPRESAID, PERCENTUALAG, ENDRUA, ENDNUMERO, ENDCOMPLEMENTO, ENDCEP, ENDBAIRRO, ENDCIDADE, ENDESTADO, ENDPAIS, COBRUA, COBNUMERO, COBCOMPLEMENTO, COBCEP, COBBAIRRO, COBCIDADE, COBESTADO, COBPAIS, CGC, INSCRESTADUAL, INSCRMUNICIPAL, PERCDESCDIARIA, CREDITO, NOMEFANTASIA, NROCONTRATO, DTINICIOCONTRATO, DTFIMCONTRATO, ENDCOBIGUALENDCORR, PARCEIROID) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strTel, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strTel1, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strTel2, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strContato, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strObs, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strTipoEmpresaId, tpDados_Longo, tpNulo_Aceita) & ", " & _
    Formata_Dados(strPercentualAg, tpDados_Longo, tpNulo_Aceita) & ", " & _
    Formata_Dados(strEndRua, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strEndNumero, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strEndComplemento, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strEndCep, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strEndBairro, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strEndCidade, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strEndEstado, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strEndPais, tpDados_Texto, tpNulo_Aceita) & ", "
  strSql = strSql & _
    Formata_Dados(strCobRua, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strCobNumero, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strCobComplemento, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strCobCep, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strCobBairro, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strCobCidade, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strCobEstado, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strCobPais, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strCGC, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strInscrEstadual, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strInscrMunicipal, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strPercDescDiaria, tpDados_Longo, tpNulo_Aceita) & ", " & _
    Formata_Dados(strCredito, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strNomeFantasia, tpDados_Texto) & ", " & _
    Formata_Dados(strNroContrato, tpDados_Texto) & ", " & _
    Formata_Dados(strDtInicioContrato, tpDados_DataHora) & ", " & _
    Formata_Dados(strDtFimContrato, tpDados_DataHora) & ", " & _
    Formata_Dados(blnEndCobIgualEndCorr, tpDados_Boolean) & ", " & _
    Formata_Dados(lngParceiroId, tpDados_Longo) & " " & _
    ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresa.InserirEmpresa]", _
            Err.Description
End Sub

Public Sub AlterarEmpresa(lngEMPRESAID As Long, strNome As String, _
                          strTel As String, _
                          strTel1 As String, _
                          strTel2 As String, _
                          strContato As String, _
                          strObs As String, _
                          strTipoEmpresaId As String, _
                          strPercentualAg As String, _
                          strEndRua As String, _
                          strEndNumero As String, _
                          strEndComplemento As String, _
                          strEndCep As String, _
                          strEndBairro As String, _
                          strEndCidade As String, _
                          strEndEstado As String, _
                          strEndPais As String, _
                          strCobRua As String, _
                          strCobNumero As String, _
                          strCobComplemento As String, _
                          strCobCep As String, _
                          strCobBairro As String, _
                          strCobCidade As String, _
                          strCobEstado As String, strCobPais As String, _
                          strCGC As String, strInscrEstadual As String, strInscrMunicipal As String, strPercDescDiaria As String, strCredito As String, _
                          strNomeFantasia As String, strNroContrato As String, strDtInicioContrato As String, strDtFimContrato As String, blnEndCobIgualEndCorr As Boolean)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE EMPRESA SET NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "TEL = " & Formata_Dados(strTel, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "TEL1 = " & Formata_Dados(strTel1, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "TEL2 = " & Formata_Dados(strTel2, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "CONTATO = " & Formata_Dados(strContato, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "OBSERVACAO = " & Formata_Dados(strObs, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "TIPOEMPRESAID = " & Formata_Dados(strTipoEmpresaId, tpDados_Longo, tpNulo_Aceita) & ", " & _
    "PERCENTUALAG = " & Formata_Dados(strPercentualAg, tpDados_Longo, tpNulo_Aceita) & ", " & _
    "ENDRUA = " & Formata_Dados(strEndRua, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "ENDNUMERO = " & Formata_Dados(strEndNumero, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "ENDCOMPLEMENTO = " & Formata_Dados(strEndComplemento, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "ENDCEP = " & Formata_Dados(strEndCep, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "ENDBAIRRO = " & Formata_Dados(strEndBairro, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "ENDCIDADE = " & Formata_Dados(strEndCidade, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "ENDESTADO = " & Formata_Dados(strEndEstado, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "ENDPAIS = " & Formata_Dados(strEndPais, tpDados_Texto, tpNulo_Aceita) & ", "
  strSql = strSql & _
    "COBRUA = " & Formata_Dados(strCobRua, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "COBNUMERO = " & Formata_Dados(strCobNumero, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "COBCOMPLEMENTO = " & Formata_Dados(strCobComplemento, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "COBCEP = " & Formata_Dados(strCobCep, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "COBBAIRRO = " & Formata_Dados(strCobBairro, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "COBCIDADE = " & Formata_Dados(strCobCidade, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "COBESTADO = " & Formata_Dados(strCobEstado, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "COBPAIS = " & Formata_Dados(strCobPais, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "CGC = " & Formata_Dados(strCGC, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "INSCRESTADUAL = " & Formata_Dados(strInscrEstadual, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "INSCRMUNICIPAL = " & Formata_Dados(strInscrMunicipal, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "PERCDESCDIARIA = " & Formata_Dados(strPercDescDiaria, tpDados_Moeda, tpNulo_Aceita) & ", " & _
    "CREDITO = " & Formata_Dados(strCredito, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "NOMEFANTASIA = " & Formata_Dados(strNomeFantasia, tpDados_Texto) & ", " & _
    "NROCONTRATO = " & Formata_Dados(strNroContrato, tpDados_Texto) & ", " & _
    "DTINICIOCONTRATO = " & Formata_Dados(strDtInicioContrato, tpDados_DataHora) & ", " & _
    "DTFIMCONTRATO = " & Formata_Dados(strDtFimContrato, tpDados_DataHora) & ", " & _
    "ENDCOBIGUALENDCORR = " & Formata_Dados(blnEndCobIgualEndCorr, tpDados_Boolean) & " " & _
    "WHERE PKID = " & lngEMPRESAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresa.AlterarEmpresa]", _
            Err.Description
End Sub

Public Function ListarEmpresa(lngEMPRESAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT EMPRESA.*, TIPOEMPRESA.DESCRICAO AS DESCRTIPOEMPRESA " & _
    "FROM EMPRESA LEFT JOIN TIPOEMPRESA ON TIPOEMPRESA.PKID = EMPRESA.TIPOEMPRESAID " & _
    "WHERE EMPRESA.PKID = " & _
    Formata_Dados(lngEMPRESAID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarEmpresa = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresa.ListarEmpresa]", _
            Err.Description
End Function
Public Function ListarEmpresaPeloNome(strNome As String, _
                                      lngEMPRESAID As Long, _
                                      lngParceiroId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT EMPRESA.* " & _
    "FROM EMPRESA " & _
    "WHERE EMPRESA.NOME = " & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "AND EMPRESA.PKID <> " & _
    Formata_Dados(lngEMPRESAID, tpDados_Longo, tpNulo_NaoAceita) & " " & _
    "AND EMPRESA.PARCEIROID <> " & _
    Formata_Dados(lngParceiroId, tpDados_Longo)
  Set ListarEmpresaPeloNome = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresa.ListarEmpresaPeloNome]", _
            Err.Description
End Function

Public Sub ExlcuirEmpresa(lngEMPRESAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM EMPRESA WHERE PKID = " & _
    Formata_Dados(lngEMPRESAID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresa.ExlcuirEmpresa]", _
            Err.Description
End Sub



