VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAssociado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirAssociado(lngASSOCIADOID As Long, _
                                 strIcAssociado As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  'CONVENIOS ASSOCIADOS
  strSql = "DELETE FROM TAB_CONVASSOC " & _
    "WHERE TAB_CONVASSOC.ASSOCIADOID = " & _
    Formata_Dados(lngASSOCIADOID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  'MENSALIDADES DE CONVÊNIO
  'strSql = "DELETE FROM MENSALIDADECONVENIO " & _
    "WHERE MENSALIDADECONVENIO.ASSOCIADOID = " & _
    Formata_Dados(lngASSOCIADOID, tpDados_Longo, tpNulo_Aceita)
  'RunSP strSql
  'MENSALIDADES PLANO
  'strSql = "DELETE FROM MENSALIDADEPLANO " & _
    "WHERE MENSALIDADEPLANO.ASSOCIADOID = " & _
    Formata_Dados(lngASSOCIADOID, tpDados_Longo, tpNulo_Aceita)
  'RunSP strSql
  'MENSALIDADES TAB_PROFASSOC
  strSql = "DELETE FROM TAB_PROFASSOC " & _
    "WHERE TAB_PROFASSOC.ASSOCIADOID = " & _
    Formata_Dados(lngASSOCIADOID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  
  If strIcAssociado = "D" Then
    strSql = "DELETE FROM DEPENDENTE " & _
      "WHERE DEPENDENTE.ASSOCIADOID = " & _
      Formata_Dados(lngASSOCIADOID, tpDados_Longo, tpNulo_Aceita)
    RunSP strSql
  ElseIf strIcAssociado = "T" Then
    '-------------------------------
    'DELETAR MENSALIDADE DE PAGAMENTO
    'AINDA NÃO IMPLEMENTADO
    '-------------------------------
    
    'DEPENDENTES
    Set objRs = SelecionarDependentePeloTitular(lngASSOCIADOID)
    Do While Not objRs.EOF
      ExcluirAssociado objRs.Fields("ASSOCIADOID").Value, "D"
      objRs.MoveNext
    Loop
    objRs.Close
    Set objRs = Nothing
    '
    strSql = "DELETE FROM TITULAR " & _
      "WHERE TITULAR.ASSOCIADOID = " & _
      Formata_Dados(lngASSOCIADOID, tpDados_Longo, tpNulo_Aceita)
    RunSP strSql
  End If
  strSql = "DELETE FROM ASSOCIADO " & _
    "WHERE ASSOCIADO.PKID = " & _
    Formata_Dados(lngASSOCIADOID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAssociado.ExcluirAssociado]", _
            Err.Description
End Function

Public Function SelecionarAssociadoPeloPkid(lngASSOCIADOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ASSOCIADO.*, TIPOSOCIO.DESCRICAO AS DESCR_TIPOSOCIO, VALORPLANO.DESCRICAO AS DESCR_VALORPLANO, ESTADOCIVIL.DESCRICAO AS DESCR_ESTADOCIVIL " & _
    "FROM ASSOCIADO LEFT JOIN TIPOSOCIO ON ASSOCIADO.TIPOSOCIOID = TIPOSOCIO.PKID " & _
    "LEFT JOIN VALORPLANO ON ASSOCIADO.VALORPLANOID = VALORPLANO.PKID " & _
    "LEFT JOIN ESTADOCIVIL ON ASSOCIADO.ESTADOCIVILID = ESTADOCIVIL.PKID " & _
    "WHERE ASSOCIADO.PKID = " & _
    Formata_Dados(lngASSOCIADOID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarAssociadoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAssociado.SelecionarAssociadoPeloPkid]", _
            Err.Description
End Function

Public Function SelecionarDependentePeloTitular(lngTitularId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT DEPENDENTE.* " & _
    "FROM DEPENDENTE " & _
    "WHERE DEPENDENTE.TITULARASSOCIADOID = " & _
    Formata_Dados(lngTitularId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarDependentePeloTitular = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAssociado.SelecionarAssociadoPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarAssociado(lngPKID As Long, _
                           lngTIPOSOCIOID As Long, _
                           lngValorPlanoId As Long, _
                           lngESTADOCIVILID As Long, _
                           strIdentidade As String, _
                           strOrgaoEmissor As String, _
                           strNaturalidade As String, _
                           strCpf As String, _
                           strNome As String, _
                           strSexo As String, _
                           strDtNascimento As String, _
                           strNomeMae As String, _
                           strEspecial As String, _
                           strEndRuaRes As String, _
                           strEndNumeroRes As String, _
                           strEndComplRes As String, _
                           strEndEstadoRes As String, _
                           strEndCepRes As String, _
                           strEndBairroRes As String, _
                           strEndCidadeRes As String, _
                           strTelefoneRes As String, _
                           strCelular As String, _
                           strEmail As String, _
                           strExcluido As String, _
                           strObservacao As String)


  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE ASSOCIADO SET " & _
    "TIPOSOCIOID = " & Formata_Dados(IIf(lngTIPOSOCIOID = 0, "", lngTIPOSOCIOID), tpDados_Longo) & ", " & _
    "VALORPLANOID = " & Formata_Dados(IIf(lngValorPlanoId = 0, "", lngValorPlanoId), tpDados_Longo) & ", " & _
    "ESTADOCIVILID = " & Formata_Dados(IIf(lngESTADOCIVILID = 0, "", lngESTADOCIVILID), tpDados_Longo) & ", " & _
    "IDENTIDADE = " & Formata_Dados(strIdentidade, tpDados_Texto) & ", " & _
    "ORGEMISS = " & Formata_Dados(strOrgaoEmissor, tpDados_Texto) & ", " & _
    "NATURALIDADE = " & Formata_Dados(strNaturalidade, tpDados_Texto) & ", " & _
    "CPF = " & Formata_Dados(strCpf, tpDados_Texto) & ", " & _
    "NOME = " & Formata_Dados(strNome, tpDados_Texto) & ", " & _
    "SEXO = " & Formata_Dados(strSexo, tpDados_Texto) & ", " & _
    "DATANASCIMENTO = " & Formata_Dados(strDtNascimento, tpDados_DataHora) & ", " & _
    "NOMEMAE = " & Formata_Dados(strNomeMae, tpDados_Texto) & ", " & _
    "ESPECIAL = " & Formata_Dados(strEspecial, tpDados_Texto) & ", " & _
    "ENDRUARES = " & Formata_Dados(strEndRuaRes, tpDados_Texto) & ", " & _
    "ENDNUMERORES = " & Formata_Dados(strEndNumeroRes, tpDados_Texto) & ", " & _
    "ENDCOMPLRES = " & Formata_Dados(strEndComplRes, tpDados_Texto) & ", " & _
    "ENDCEPRES = " & Formata_Dados(strEndCepRes, tpDados_Texto) & ", "
  strSql = strSql & "ENDBAIRRORES = " & Formata_Dados(strEndBairroRes, tpDados_Texto) & ", " & _
    "ENDCIDADERES = " & Formata_Dados(strEndCidadeRes, tpDados_Texto) & ", " & _
    "ENDESTADORES = " & Formata_Dados(strEndEstadoRes, tpDados_Texto) & ", " & _
    "TELEFONERES1 = " & Formata_Dados(strTelefoneRes, tpDados_Texto) & ", " & _
    "CELULAR = " & Formata_Dados(strCelular, tpDados_Texto) & ", " & _
    "EMAIL = " & Formata_Dados(strEmail, tpDados_Texto) & ", " & _
    "EXCLUIDO = " & Formata_Dados(strExcluido, tpDados_Texto) & ", " & _
    "OBSERVACAO = " & Formata_Dados(strObservacao, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAssociado.AlterarAssociado]", _
            Err.Description
End Sub

Public Sub InserirAssociado(lngTIPOSOCIOID As Long, _
                           lngValorPlanoId As Long, _
                           lngESTADOCIVILID As Long, _
                           strIdentidade As String, _
                           strOrgaoEmissor As String, _
                           strNaturalidade As String, _
                           strCpf As String, _
                           strNome As String, _
                           strSexo As String, _
                           strDtNascimento As String, _
                           strNomeMae As String, _
                           strEspecial As String, _
                           strEndRuaRes As String, _
                           strEndNumeroRes As String, _
                           strEndComplRes As String, _
                           strEndEstadoRes As String, _
                           strEndCepRes As String, _
                           strEndBairroRes As String, _
                           strEndCidadeRes As String, _
                           strTelefoneRes As String, _
                           strCelular As String, _
                           strEmail As String, _
                           strObservacao As String, _
                           strIcAssociado As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO ASSOCIADO (TIPOSOCIOID, VALORPLANOID, ESTADOCIVILID, IDENTIDADE, ORGEMISS, NATURALIDADE, CPF, NOME, SEXO, DATANASCIMENTO, NOMEMAE, ESPECIAL, ENDRUARES, ENDNUMERORES, ENDCOMPLRES, ENDCEPRES, ENDBAIRRORES, ENDCIDADERES, " & _
    "ENDESTADORES, TELEFONERES1, CELULAR, EMAIL, EXCLUIDO, OBSERVACAO, ICASSOCIADO) VALUES (" & _
    Formata_Dados(IIf(lngTIPOSOCIOID = 0, "", lngTIPOSOCIOID), tpDados_Longo) & ", " & _
    Formata_Dados(IIf(lngValorPlanoId = 0, "", lngValorPlanoId), tpDados_Longo) & ", " & _
    Formata_Dados(IIf(lngESTADOCIVILID = 0, "", lngESTADOCIVILID), tpDados_Longo) & ", " & _
    Formata_Dados(strIdentidade, tpDados_Texto) & ", " & _
    Formata_Dados(strOrgaoEmissor, tpDados_Texto) & ", " & _
    Formata_Dados(strNaturalidade, tpDados_Texto) & ", " & _
    Formata_Dados(strCpf, tpDados_Texto) & ", " & _
    Formata_Dados(strNome, tpDados_Texto) & ", " & _
    Formata_Dados(strSexo, tpDados_Texto) & ", " & _
    Formata_Dados(strDtNascimento, tpDados_DataHora) & ", " & _
    Formata_Dados(strNomeMae, tpDados_Texto) & ", " & _
    Formata_Dados(strEspecial, tpDados_Texto) & ", " & _
    Formata_Dados(strEndRuaRes, tpDados_Texto) & ", " & _
    Formata_Dados(strEndNumeroRes, tpDados_Texto) & ", " & _
    Formata_Dados(strEndComplRes, tpDados_Texto) & ", " & _
    Formata_Dados(strEndCepRes, tpDados_Texto) & ", "
  strSql = strSql & Formata_Dados(strEndBairroRes, tpDados_Texto) & ", " & _
    Formata_Dados(strEndCidadeRes, tpDados_Texto) & ", " & _
    Formata_Dados(strEndEstadoRes, tpDados_Texto) & ", " & _
    Formata_Dados(strTelefoneRes, tpDados_Texto) & ", " & _
    Formata_Dados(strCelular, tpDados_Texto) & ", " & _
    Formata_Dados(strEmail, tpDados_Texto) & ", " & _
    Formata_Dados("N", tpDados_Texto) & ", " & _
    Formata_Dados(strObservacao, tpDados_Texto) & ", " & _
    Formata_Dados(strIcAssociado, tpDados_Texto) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAssociado.InserirAssociado]", _
            Err.Description
End Sub

Public Function SelecionarAssociado(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ASSOCIADO.*, TIPOSOCIO.NOME AS DESC_TIPOSOSIO " & _
    "FROM ASSOCIADO LEFT JOIN TIPOSOSIO ON ASSOCIADO.TIPOSOCIOID = TIPOSOCIO.PKID " & _
    "WHERE ASSOCIADO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarAssociado = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAssociado.SelecionarAssociado]", _
            Err.Description
End Function


Public Function SelecionarAssociadoPeloNome(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ASSOCIADO.* " & _
    "FROM ASSOCIADO " & _
    "WHERE ASSOCIADO.NOME = " & _
      Formata_Dados(strNome, tpDados_Texto) & _
      " AND ASSOCIADO.EXCLUIDO = " & Formata_Dados("N", tpDados_Texto)
  Set SelecionarAssociadoPeloNome = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPlano.SelecionarAssociadoPeloNome]", _
            Err.Description
End Function


