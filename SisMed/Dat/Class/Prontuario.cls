VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProntuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirProntuario(lngPRONTUARIOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  'CONVENIOS PRONTUARIOS
''  strSql = "DELETE FROM TAB_CONVASSOC " & _
''    "WHERE TAB_CONVASSOC.PRONTUARIOID = " & _
''    Formata_Dados(lngPRONTUARIOID, tpDados_Longo, tpNulo_Aceita)
''  RunSP strSql
  'MENSALIDADES DE CONVÊNIO
  'strSql = "DELETE FROM MENSALIDADECONVENIO " & _
    "WHERE MENSALIDADECONVENIO.PRONTUARIOID = " & _
    Formata_Dados(lngPRONTUARIOID, tpDados_Longo, tpNulo_Aceita)
  'RunSP strSql
  'MENSALIDADES PLANO
  'strSql = "DELETE FROM MENSALIDADEPLANO " & _
    "WHERE MENSALIDADEPLANO.PRONTUARIOID = " & _
    Formata_Dados(lngPRONTUARIOID, tpDados_Longo, tpNulo_Aceita)
  'RunSP strSql
  'MENSALIDADES TAB_PROFASSOC
'''  strSql = "DELETE FROM TAB_PROFASSOC " & _
'''    "WHERE TAB_PROFASSOC.PRONTUARIOID = " & _
'''    Formata_Dados(lngPRONTUARIOID, tpDados_Longo, tpNulo_Aceita)
'''  RunSP strSql
  
  strSql = "DELETE FROM FUNCIONARIO " & _
    "WHERE FUNCIONARIO.PRONTUARIOID = " & _
    Formata_Dados(lngPRONTUARIOID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  strSql = "DELETE FROM PACIENTE " & _
    "WHERE PACIENTE.PRONTUARIOID = " & _
    Formata_Dados(lngPRONTUARIOID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  strSql = "DELETE FROM PRESTADOR " & _
    "WHERE PRESTADOR.PRONTUARIOID = " & _
    Formata_Dados(lngPRONTUARIOID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  
  strSql = "DELETE FROM PRONTUARIO " & _
    "WHERE PRONTUARIO.PKID = " & _
    Formata_Dados(lngPRONTUARIOID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProntuario.ExcluirProntuario]", _
            Err.Description
End Function

Public Function SelecionarProntuarioPeloPkid(lngPRONTUARIOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  strSql = "SELECT PRONTUARIO.* " & _
    "FROM PRONTUARIO " & _
    "WHERE PRONTUARIO.PKID = " & _
    Formata_Dados(lngPRONTUARIOID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarProntuarioPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProntuario.SelecionarProntuarioPeloPkid]", _
            Err.Description
End Function

Public Sub AlterarProntuario(lngPKID As Long, _
                           strDtNascimento As String, _
                           strNome As String, _
                           strRgNumero As String, _
                           strRgOrgao As String, _
                           strRgDtExpedicao As String, _
                           strTipoPessoa As String, _
                           strCNPJ As String, _
                           strCPF As String, _
                           strSexo As String, _
                           strTelefone As String, _
                           strCelular As String, _
                           strEndRuaRes As String, _
                           strEndNumeroRes As String, _
                           strEndComplRes As String, _
                           strEndEstadoRes As String, _
                           strEndCepRes As String, _
                           strEndBairroRes As String, _
                           strEndCidadeRes As String, _
                           strObservacao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE PRONTUARIO SET " & _
    "DTNASCIMENTO = " & Formata_Dados(strDtNascimento, tpDados_DataHora) & ", " & _
    "NOME = " & Formata_Dados(strNome, tpDados_Texto) & ", " & _
    "RGNUMERO = " & Formata_Dados(strRgNumero, tpDados_Texto) & ", " & _
    "RGORGAO = " & Formata_Dados(strRgOrgao, tpDados_Texto) & ", " & _
    "RGDTEXPEDICAO = " & Formata_Dados(strRgDtExpedicao, tpDados_DataHora) & ", " & _
    "TIPO_PESSOA = " & Formata_Dados(strTipoPessoa, tpDados_Texto) & ", " & _
    "CNPJ = " & Formata_Dados(strCNPJ, tpDados_Texto) & ", " & _
    "CPF = " & Formata_Dados(strCPF, tpDados_Texto) & ", " & _
    "SEXO = " & Formata_Dados(strSexo, tpDados_Texto) & ", " & _
    "TELEFONE = " & Formata_Dados(strTelefone, tpDados_Texto) & ", " & _
    "CELULAR = " & Formata_Dados(strCelular, tpDados_Texto) & ", " & _
    "ENDRUA = " & Formata_Dados(strEndRuaRes, tpDados_Texto) & ", "
  strSql = strSql & "ENDNUMERO = " & Formata_Dados(strEndNumeroRes, tpDados_Texto) & ", " & _
    "ENDCOMPLEMENTO = " & Formata_Dados(strEndComplRes, tpDados_Texto) & ", " & _
    "ENDCEP = " & Formata_Dados(strEndCepRes, tpDados_Texto) & ", " & _
    "ENDBAIRRO = " & Formata_Dados(strEndBairroRes, tpDados_Texto) & ", " & _
    "ENDCIDADE = " & Formata_Dados(strEndCidadeRes, tpDados_Texto) & ", " & _
    "ENDESTADO = " & Formata_Dados(strEndEstadoRes, tpDados_Texto) & ", " & _
    "ENDPAIS = " & Formata_Dados("", tpDados_Texto) & ", " & _
    "OBSERVACAO = " & Formata_Dados(strObservacao, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProntuario.AlterarProntuario]", _
            Err.Description
End Sub

Public Sub InserirProntuario(lngPKID As Long, _
                           strDtNascimento As String, _
                           strNome As String, _
                           strRgNumero As String, _
                           strRgOrgao As String, _
                           strRgDtExpedicao As String, _
                           strTipoPessoa As String, _
                           strCNPJ As String, _
                           strCPF As String, _
                           strSexo As String, _
                           strTelefone As String, _
                           strCelular As String, _
                           strEndRuaRes As String, _
                           strEndNumeroRes As String, _
                           strEndComplRes As String, _
                           strEndEstadoRes As String, _
                           strEndCepRes As String, _
                           strEndBairroRes As String, _
                           strEndCidadeRes As String, _
                           strObservacao As String)
  On Error GoTo trata
  Dim strSql  As String
  Dim objRs   As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO PRONTUARIO (DTCADASTRO,DTNASCIMENTO,NOME,RGNUMERO,RGORGAO,RGDTEXPEDICAO,TIPO_PESSOA,CNPJ,CPF,SEXO,TELEFONE,CELULAR,ENDRUA,ENDNUMERO,ENDCOMPLEMENTO,ENDCEP,ENDBAIRRO,ENDCIDADE,ENDESTADO,ENDPAIS,OBSERVACAO) VALUES (" & _
    " " & Formata_Dados(Format(Now, "DD/MM/YYYY hh:mm"), tpDados_DataHora) & _
    ", " & Formata_Dados(strDtNascimento, tpDados_DataHora) & _
    ", " & Formata_Dados(strNome, tpDados_Texto) & _
    ", " & Formata_Dados(strRgNumero, tpDados_Texto) & _
    ", " & Formata_Dados(strRgOrgao, tpDados_Texto) & _
    ", " & Formata_Dados(strRgDtExpedicao, tpDados_DataHora) & _
    ", " & Formata_Dados(strTipoPessoa, tpDados_Texto) & _
    ", " & Formata_Dados(strCNPJ, tpDados_Texto) & _
    ", " & Formata_Dados(strCPF, tpDados_Texto) & _
    ", " & Formata_Dados(strSexo, tpDados_Texto) & _
    ", " & Formata_Dados(strTelefone, tpDados_Texto) & _
    ", " & Formata_Dados(strCelular, tpDados_Texto) & _
    ", " & Formata_Dados(strEndRuaRes, tpDados_Texto)
  strSql = strSql & ", " & Formata_Dados(strEndNumeroRes, tpDados_Texto) & _
    ", " & Formata_Dados(strEndComplRes, tpDados_Texto) & _
    ", " & Formata_Dados(strEndCepRes, tpDados_Texto) & _
    ", " & Formata_Dados(strEndBairroRes, tpDados_Texto) & _
    ", " & Formata_Dados(strEndCidadeRes, tpDados_Texto) & _
    ", " & Formata_Dados(strEndEstadoRes, tpDados_Texto) & _
    ", " & Formata_Dados("", tpDados_Texto) & _
    ", " & Formata_Dados(strObservacao, tpDados_Texto) & ");"
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngPKID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProntuario.InserirProntuario]", _
            Err.Description
End Sub

Public Function SelecionarProntuario(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PRONTUARIO.*, TIPOSOCIO.NOME AS DESC_TIPOSOSIO " & _
    "FROM PRONTUARIO LEFT JOIN TIPOSOSIO ON PRONTUARIO.TIPOSOCIOID = TIPOSOCIO.PKID " & _
    "WHERE PRONTUARIO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarProntuario = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProntuario.SelecionarProntuario]", _
            Err.Description
End Function


Public Function SelecionarProntuarioPeloNome(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PRONTUARIO.* " & _
    "FROM PRONTUARIO " & _
    "WHERE PRONTUARIO.NOME = " & _
      Formata_Dados(strNome, tpDados_Texto)
  Set SelecionarProntuarioPeloNome = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPlano.SelecionarProntuarioPeloNome]", _
            Err.Description
End Function


