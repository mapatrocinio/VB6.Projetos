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

Public Function ExcluirEmpresa(strNome As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM EMPRESA " & _
    "WHERE EMPRESA.PKID = " & _
    Formata_Dados(strNome, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresa.ExcluirEmpresa]", _
            Err.Description
End Function


Public Function SelecionarEmpresaPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT EMPRESA.*, TIPOEMPRESA.DESCRICAO AS DESC_TIPOEMPRESA " & _
    "FROM EMPRESA LEFT JOIN TIPOEMPRESA ON TIPOEMPRESA.PKID = EMPRESA.TIPOEMPRESAID " & _
    "WHERE EMPRESA.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarEmpresaPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresa.SelecionarEmpresaPeloPkid]", _
            Err.Description
End Function

Public Function SelecionarEmpresaPeloNome(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT EMPRESA.*, TIPOEMPRESA.DESCRICAO AS DESC_TIPOEMPRESA " & _
    "FROM EMPRESA LEFT JOIN TIPOEMPRESA ON TIPOEMPRESA.PKID = EMPRESA.TIPOEMPRESAID " & _
    "WHERE EMPRESA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
    " AND EMPRESA.NOME = " & _
    Formata_Dados(strNome, tpDados_Texto)
  Set SelecionarEmpresaPeloNome = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresa.SelecionarEmpresaPeloNome]", _
            Err.Description
End Function


Public Sub AlterarEmpresa(lngPKID As Long, _
                           lngTIPOEMPRESAID As Long, _
                           strNome As String, _
                           strNomeFantasia As String, _
                           strCnpj As String, _
                           strInscrEstd As String, _
                           strInscrMunic As String, _
                           strTel1 As String, strTel2 As String, strTel3 As String, strFax As String, _
                           strEmail As String, _
                           strContato As String, _
                           strTelContato As String, _
                           strEndRua As String, _
                           strEndNumero As String, _
                           strEndCompl As String, _
                           strEndEstado As String, _
                           strEndCep As String, _
                           strEndBairro As String, _
                           strEndCidade As String, _
                           strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE EMPRESA SET " & _
    "TIPOEMPRESAID = " & Formata_Dados(IIf(lngTIPOEMPRESAID = 0, "", lngTIPOEMPRESAID), tpDados_Longo) & ", " & _
    "NOME = " & Formata_Dados(strNome, tpDados_Texto) & ", " & _
    "NOMEFANTASIA = " & Formata_Dados(strNomeFantasia, tpDados_Texto) & ", " & _
    "CNPJ = " & Formata_Dados(strCnpj, tpDados_Texto) & ", " & _
    "INSCRICAOESTADUAL = " & Formata_Dados(strInscrEstd, tpDados_Texto) & ", " & _
    "INSCRICAOMUNICIPAL = " & Formata_Dados(strInscrMunic, tpDados_Texto) & ", " & _
    "TELEFONE1 = " & Formata_Dados(strTel1, tpDados_Texto) & ", " & _
    "TELEFONE2 = " & Formata_Dados(strTel2, tpDados_Texto) & ", " & _
    "TELEFONE3 = " & Formata_Dados(strTel3, tpDados_Texto) & ", "
  strSql = strSql & "FAX = " & Formata_Dados(strFax, tpDados_Texto) & ", " & _
    "EMAIL = " & Formata_Dados(strEmail, tpDados_Texto) & ", " & _
    "CONTATO = " & Formata_Dados(strContato, tpDados_Texto) & ", " & _
    "TELEFONECONTATO = " & Formata_Dados(strTelContato, tpDados_Texto) & ", " & _
    "ENDRUA = " & Formata_Dados(strEndRua, tpDados_Texto) & ", " & _
    "ENDNUMERO = " & Formata_Dados(strEndNumero, tpDados_Texto) & ", " & _
    "ENDCOMPL = " & Formata_Dados(strEndCompl, tpDados_Texto) & ", " & _
    "ENDESTADO = " & Formata_Dados(strEndEstado, tpDados_Texto) & ", " & _
    "ENDCEP = " & Formata_Dados(strEndCep, tpDados_Texto) & ", " & _
    "ENDBAIRRO = " & Formata_Dados(strEndBairro, tpDados_Texto) & ", " & _
    "ENDCIDADE = " & Formata_Dados(strEndCidade, tpDados_Texto) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresa.AlterarEmpresa]", _
            Err.Description
End Sub

Public Sub InserirEmpresa(ByRef lngPKID As Long, _
                           lngTIPOEMPRESAID As Long, _
                           strNome As String, _
                           strNomeFantasia As String, _
                           strCnpj As String, _
                           strInscrEstd As String, _
                           strInscrMunic As String, _
                           strTel1 As String, strTel2 As String, strTel3 As String, strFax As String, _
                           strEmail As String, _
                           strContato As String, _
                           strTelContato As String, _
                           strEndRua As String, _
                           strEndNumero As String, _
                           strEndCompl As String, _
                           strEndEstado As String, _
                           strEndCep As String, _
                           strEndBairro As String, _
                           strEndCidade As String, _
                           strStatus As String)

  On Error GoTo trata
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO EMPRESA (TIPOEMPRESAID, NOME, NOMEFANTASIA, CNPJ, INSCRICAOESTADUAL, INSCRICAOMUNICIPAL, TELEFONE1, TELEFONE2, TELEFONE3, " & _
    "FAX, EMAIL, CONTATO, TELEFONECONTATO, ENDRUA, ENDNUMERO, ENDCOMPL, ENDESTADO, ENDCEP, ENDBAIRRO, ENDCIDADE, STATUS) VALUES (" & _
    Formata_Dados(IIf(lngTIPOEMPRESAID = 0, "", lngTIPOEMPRESAID), tpDados_Longo) & ", " & _
    Formata_Dados(strNome, tpDados_Texto) & ", " & _
    Formata_Dados(strNomeFantasia, tpDados_Texto) & ", " & _
    Formata_Dados(strCnpj, tpDados_Texto) & ", " & _
    Formata_Dados(strInscrEstd, tpDados_Texto) & ", " & _
    Formata_Dados(strInscrMunic, tpDados_Texto) & ", " & _
    Formata_Dados(strTel1, tpDados_Texto) & ", " & _
    Formata_Dados(strTel2, tpDados_Texto) & ", " & _
    Formata_Dados(strTel3, tpDados_Texto) & ", "
  strSql = strSql & Formata_Dados(strFax, tpDados_Texto) & ", " & _
    Formata_Dados(strEmail, tpDados_Texto) & ", " & _
    Formata_Dados(strContato, tpDados_Texto) & ", " & _
    Formata_Dados(strTelContato, tpDados_Texto) & ", " & _
    Formata_Dados(strEndRua, tpDados_Texto) & ", " & _
    Formata_Dados(strEndNumero, tpDados_Texto) & ", " & _
    Formata_Dados(strEndCompl, tpDados_Texto) & ", " & _
    Formata_Dados(strEndEstado, tpDados_Texto) & ", " & _
    Formata_Dados(strEndCep, tpDados_Texto) & ", " & _
    Formata_Dados(strEndBairro, tpDados_Texto) & ", " & _
    Formata_Dados(strEndCidade, tpDados_Texto) & ", " & _
    Formata_Dados(strStatus, tpDados_Texto) & ");"
  '
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngPKID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEmpresa.InserirEmpresa]", _
            Err.Description
End Sub





