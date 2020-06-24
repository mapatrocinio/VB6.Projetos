VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarLoja(intTipoLoja As Integer, _
                               strNome As String, _
                               lngLOJAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT LOJA.* " & _
    "FROM LOJA "
  '
  Select Case intTipoLoja
  Case tpLoja.tpLoja_Anodizadora
    strSql = strSql & " INNER JOIN ANODIZADORA ON ANODIZADORA.LOJAID = LOJA.PKID "
  Case tpLoja.tpLoja_Empresa
    strSql = strSql & " INNER JOIN EMPRESA ON EMPRESA.LOJAID = LOJA.PKID "
  Case tpLoja.tpLoja_Fabrica
    strSql = strSql & " INNER JOIN FABRICA ON FABRICA.LOJAID = LOJA.PKID "
  Case tpLoja.tpLoja_Filial
    strSql = strSql & " INNER JOIN FILIAL ON FILIAL.LOJAID = LOJA.PKID "
  Case tpLoja.tpLoja_Fornecedor
    strSql = strSql & " INNER JOIN FORNECEDOR ON FORNECEDOR.LOJAID = LOJA.PKID "
  End Select
  '
  strSql = strSql & "WHERE " & _
    " LOJA.NOME = " & _
      Formata_Dados(strNome, tpDados_Texto) & _
    " AND LOJA.PKID <> " & Formata_Dados(lngLOJAID, tpDados_Longo)
  Set SelecionarLoja = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLoja.SelecionarLoja]", _
            Err.Description
End Function

Public Function ExcluirLoja(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  'ANODIZADORA
  strSql = "DELETE FROM ANODIZADORA " & _
    "WHERE LOJAID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  'FABRICA
  strSql = "DELETE FROM FABRICA " & _
    "WHERE LOJAID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  'FILIAL
  strSql = "DELETE FROM FILIAL " & _
    "WHERE LOJAID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  'FORNECEDOR
  strSql = "DELETE FROM FORNECEDOR " & _
    "WHERE LOJAID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  'LOJA
  strSql = "DELETE FROM LOJA " & _
    "WHERE PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLoja.ExcluirLoja]", _
            Err.Description
End Function


Public Function SelecionarLojaPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT LOJA.*, FORNECEDOR.VALOR_KG " & _
    "FROM LOJA LEFT JOIN FORNECEDOR ON LOJA.PKID = FORNECEDOR.LOJAID " & _
    "WHERE LOJA.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarLojaPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLoja.SelecionarLojaPeloPkid]", _
            Err.Description
End Function



Public Function SelecionarFornecedorPeloNome(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT LOJA.*, FORNECEDOR.VALOR_KG " & _
    "FROM LOJA " & _
    " INNER JOIN FORNECEDOR ON LOJA.PKID = FORNECEDOR.LOJAID " & _
    "WHERE LOJA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
    " AND LOJA.NOME = " & _
    Formata_Dados(strNome, tpDados_Texto)
  Set SelecionarFornecedorPeloNome = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLoja.SelecionarFornecedorPeloNome]", _
            Err.Description
End Function

Public Function SelecionarLojaPeloNome(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT LOJA.* " & _
    "FROM LOJA " & _
    "WHERE LOJA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
    " AND LOJA.NOME = " & _
    Formata_Dados(strNome, tpDados_Texto)
  Set SelecionarLojaPeloNome = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLoja.SelecionarLojaPeloNome]", _
            Err.Description
End Function


Public Sub AlterarLoja(lngPKID As Long, _
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

  strSql = "UPDATE LOJA SET " & _
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
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLoja.AlterarLoja]", _
            Err.Description
End Sub

Public Sub InserirLoja(ByRef lngPKID As Long, _
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
  strSql = strSql & "INSERT INTO LOJA (NOME, NOMEFANTASIA, CNPJ, INSCRICAOESTADUAL, INSCRICAOMUNICIPAL, TELEFONE1, TELEFONE2, TELEFONE3, " & _
    "FAX, EMAIL, CONTATO, TELEFONECONTATO, ENDRUA, ENDNUMERO, ENDCOMPL, ENDESTADO, ENDCEP, ENDBAIRRO, ENDCIDADE, STATUS) VALUES (" & _
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
            Err.Source & ".[clsLoja.InserirLoja]", _
            Err.Description
End Sub

Public Sub AlterarFornecedor(lngPKID As Long, _
                             strValorKg As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "UPDATE FORNECEDOR " & _
    "SET VALOR_KG = " & Formata_Dados(strValorKg, tpDados_Moeda) & " " & _
    " WHERE LOJAID = " & Formata_Dados(lngPKID, tpDados_Longo)
  '

  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLoja.InserirFornecedor]", _
            Err.Description
End Sub

Public Sub InserirFornecedor(lngPKID As Long, _
                             strValorKg As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO FORNECEDOR (LOJAID, VALOR_KG) VALUES (" & _
    Formata_Dados(lngPKID, tpDados_Longo) & ", " & _
    Formata_Dados(strValorKg, tpDados_Moeda) & " " & ");"
  '

  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLoja.InserirFornecedor]", _
            Err.Description
End Sub


Public Sub InserirFilial(ByRef lngPKID As Long)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO FILIAL (LOJAID) VALUES (" & _
    Formata_Dados(lngPKID, tpDados_Longo) & " " & ");"
  '

  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLoja.InserirFilial]", _
            Err.Description
End Sub
Public Sub InserirEmpresa(ByRef lngPKID As Long)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO EMPRESA (LOJAID) VALUES (" & _
    Formata_Dados(lngPKID, tpDados_Longo) & " " & ");"
  '

  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLoja.InserirEmpresa]", _
            Err.Description
End Sub

Public Sub InserirFabrica(ByRef lngPKID As Long)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO FABRICA (LOJAID) VALUES (" & _
    Formata_Dados(lngPKID, tpDados_Longo) & " " & ");"
  '

  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLoja.InserirFabrica]", _
            Err.Description
End Sub

Public Sub InserirAnodizadora(ByRef lngPKID As Long)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO ANODIZADORA (LOJAID) VALUES (" & _
    Formata_Dados(lngPKID, tpDados_Longo) & " " & ");"
  '

  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLoja.InserirAnodizadora]", _
            Err.Description
End Sub




