VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAgencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public Function ExcluirAgencia(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  'Delete AGENCIACNPJ
  strSql = "DELETE FROM AGENCIACNPJ " & _
    "WHERE AGENCIACNPJ.AGENCIAID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  'Delete AGENCIA
  strSql = "DELETE FROM AGENCIA " & _
    "WHERE AGENCIA.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAgencia.ExcluirAgencia]", _
            Err.Description
End Function


Public Function SelecionarAgenciaPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT AGENCIA.* " & _
    "FROM AGENCIA " & _
    "WHERE AGENCIA.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarAgenciaPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAgencia.SelecionarAgenciaPeloPkid]", _
            Err.Description
End Function

Public Function SelecionarAgenciaPeloNome(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT AGENCIA.* " & _
    "FROM AGENCIA " & _
    "WHERE AGENCIA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
    " AND AGENCIA.NOME = " & _
    Formata_Dados(strNome, tpDados_Texto)
  Set SelecionarAgenciaPeloNome = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAgencia.SelecionarAgenciaPeloNome]", _
            Err.Description
End Function


Public Sub AlterarAgencia(lngPKID As Long, _
                          strNome As String, _
                          strTel1 As String, strTel2 As String, strTel3 As String, _
                          strEndRua As String, _
                          strEndNumero As String, _
                          strEndCompl As String, _
                          strEndEstado As String, _
                          strEndCep As String, _
                          strEndBairro As String, _
                          strEndCidade As String, _
                          strObservacao As String, _
                          strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE AGENCIA SET " & _
    "NOME = " & Formata_Dados(strNome, tpDados_Texto) & ", " & _
    "TELEFONE1 = " & Formata_Dados(strTel1, tpDados_Texto) & ", " & _
    "TELEFONE2 = " & Formata_Dados(strTel2, tpDados_Texto) & ", " & _
    "TELEFONE3 = " & Formata_Dados(strTel3, tpDados_Texto) & ", "
  strSql = strSql & "ENDRUA = " & Formata_Dados(strEndRua, tpDados_Texto) & ", " & _
    "ENDNUMERO = " & Formata_Dados(strEndNumero, tpDados_Texto) & ", " & _
    "ENDCOMPLEMENTO = " & Formata_Dados(strEndCompl, tpDados_Texto) & ", " & _
    "ENDESTADO = " & Formata_Dados(strEndEstado, tpDados_Texto) & ", " & _
    "ENDCEP = " & Formata_Dados(strEndCep, tpDados_Texto) & ", " & _
    "ENDBAIRRO = " & Formata_Dados(strEndBairro, tpDados_Texto) & ", " & _
    "ENDCIDADE = " & Formata_Dados(strEndCidade, tpDados_Texto) & ", " & _
    "OBSERVACAO = " & Formata_Dados(strObservacao, tpDados_Texto) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAgencia.AlterarAgencia]", _
            Err.Description
End Sub

Public Sub InserirAgencia(lngPKID As Long, _
                          strNome As String, _
                          strTel1 As String, strTel2 As String, strTel3 As String, _
                          strEndRua As String, _
                          strEndNumero As String, _
                          strEndCompl As String, _
                          strEndEstado As String, _
                          strEndCep As String, _
                          strEndBairro As String, _
                          strEndCidade As String, _
                          strObservacao As String, _
                          strStatus As String)

  On Error GoTo trata
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO AGENCIA (NOME, TELEFONE1, TELEFONE2, TELEFONE3, " & _
    "ENDRUA, ENDNUMERO, ENDCOMPLEMENTO, ENDESTADO, ENDCEP, ENDBAIRRO, OBSERVACAO, ENDCIDADE, STATUS) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto) & ", " & _
    Formata_Dados(strTel1, tpDados_Texto) & ", " & _
    Formata_Dados(strTel2, tpDados_Texto) & ", " & _
    Formata_Dados(strTel3, tpDados_Texto) & ", "
  strSql = strSql & Formata_Dados(strEndRua, tpDados_Texto) & ", " & _
    Formata_Dados(strEndNumero, tpDados_Texto) & ", " & _
    Formata_Dados(strEndCompl, tpDados_Texto) & ", " & _
    Formata_Dados(strEndEstado, tpDados_Texto) & ", " & _
    Formata_Dados(strEndCep, tpDados_Texto) & ", " & _
    Formata_Dados(strEndBairro, tpDados_Texto) & ", " & _
    Formata_Dados(strEndCidade, tpDados_Texto) & ", " & _
    Formata_Dados(strObservacao, tpDados_Texto) & ", " & _
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
            Err.Source & ".[clsAgencia.InserirAgencia]", _
            Err.Description
End Sub



