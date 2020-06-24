VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConfiguracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirConfiguracao(lngCONFIGURACAOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui CONFIGURACAO
  strSql = "DELETE FROM CONFIGURACAO WHERE CONFIGURACAO.PKID = " & _
    Formata_Dados(lngCONFIGURACAOID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConfiguracao.ExcluirConfiguracao]", _
            Err.Description
End Sub

Public Function VerificaExclusaoConfiguracao(lngCONFIGURACAOID As Long, _
                                             ByRef strDescricaoErro As String) As Boolean
  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  Dim blnRet As Boolean
  strDescricaoErro = ""
  blnRet = True
'''  'LOCAÇÃO
'''  strSql = "SELECT COUNT(*) AS TOTALLANC FROM LOCACAO WHERE CONFIGURACAOID = " & _
'''    Formata_Dados(lngCONFIGURACAOID, tpDados_Longo)
'''  Set objRs = RunSPReturnRS(strSql)
'''  If objRs.EOF Then
'''    blnRet = True
'''  Else
'''    If objRs.Fields("TOTALLANC").Value = 0 Then
'''      blnRet = True
'''    Else
'''      blnRet = False
'''      strDescricaoErro = "Tabela de Locação"
'''    End If
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
  VerificaExclusaoConfiguracao = blnRet
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConfiguracao.VerificaExclusaoConfiguracao]", _
            Err.Description
End Function
Public Sub AlterarConfiguracaoCadastro(lngCONFIGURACAOID As Long, _
                                       strEmpresa As String, _
                                       strTitulo As String, _
                                       strCnpj As String, _
                                       strInscrMunicipal As String, _
                                       strEndereco As String, _
                                       strBairro As String, _
                                       strCidade As String, _
                                       strEstado As String, _
                                       strCep As String, _
                                       strTel As String)
  On Error GoTo trata
  '
  Dim strSql As String
  strSql = "UPDATE CONFIGURACAO SET " & _
    "Empresa = " & Formata_Dados(strEmpresa, tpDados_Texto) & _
    ", Titulo= " & Formata_Dados(strTitulo, tpDados_Texto) & _
    ", Cnpj = " & Formata_Dados(strCnpj, tpDados_Texto) & _
    ", InscrMunicipal = " & Formata_Dados(strInscrMunicipal, tpDados_Texto) & _
    ", Endereco = " & Formata_Dados(strEndereco, tpDados_Texto) & _
    ", Bairro = " & Formata_Dados(strBairro, tpDados_Texto) & _
    ", Cidade = " & Formata_Dados(strCidade, tpDados_Texto) & _
    ", Estado = " & Formata_Dados(strEstado, tpDados_Texto) & _
    ", Cep = " & Formata_Dados(strCep, tpDados_Texto) & _
    ", Tel = " & Formata_Dados(strTel, tpDados_Texto) & _
    " WHERE PKID = " & lngCONFIGURACAOID
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConfiguracao.AlterarConfiguracaoCadastro]", _
            Err.Description
End Sub

Public Sub InserirConfiguracao()
  On Error GoTo trata
  '
  Dim strSql As String
  strSql = "INSERT INTO CONFIGURACAO (Empresa, Endereco, Bairro, Cidade, Estado, Tel, Cep, TITULO) " & _
    "VALUES ('Empresa', 'Endereco', 'Bairro', 'Cidade', 'Es', 'Tel', 'Cep', 'TITULO')"
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConfiguracao.InserirConfiguracao]", _
            Err.Description
End Sub

Public Function ListarConfiguracao() As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CONFIGURACAO.* " & _
    "FROM CONFIGURACAO "
  Set ListarConfiguracao = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConfiguracao.ListarConfiguracao]", _
            Err.Description
End Function

Public Function ListarConfiguracaoCadastro(lngCONFIGURACAOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CONFIGURACAO.* " & _
    "FROM CONFIGURACAO " & _
    "WHERE CONFIGURACAO.PKID = " & _
      Formata_Dados(lngCONFIGURACAOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarConfiguracaoCadastro = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConfiguracao.ListarConfiguracaoCadastro]", _
            Err.Description
End Function
