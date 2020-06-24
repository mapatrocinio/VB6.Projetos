VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPessoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarPessoaPeloNome(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objPessoa As datSisMetal.clsPessoa
  Set objPessoa = New datSisMetal.clsPessoa
  
  'Listar Pessoa
  Set SelecionarPessoaPeloNome = objPessoa.SelecionarPessoaPeloNome(strNome)
  '
  Set objPessoa = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarPessoaPeloPkid(lngPESSOAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objPessoa As datSisMetal.clsPessoa
  Set objPessoa = New datSisMetal.clsPessoa
  
  'Listar Pessoa
  Set SelecionarPessoaPeloPkid = objPessoa.SelecionarPessoaPeloPkid(lngPESSOAID)
  '
  Set objPessoa = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirPessoa(lngPESSOAID As Long)
  On Error GoTo trata
  Dim objPessoa As datSisMetal.clsPessoa
  Set objPessoa = New datSisMetal.clsPessoa
  
  'Excluir Pessoa
  objPessoa.ExcluirPessoa lngPESSOAID
  '
  Set objPessoa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirPessoa(lngPKID As Long, _
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
  Dim objPessoa     As datSisMetal.clsPessoa
  '
  Set objPessoa = New datSisMetal.clsPessoa
  'Inserir Pessoa
  objPessoa.InserirPessoa lngPKID, _
                           strDtNascimento, _
                           strNome, _
                           strRgNumero, _
                           strRgOrgao, _
                           strRgDtExpedicao, _
                           strTipoPessoa, _
                           strCNPJ, _
                           strCPF, _
                           strSexo, _
                           strTelefone, _
                           strCelular, _
                           strEndRuaRes, _
                           strEndNumeroRes, _
                           strEndComplRes, _
                           strEndEstadoRes, _
                           strEndCepRes, _
                           strEndBairroRes, _
                           strEndCidadeRes, _
                           strObservacao
  '
  Set objPessoa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarPessoa(lngPKID As Long, _
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
  Dim objPessoa     As datSisMetal.clsPessoa
  '
  Set objPessoa = New datSisMetal.clsPessoa
  'Atualizar Pessoa
  objPessoa.AlterarPessoa lngPKID, _
                           strDtNascimento, _
                           strNome, _
                           strRgNumero, _
                           strRgOrgao, _
                           strRgDtExpedicao, _
                           strTipoPessoa, _
                           strCNPJ, _
                           strCPF, _
                           strSexo, _
                           strTelefone, _
                           strCelular, _
                           strEndRuaRes, _
                           strEndNumeroRes, _
                           strEndComplRes, _
                           strEndEstadoRes, _
                           strEndCepRes, _
                           strEndBairroRes, _
                           strEndCidadeRes, _
                           strObservacao
  '
  Set objPessoa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

