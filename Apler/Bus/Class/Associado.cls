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

Public Function SelecionarAssociadoPeloNome(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objAssociado As datApler.clsAssociado
  Set objAssociado = New datApler.clsAssociado
  
  'Listar Associado
  Set SelecionarAssociadoPeloNome = objAssociado.SelecionarAssociadoPeloNome(strNome)
  '
  Set objAssociado = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarAssociadoPeloPkid(lngASSOCIADOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objAssociado As datApler.clsAssociado
  Set objAssociado = New datApler.clsAssociado
  
  'Listar Associado
  Set SelecionarAssociadoPeloPkid = objAssociado.SelecionarAssociadoPeloPkid(lngASSOCIADOID)
  '
  Set objAssociado = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirAssociado(lngASSOCIADOID As Long, _
                            strIcAssociado As String)
  On Error GoTo trata
  Dim objAssociado As datApler.clsAssociado
  Set objAssociado = New datApler.clsAssociado
  
  'Excluir Associado
  objAssociado.ExcluirAssociado lngASSOCIADOID, _
                                strIcAssociado
  '
  Set objAssociado = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
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
  Dim objAssociado     As datApler.clsAssociado
  '
  Set objAssociado = New datApler.clsAssociado
  'Inserir Associado
  objAssociado.InserirAssociado lngTIPOSOCIOID, _
                                lngValorPlanoId, _
                                lngESTADOCIVILID, _
                                strIdentidade, _
                                strOrgaoEmissor, _
                                strNaturalidade, _
                                strCpf, _
                                strNome, _
                                strSexo, _
                                strDtNascimento, _
                                strNomeMae, _
                                strEspecial, _
                                strEndRuaRes, _
                                strEndNumeroRes, _
                                strEndComplRes, _
                                strEndEstadoRes, _
                                strEndCepRes, _
                                strEndBairroRes, _
                                strEndCidadeRes, _
                                strTelefoneRes, _
                                strCelular, _
                                strEmail, _
                                strObservacao, _
                                strIcAssociado
  '
  Set objAssociado = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


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
  Dim objAssociado     As datApler.clsAssociado
  '
  Set objAssociado = New datApler.clsAssociado
  'Atualizar Associado
  objAssociado.AlterarAssociado lngPKID, _
                                lngTIPOSOCIOID, _
                                lngValorPlanoId, _
                                lngESTADOCIVILID, _
                                strIdentidade, _
                                strOrgaoEmissor, _
                                strNaturalidade, _
                                strCpf, _
                                strNome, _
                                strSexo, _
                                strDtNascimento, _
                                strNomeMae, _
                                strEspecial, _
                                strEndRuaRes, _
                                strEndNumeroRes, _
                                strEndComplRes, _
                                strEndEstadoRes, _
                                strEndCepRes, _
                                strEndBairroRes, _
                                strEndCidadeRes, _
                                strTelefoneRes, _
                                strCelular, _
                                strEmail, _
                                strExcluido, _
                                strObservacao
  '
  Set objAssociado = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

