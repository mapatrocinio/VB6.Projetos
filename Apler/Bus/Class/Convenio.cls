VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarConvenioPeloNome(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objConvenio As datApler.clsConvenio
  Set objConvenio = New datApler.clsConvenio
  
  'Listar Convenio
  Set SelecionarConvenioPeloNome = objConvenio.SelecionarConvenioPeloNome(strNome)
  '
  Set objConvenio = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarConvenioPeloPkid(lngCONVENIOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objConvenio As datApler.clsConvenio
  Set objConvenio = New datApler.clsConvenio
  
  'Listar Convenio
  Set SelecionarConvenioPeloPkid = objConvenio.SelecionarConvenioPeloPkid(lngCONVENIOID)
  '
  Set objConvenio = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirConvenio(lngCONVENIOID As Long)
  On Error GoTo trata
  Dim objConvenio As datApler.clsConvenio
  Set objConvenio = New datApler.clsConvenio
  
  'Excluir Convenio
  objConvenio.ExcluirConvenio lngCONVENIOID
  '
  Set objConvenio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirConvenio(ByRef lngPKID As Long, _
                           lngTIPOCONVENIOID As Long, _
                           strNome As String, _
                           strNomeFantasia As String, _
                           strNumeroContrato As String, _
                           strCodigo As String, _
                           strDtInicio As String, _
                           strDtFim As String, _
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
                           strStatus As String, _
                           strStatusAnterior As String)
  On Error GoTo trata
  Dim objConvenio     As datApler.clsConvenio
  '
  Set objConvenio = New datApler.clsConvenio
  'Inserir Convenio
  objConvenio.InserirConvenio lngPKID, _
                              lngTIPOCONVENIOID, _
                              strNome, _
                              strNomeFantasia, _
                              strNumeroContrato, _
                              strCodigo, _
                              strDtInicio, _
                              strDtFim, _
                              strCnpj, _
                              strInscrEstd, _
                              strInscrMunic, _
                              strTel1, strTel2, strTel3, strFax, _
                              strEmail, _
                              strContato, _
                              strTelContato, _
                              strEndRua, _
                              strEndNumero, _
                              strEndCompl, _
                              strEndEstado, _
                              strEndCep, _
                              strEndBairro, _
                              strEndCidade, _
                              strStatus, _
                              strStatusAnterior
  '
  Set objConvenio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarConvenio(lngPKID As Long, _
                           lngTIPOCONVENIOID As Long, _
                           strNome As String, _
                           strNomeFantasia As String, _
                           strNumeroContrato As String, _
                           strCodigo As String, _
                           strDtInicio As String, _
                           strDtFim As String, _
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
                           strStatus As String, _
                           strStatusAnterior As String)

  On Error GoTo trata
  Dim objConvenio     As datApler.clsConvenio
  '
  Set objConvenio = New datApler.clsConvenio
  'Atualizar Convenio
  objConvenio.AlterarConvenio lngPKID, _
                              lngTIPOCONVENIOID, _
                              strNome, _
                              strNomeFantasia, _
                              strNumeroContrato, _
                              strCodigo, _
                              strDtInicio, _
                              strDtFim, _
                              strCnpj, _
                              strInscrEstd, _
                              strInscrMunic, _
                              strTel1, strTel2, strTel3, strFax, _
                              strEmail, _
                              strContato, _
                              strTelContato, _
                              strEndRua, _
                              strEndNumero, _
                              strEndCompl, _
                              strEndEstado, _
                              strEndCep, _
                              strEndBairro, _
                              strEndCidade, _
                              strStatus, _
                              strStatusAnterior
  '
  Set objConvenio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

