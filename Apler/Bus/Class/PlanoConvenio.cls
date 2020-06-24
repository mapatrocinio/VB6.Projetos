VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlanoConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarPlanoConvenioPeloPkid(lngPlanoConvenioId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objPlanoConvenio As datApler.clsPlanoConvenio
  Set objPlanoConvenio = New datApler.clsPlanoConvenio
  
  'Listar PlanoConvenio
  Set SelecionarPlanoConvenioPeloPkid = objPlanoConvenio.SelecionarPlanoConvenioPeloPkid(lngPlanoConvenioId)
  '
  Set objPlanoConvenio = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirPlanoConvenio(lngPlanoConvenioId As Long)
  On Error GoTo trata
  Dim objPlanoConvenio As datApler.clsPlanoConvenio
  Set objPlanoConvenio = New datApler.clsPlanoConvenio
  
  'Excluir PlanoConvenio
  objPlanoConvenio.ExcluirPlanoConvenio lngPlanoConvenioId
  '
  Set objPlanoConvenio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirPlanoConvenio(lngCONVENIOID As Long, _
                                strNome As String, _
                                strValorSocio As String, _
                                strValorDependente As String, _
                                strValorAplerSocio As String, _
                                strValorAplerDependente As String, _
                                strDataInicio As String, _
                                strDataFim As String)
  On Error GoTo trata
  Dim objPlanoConvenio     As datApler.clsPlanoConvenio
  '
  Set objPlanoConvenio = New datApler.clsPlanoConvenio
  'Inserir PlanoConvenio
  objPlanoConvenio.InserirPlanoConvenio lngCONVENIOID, _
                                        strNome, _
                                        strValorSocio, _
                                        strValorDependente, _
                                        strValorAplerSocio, _
                                        strValorAplerDependente, _
                                        strDataInicio, _
                                        strDataFim
  '
  Set objPlanoConvenio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarPlanoConvenio(lngPKID As Long, _
                                strNome As String, _
                                strValorSocio As String, _
                                strValorDependente As String, _
                                strValorAplerSocio As String, _
                                strValorAplerDependente As String, _
                                strDataInicio As String, _
                                strDataFim As String, _
                                strStatus As String)

  On Error GoTo trata
  Dim objPlanoConvenio     As datApler.clsPlanoConvenio
  '
  Set objPlanoConvenio = New datApler.clsPlanoConvenio
  'Atualizar PlanoConvenio
  objPlanoConvenio.AlterarPlanoConvenio lngPKID, _
                                        strNome, _
                                        strValorSocio, _
                                        strValorDependente, _
                                        strValorAplerSocio, _
                                        strValorAplerDependente, _
                                        strDataInicio, _
                                        strDataFim, _
                                        strStatus
  '
  Set objPlanoConvenio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

