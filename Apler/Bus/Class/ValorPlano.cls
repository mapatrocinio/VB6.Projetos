VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsValorPlano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarValorPlanoPelaDescricao(lngTIPOVALORPLANOID As Long, _
                                                  strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objValorPlano As datApler.clsValorPlano
  Set objValorPlano = New datApler.clsValorPlano
  
  'Listar ValorPlano
  Set SelecionarValorPlanoPelaDescricao = objValorPlano.SelecionarValorPlanoPelaDescricao(lngTIPOVALORPLANOID, _
                                                                                          strDescricao)
  '
  Set objValorPlano = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarValorPlanoPeloPkid(lngVALORPLANOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objValorPlano As datApler.clsValorPlano
  Set objValorPlano = New datApler.clsValorPlano
  
  'Listar ValorPlano
  Set SelecionarValorPlanoPeloPkid = objValorPlano.SelecionarValorPlanoPeloPkid(lngVALORPLANOID)
  '
  Set objValorPlano = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirValorPlano(lngVALORPLANOID As Long)
  On Error GoTo trata
  Dim objValorPlano As datApler.clsValorPlano
  Set objValorPlano = New datApler.clsValorPlano
  
  'Excluir ValorPlano
  objValorPlano.ExcluirValorPlano lngVALORPLANOID
  '
  Set objValorPlano = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirValorPlano(lngTipoPlanoId As Long, _
                             strDescricao As String, _
                             strValorTitular As String, _
                             strValorDependente As String)
  On Error GoTo trata
  Dim objValorPlano     As datApler.clsValorPlano
  '
  Set objValorPlano = New datApler.clsValorPlano
  'Inserir ValorPlano
  objValorPlano.InserirValorPlano lngTipoPlanoId, _
                                  strDescricao, _
                                  strValorTitular, _
                                  strValorDependente
  '
  Set objValorPlano = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarValorPlano(lngPKID As Long, _
                             strDescricao As String, _
                             strValorTitular As String, _
                             strValorDependente As String, _
                             strStatus As String)

  On Error GoTo trata
  Dim objValorPlano     As datApler.clsValorPlano
  '
  Set objValorPlano = New datApler.clsValorPlano
  'Atualizar ValorPlano
  objValorPlano.AlterarValorPlano lngPKID, _
                                  strDescricao, _
                                  strValorTitular, _
                                  strValorDependente, _
                                  strStatus
  '
  Set objValorPlano = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

