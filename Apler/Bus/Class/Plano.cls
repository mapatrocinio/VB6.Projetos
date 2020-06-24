VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarPlanoPelaDescricao(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objPlano As datApler.clsPlano
  Set objPlano = New datApler.clsPlano
  
  'Listar Plano
  Set SelecionarPlanoPelaDescricao = objPlano.SelecionarPlanoPelaDescricao(strDescricao)
  '
  Set objPlano = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarPlanoPeloPkid(lngPlanoId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objPlano As datApler.clsPlano
  Set objPlano = New datApler.clsPlano
  
  'Listar Plano
  Set SelecionarPlanoPeloPkid = objPlano.SelecionarPlanoPeloPkid(lngPlanoId)
  '
  Set objPlano = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirPlano(lngPlanoId As Long)
  On Error GoTo trata
  Dim objPlano As datApler.clsPlano
  Set objPlano = New datApler.clsPlano
  
  'Excluir Plano
  objPlano.ExcluirPlano lngPlanoId
  '
  Set objPlano = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirPlano(strDescricao As String, _
                        strObservacao As String)
  On Error GoTo trata
  Dim objPlano     As datApler.clsPlano
  '
  Set objPlano = New datApler.clsPlano
  'Inserir Plano
  objPlano.InserirPlano strDescricao, _
                        strObservacao
  '
  Set objPlano = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarPlano(lngPKID As Long, _
                        strDescricao As String, _
                        strObservacao As String, _
                        strStatus As String)

  On Error GoTo trata
  Dim objPlano     As datApler.clsPlano
  '
  Set objPlano = New datApler.clsPlano
  'Atualizar Plano
  objPlano.AlterarPlano lngPKID, _
                        strDescricao, _
                        strObservacao, _
                        strStatus
  '
  Set objPlano = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

