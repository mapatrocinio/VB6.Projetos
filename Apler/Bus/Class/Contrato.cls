VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarContratoPelaDescricao(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objContrato As datApler.clsContrato
  Set objContrato = New datApler.clsContrato
  
  'Listar Contrato
  Set SelecionarContratoPelaDescricao = objContrato.SelecionarContratoPelaDescricao(strDescricao)
  '
  Set objContrato = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarContratoPeloPkid(lngCONTRATOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objContrato As datApler.clsContrato
  Set objContrato = New datApler.clsContrato
  
  'Listar Contrato
  Set SelecionarContratoPeloPkid = objContrato.SelecionarContratoPeloPkid(lngCONTRATOID)
  '
  Set objContrato = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirContrato(lngCONTRATOID As Long)
  On Error GoTo trata
  Dim objContrato As datApler.clsContrato
  Set objContrato = New datApler.clsContrato
  
  'Excluir Contrato
  objContrato.ExcluirContrato lngCONTRATOID
  '
  Set objContrato = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirContrato(strDescricao As String, _
                        strObservacao As String, _
                        lngParceiroId As Long)
  On Error GoTo trata
  Dim objContrato     As datApler.clsContrato
  '
  Set objContrato = New datApler.clsContrato
  'Inserir Contrato
  objContrato.InserirContrato strDescricao, _
                        strObservacao, _
                        lngParceiroId
  '
  Set objContrato = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarContrato(lngPKID As Long, _
                        strDescricao As String, _
                        strObservacao As String, _
                        strStatus As String)

  On Error GoTo trata
  Dim objContrato     As datApler.clsContrato
  '
  Set objContrato = New datApler.clsContrato
  'Atualizar Contrato
  objContrato.AlterarContrato lngPKID, _
                        strDescricao, _
                        strObservacao, _
                        strStatus
  '
  Set objContrato = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

