VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBeneficiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarBeneficiarioPeloPkid(lngBeneficiarioId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objBeneficiario As datApler.clsBeneficiario
  Set objBeneficiario = New datApler.clsBeneficiario
  
  'Listar Beneficiario
  Set SelecionarBeneficiarioPeloPkid = objBeneficiario.SelecionarBeneficiarioPeloPkid(lngBeneficiarioId)
  '
  Set objBeneficiario = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirBeneficiario(lngBeneficiarioId As Long)
  On Error GoTo trata
  Dim objBeneficiario As datApler.clsBeneficiario
  Set objBeneficiario = New datApler.clsBeneficiario
  
  'Excluir Beneficiario
  objBeneficiario.ExcluirBeneficiario lngBeneficiarioId
  '
  Set objBeneficiario = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirBeneficiario(lngTABCONVASSOCID As Long, _
                               lngGRAUPARENTESCOID As Long, _
                               strNome As String, _
                               strCpf As String, _
                               strIdentidade As String, _
                               strDataNascimento As String, _
                               strSexo As String)
  On Error GoTo trata
  Dim objBeneficiario     As datApler.clsBeneficiario
  '
  Set objBeneficiario = New datApler.clsBeneficiario
  'Inserir Beneficiario
  objBeneficiario.InserirBeneficiario lngTABCONVASSOCID, _
                                      lngGRAUPARENTESCOID, _
                                      strNome, _
                                      strCpf, _
                                      strIdentidade, _
                                      strDataNascimento, _
                                      strSexo
  '
  Set objBeneficiario = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarBeneficiario(lngPKID As Long, _
                               lngGRAUPARENTESCOID As Long, _
                               strNome As String, _
                               strCpf As String, _
                               strIdentidade As String, _
                               strDataNascimento As String, _
                               strSexo As String)

  On Error GoTo trata
  Dim objBeneficiario     As datApler.clsBeneficiario
  '
  Set objBeneficiario = New datApler.clsBeneficiario
  'Atualizar Beneficiario
  objBeneficiario.AlterarBeneficiario lngPKID, _
                                      lngGRAUPARENTESCOID, _
                                      strNome, _
                                      strCpf, _
                                      strIdentidade, _
                                      strDataNascimento, _
                                      strSexo
  '
  Set objBeneficiario = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

