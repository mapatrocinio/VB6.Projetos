VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAssociadoDependente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarDependentePeloPkid(lngDependenteId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objDependente As datApler.clsAssociadoDependente
  Set objDependente = New datApler.clsAssociadoDependente
  
  'Listar Dependente
  Set SelecionarDependentePeloPkid = objDependente.SelecionarDependentePeloPkid(lngDependenteId)
  '
  Set objDependente = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function



Public Sub InserirDependente(lngPKID As Long, _
                             lngGRAUPARENTESCOID As Long, _
                             lngTITULARASSOCIADOID As Long, _
                             strMatricula As String)
  On Error GoTo trata
  Dim objDependente     As datApler.clsAssociadoDependente
  '
  Set objDependente = New datApler.clsAssociadoDependente
  'Inserir Dependente
  objDependente.InserirDependente lngPKID, _
                                  lngGRAUPARENTESCOID, _
                                  lngTITULARASSOCIADOID, _
                                  strMatricula
  '
  Set objDependente = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarDependente(lngPKID As Long, _
                             lngGRAUPARENTESCOID As Long, _
                             strMatricula As String)
  On Error GoTo trata
  Dim objDependente     As datApler.clsAssociadoDependente
  '
  Set objDependente = New datApler.clsAssociadoDependente
  'Atualizar Dependente
  objDependente.AlterarDependente lngPKID, _
                                  lngGRAUPARENTESCOID, _
                                  strMatricula
  '
  Set objDependente = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

