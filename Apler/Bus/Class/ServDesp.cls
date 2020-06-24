VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServDesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarServicoDespertador(lngSERVICODESPERTADORID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objServicoDespertador As datApler.clsServDesp
  Set objServicoDespertador = New datApler.clsServDesp
  
  'Selecionar
  Set SelecionarServicoDespertador = objServicoDespertador.SelecionarServicoDespertador(lngSERVICODESPERTADORID)
  '
  Set objServicoDespertador = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Sub IncluirServicoDespertador(lngLOCACAOID As Long, _
                                     strData As String, _
                                     blnStatus As Boolean)
  On Error GoTo trata
  Dim objServicoDespertador As datApler.clsServDesp
  Set objServicoDespertador = New datApler.clsServDesp
  
  'Inserir ServicoDespertador
  objServicoDespertador.IncluirServicoDespertador lngLOCACAOID, _
                                                  strData, _
                                                  blnStatus

  '
  Set objServicoDespertador = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub AlterarStatusServicoDespertador(lngSERVICODESPERTADORID As Long, _
                                           blnStatus As Boolean)
  On Error GoTo trata
  Dim objServicoDespertador As datApler.clsServDesp
  Set objServicoDespertador = New datApler.clsServDesp
  
  'Alterar Status ServicoDespertador
  objServicoDespertador.AlterarStatusServicoDespertador lngSERVICODESPERTADORID, _
                                                        blnStatus

  '
  Set objServicoDespertador = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarServicoDespertador(lngSERVICODESPERTADORID As Long, _
                                     strData As String)
  On Error GoTo trata
  Dim objServicoDespertador As datApler.clsServDesp
  Set objServicoDespertador = New datApler.clsServDesp
  
  'Inserir ServicoDespertador
  objServicoDespertador.AlterarServicoDespertador lngSERVICODESPERTADORID, _
                                                  strData

  '
  Set objServicoDespertador = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub ExcluirServicoDespertador(lngSERVICODESPERTADORID As Long)
  On Error GoTo trata
  Dim objServicoDespertador As datApler.clsServDesp
  Set objServicoDespertador = New datApler.clsServDesp
  
  'Excluir ServicoDespertador
  objServicoDespertador.ExcluirServicoDespertador lngSERVICODESPERTADORID

  '
  Set objServicoDespertador = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub



