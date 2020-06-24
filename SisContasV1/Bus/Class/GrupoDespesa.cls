VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrupoDespesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarGrupoDespesa(lngGRUPODESPESAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objGrupoDespesa As datSisContas.clsGrupoDespesa
  Set objGrupoDespesa = New datSisContas.clsGrupoDespesa
  
  'Listar Grupo Despesa
  Set SelecionarGrupoDespesa = objGrupoDespesa.SelecionarGrupoDespesa(lngGRUPODESPESAID)
  '
  Set objGrupoDespesa = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirGrupoDespesa(strCodigo As String, _
                               strDescricao As String, _
                               strTipo As String)
  
  On Error GoTo trata
  Dim objGrupoDespesa As datSisContas.clsGrupoDespesa
  Set objGrupoDespesa = New datSisContas.clsGrupoDespesa
  
  'Inserir Grupo Despesa
  objGrupoDespesa.IncluirGrupoDespesa strCodigo, _
                                      strDescricao, _
                                      strTipo
  '
  Set objGrupoDespesa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarGrupoDespesa(lngGRUPODESPESAID As Long, _
                               strCodigo As String, _
                               strDescricao As String, _
                               strTipo As String)
  On Error GoTo trata
  Dim objGrupoDespesa As datSisContas.clsGrupoDespesa
  Set objGrupoDespesa = New datSisContas.clsGrupoDespesa
  
  'Inserir Grupo Despesa
  objGrupoDespesa.AlterarGrupoDespesa lngGRUPODESPESAID, _
                                      strCodigo, _
                                      strDescricao, _
                                      strTipo
  '
  Set objGrupoDespesa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub
Public Sub ExcluirGrupoDespesa(lngGRUPODESPESAID As Long)
  On Error GoTo trata
  Dim objGrupoDespesa As datSisContas.clsGrupoDespesa
  Set objGrupoDespesa = New datSisContas.clsGrupoDespesa
  
  'Excluir Grupo Despesa
  objGrupoDespesa.ExcluirGrupoDespesa lngGRUPODESPESAID

  '
  Set objGrupoDespesa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


