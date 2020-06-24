VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubGrupoDespesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarTipoGrupo(strCodigoGrupo As String, _
                                    strCodigoSubGrupo As String) As String
  On Error GoTo trata
  Dim objSubGrupoDespesa As datSisContas.clsSubGrupoDespesa
  Dim objRs As ADODB.Recordset
  SelecionarTipoGrupo = ""
  Set objSubGrupoDespesa = New datSisContas.clsSubGrupoDespesa
  
  'Selecionar Tipo do Grupo Despesa
  Set objRs = objSubGrupoDespesa.SelecionarTipoGrupo(strCodigoGrupo, _
                                                     strCodigoSubGrupo)
  If Not objRs.EOF Then
    SelecionarTipoGrupo = objRs.Fields("TIPO").Value & ""
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set objSubGrupoDespesa = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Function SelecionarSubGrupoDespesa(lngSubGrupoDespesaId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objSubGrupoDespesa As datSisContas.clsSubGrupoDespesa
  Set objSubGrupoDespesa = New datSisContas.clsSubGrupoDespesa
  
  'Listar Sub Grupo Despesa
  Set SelecionarSubGrupoDespesa = objSubGrupoDespesa.SelecionarSubGrupoDespesa(lngSubGrupoDespesaId)
  '
  Set objSubGrupoDespesa = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirSubGrupoDespesa(lngGRUPODESPESAID As Long, _
                                  strCodigo As String, _
                                  strDescricao As String)
  
  On Error GoTo trata
  Dim objSubGrupoDespesa As datSisContas.clsSubGrupoDespesa
  Set objSubGrupoDespesa = New datSisContas.clsSubGrupoDespesa
  
  'Inserir Sub Grupo Despesa
  objSubGrupoDespesa.IncluirSubGrupoDespesa lngGRUPODESPESAID, _
                                            strCodigo, _
                                            strDescricao
  '
  Set objSubGrupoDespesa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarSubGrupoDespesa(lngSubGrupoDespesaId As Long, _
                                  strCodigo As String, _
                                  strDescricao As String)
  On Error GoTo trata
  Dim objSubGrupoDespesa As datSisContas.clsSubGrupoDespesa
  Set objSubGrupoDespesa = New datSisContas.clsSubGrupoDespesa
  
  'Inserir Sub Grupo Despesa
  objSubGrupoDespesa.AlterarSubGrupoDespesa lngSubGrupoDespesaId, _
                                            strCodigo, _
                                            strDescricao
  '
  Set objSubGrupoDespesa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub
Public Sub ExcluirSubGrupoDespesa(lngSubGrupoDespesaId As Long)
  On Error GoTo trata
  Dim objSubGrupoDespesa As datSisContas.clsSubGrupoDespesa
  Set objSubGrupoDespesa = New datSisContas.clsSubGrupoDespesa
  
  'Excluir Sub Grupo Despesa
  objSubGrupoDespesa.ExcluirSubGrupoDespesa lngSubGrupoDespesaId

  '
  Set objSubGrupoDespesa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


