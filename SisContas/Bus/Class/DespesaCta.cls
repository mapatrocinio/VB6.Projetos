VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDespesaCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarDespesa(lngDESPESAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objDespesa As datSisContas.clsDespesaCta
  Set objDespesa = New datSisContas.clsDespesaCta
  
  'Listar Despesa
  Set SelecionarDespesa = objDespesa.SelecionarDespesa(lngDESPESAID)
  '
  Set objDespesa = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirDespesa(strDtVencimento As String, _
                          strValorPagar As String, _
                          lngLIVROID As Long, _
                          strNumeroCheque As String, _
                          strTipo As String, _
                          strDtPagamento As String, _
                          strDescricao As String, _
                          strValorPago As String, _
                          strVale As String, _
                          lngSubGrupoDespesaId As Long, _
                          lngFormaPgtoId As Long, _
                          strUsuario As String, _
                          ByRef lngSEQUENCIALEXTERNO As Long, _
                          lngParceiroId As Long)
  
  On Error GoTo trata
  Dim objDespesa As datSisContas.clsDespesaCta
  Set objDespesa = New datSisContas.clsDespesaCta
  
  'Inserir Despesa
  objDespesa.IncluirDespesa strDtVencimento, _
                            strValorPagar, _
                            lngLIVROID, _
                            strNumeroCheque, _
                            strTipo, _
                            strDtPagamento, _
                            strDescricao, _
                            strValorPago, _
                            strVale, _
                            lngSubGrupoDespesaId, _
                            lngFormaPgtoId, _
                            strUsuario, _
                            lngSEQUENCIALEXTERNO, _
                            lngParceiroId

  '
  Set objDespesa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarDespesa(strDtPagamento As String, _
                          strDtVencimento As String, _
                          strValorPagar As String, _
                          lngLIVROID As Long, _
                          strNumeroCheque As String, _
                          lngDESPESAID As Long, _
                          strDescricao As String, _
                          strValorPago As String, _
                          strVale As String, _
                          lngSubGrupoDespesaId As Long, _
                          lngFormaPgtoId As Long, _
                          strUsuarioAlterouDescricao As String)
  On Error GoTo trata
  Dim objDespesa As datSisContas.clsDespesaCta
  Set objDespesa = New datSisContas.clsDespesaCta
  
  'Alterar Despesa
  objDespesa.AlterarDespesa strDtPagamento, _
                            strDtVencimento, _
                            strValorPagar, _
                            lngLIVROID, _
                            strNumeroCheque, _
                            lngDESPESAID, _
                            strDescricao, _
                            strValorPago, _
                            strVale, _
                            lngSubGrupoDespesaId, _
                            lngFormaPgtoId, _
                            strUsuarioAlterouDescricao

  '
  Set objDespesa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub
Public Sub ExcluirDespesa(lngDESPESAID As Long)
  On Error GoTo trata
  Dim objDespesa As datSisContas.clsDespesaCta
  Set objDespesa = New datSisContas.clsDespesaCta
  
  'Excluir Despesa
  objDespesa.ExcluirDespesa lngDESPESAID

  '
  Set objDespesa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


