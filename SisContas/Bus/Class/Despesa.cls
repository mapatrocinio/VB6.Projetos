VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDespesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarDespesa(lngDESPESAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objDespesa As datSisContas.clsDespesa
  Set objDespesa = New datSisContas.clsDespesa
  
  'Listar Despesa
  Set SelecionarDespesa = objDespesa.SelecionarDespesa(lngDESPESAID)
  '
  Set objDespesa = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function IncluirDespesa(strSequencial As String, _
                               strTipo As String, _
                               lngTurnoId As Long, _
                               strDtPagamento As String, _
                               strDescricao As String, _
                               strValorPago As String, _
                               strVale As String, _
                               lngSubGrupoDespesaId As Long, _
                               lngFormaPgtoId As Long, _
                               strUsuario As String, _
                               strUsuarioAutorizacao As String, _
                               lngFUNCIONARIOID As Long) As Long
  On Error GoTo trata
  Dim objDespesa As datSisContas.clsDespesa
  Set objDespesa = New datSisContas.clsDespesa
  
  'Inserir Despesa
  IncluirDespesa = objDespesa.IncluirDespesa(strSequencial, _
                                             strTipo, _
                                             lngTurnoId, _
                                             strDtPagamento, _
                                             strDescricao, _
                                             strValorPago, _
                                             strVale, _
                                             lngSubGrupoDespesaId, _
                                             lngFormaPgtoId, _
                                             strUsuario, _
                                             strUsuarioAutorizacao, _
                                             lngFUNCIONARIOID)

  '
  Set objDespesa = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub AlterarDespesa(lngDESPESAID As Long, _
                          strDescricao As String, _
                          strValorPago As String, _
                          strVale As String, _
                          lngSubGrupoDespesaId As Long, _
                          lngFormaPgtoId As Long, _
                          strUsuarioAutorizacao As String, _
                          lngFUNCIONARIOID As Long)
  On Error GoTo trata
  Dim objDespesa As datSisContas.clsDespesa
  Set objDespesa = New datSisContas.clsDespesa
  
  'Inserir Despesa
  objDespesa.AlterarDespesa lngDESPESAID, _
                            strDescricao, _
                            strValorPago, _
                            strVale, _
                            lngSubGrupoDespesaId, _
                            lngFormaPgtoId, _
                            strUsuarioAutorizacao, _
                            lngFUNCIONARIOID

  '
  Set objDespesa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub ExcluirDespesa(lngDESPESAID As Long)
  On Error GoTo trata
  Dim objDespesa  As datSisContas.clsDespesa
  Dim objCC       As busSisContas.clsContaCorrente
  Dim objRs       As ADODB.Recordset
  '
  'Excluir pagamentos da despesa
  Set objCC = New busSisContas.clsContaCorrente
  Set objRs = objCC.SelecionarPagamentos("DE", _
                                         lngDESPESAID)
  Do While Not objRs.EOF
    objCC.ExcluirContaCorrente "DE", _
                               objRs.Fields("PKID").Value
    objRs.MoveNext
  Loop
  objRs.Close
  Set objRs = Nothing
  Set objCC = Nothing
  '
  Set objDespesa = New datSisContas.clsDespesa
  'Excluir Despesa
  objDespesa.ExcluirDespesa lngDESPESAID

  '
  Set objDespesa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


