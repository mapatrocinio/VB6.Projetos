VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReceita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarReceita(lngRECEITAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objReceita As datSisContas.clsReceita
  Set objReceita = New datSisContas.clsReceita
  
  'Listar Receita
  Set SelecionarReceita = objReceita.SelecionarReceita(lngRECEITAID)
  '
  Set objReceita = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirReceita(strDtVencimento As String, _
                          strValorReceber As String, _
                          lngLIVROID As Long, _
                          strNumeroCheque As String, _
                          strTipo As String, _
                          strDtRecebimento As String, _
                          strDescricao As String, _
                          strValorRecebido As String, _
                          lngSubGrupoDespesaId As Long, _
                          lngFormaPgtoId As Long, _
                          strUsuario As String, _
                          ByRef lngSEQUENCIALEXTERNO As Long, _
                          lngEMPRESACTRTOID As Long, _
                          lngParceiroId As Long)
  
  On Error GoTo trata
  Dim objReceita As datSisContas.clsReceita
  Set objReceita = New datSisContas.clsReceita
  
  'Inserir Receita
  objReceita.IncluirReceita strDtVencimento, _
                            strValorReceber, _
                            lngLIVROID, _
                            strNumeroCheque, _
                            strTipo, _
                            strDtRecebimento, _
                            strDescricao, _
                            strValorRecebido, _
                            lngSubGrupoDespesaId, _
                            lngFormaPgtoId, _
                            strUsuario, _
                            lngSEQUENCIALEXTERNO, _
                            lngEMPRESACTRTOID, _
                            lngParceiroId

  '
  Set objReceita = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarReceita(strDtRecebimento As String, _
                          strDtVencimento As String, _
                          strValorReceber As String, _
                          lngLIVROID As Long, _
                          strNumeroCheque As String, _
                          lngRECEITAID As Long, _
                          strDescricao As String, _
                          strValorRecebido As String, _
                          lngSubGrupoDespesaId As Long, _
                          lngFormaPgtoId As Long, _
                          strUsuarioAlterouDescricao As String, _
                          lngEMPRESACTRTOID As Long)
  On Error GoTo trata
  Dim objReceita As datSisContas.clsReceita
  Set objReceita = New datSisContas.clsReceita
  
  'Alterar Receita
  objReceita.AlterarReceita strDtRecebimento, _
                            strDtVencimento, _
                            strValorReceber, _
                            lngLIVROID, _
                            strNumeroCheque, _
                            lngRECEITAID, _
                            strDescricao, _
                            strValorRecebido, _
                            lngSubGrupoDespesaId, _
                            lngFormaPgtoId, _
                            strUsuarioAlterouDescricao, _
                            lngEMPRESACTRTOID

  '
  Set objReceita = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub
Public Sub ExcluirReceita(lngRECEITAID As Long)
  On Error GoTo trata
  Dim objReceita As datSisContas.clsReceita
  Set objReceita = New datSisContas.clsReceita
  
  'Excluir Receita
  objReceita.ExcluirReceita lngRECEITAID

  '
  Set objReceita = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


