VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMovimentacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarMovimentacao(lngMOVIMENTACAOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objMovimentacao As datSisContas.clsMovimentacao
  Set objMovimentacao = New datSisContas.clsMovimentacao
  
  'Listar Movimentacao
  Set SelecionarMovimentacao = objMovimentacao.SelecionarMovimentacao(lngMOVIMENTACAOID)
  '
  Set objMovimentacao = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirMovimentacao(strStatus As String, _
                               strData As String, _
                               strDocumento As String, _
                               strCONTADEBITOID As String, _
                               strCONTACREDITOID As String, _
                               strValor As String, _
                               strDescricao As String)
  
  On Error GoTo trata
  Dim objMovimentacao As datSisContas.clsMovimentacao
  Set objMovimentacao = New datSisContas.clsMovimentacao
  
  'Inserir Movimentacao
  objMovimentacao.IncluirMovimentacao strStatus, _
                                      strData, _
                                      strDocumento, _
                                      strCONTADEBITOID, _
                                      strCONTACREDITOID, _
                                      strValor, _
                                      strDescricao
  '
  Set objMovimentacao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarMovimentacao(strData As String, _
                               strDocumento As String, _
                               strCONTADEBITOID As String, _
                               strCONTACREDITOID As String, _
                               strValor As String, _
                               strDescricao As String, _
                               lngMOVIMENTACAOID As Long)
  On Error GoTo trata
  Dim objMovimentacao As datSisContas.clsMovimentacao
  Set objMovimentacao = New datSisContas.clsMovimentacao
  
  'Alterar Movimentacao
  objMovimentacao.AlterarMovimentacao strData, _
                                      strDocumento, _
                                      strCONTADEBITOID, _
                                      strCONTACREDITOID, _
                                      strValor, _
                                      strDescricao, _
                                      lngMOVIMENTACAOID
  '
  Set objMovimentacao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub
Public Sub ExcluirMovimentacao(lngMOVIMENTACAOID As Long)
  On Error GoTo trata
  Dim objMovimentacao As datSisContas.clsMovimentacao
  Set objMovimentacao = New datSisContas.clsMovimentacao
  
  'Excluir Movimentacao
  objMovimentacao.ExcluirMovimentacao lngMOVIMENTACAOID

  '
  Set objMovimentacao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


