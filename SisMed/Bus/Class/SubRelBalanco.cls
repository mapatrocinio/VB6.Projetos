VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRelBalanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub SelecionarSaldoBalanco(ByRef curSaldoAnterior As Currency, _
                                  ByRef datDataSaldoAnterior As Date, _
                                  ByRef curReceita As Currency, _
                                  ByRef curPrestador As Currency, _
                                  ByRef curDespesa As Currency, _
                                  ByVal strDataInicial As String, _
                                  ByVal strDataFinal As String)
  On Error GoTo trata
  Dim objRelBalanco As datSisMed.clsRelBalanco
  Set objRelBalanco = New datSisMed.clsRelBalanco
  
  'Listar Sub Grupo Despesa
  objRelBalanco.SelecionarSaldoBalanco curSaldoAnterior, _
                                       datDataSaldoAnterior, _
                                       curReceita, _
                                       curPrestador, _
                                       curDespesa, _
                                       strDataInicial, _
                                       strDataFinal
  '
  Set objRelBalanco = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Function SelecionarRelBalanco(lngRelBalancoId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objRelBalanco As datSisMed.clsRelBalanco
  Set objRelBalanco = New datSisMed.clsRelBalanco
  
  'Listar Sub Grupo Despesa
  Set SelecionarRelBalanco = objRelBalanco.SelecionarRelBalanco(lngRelBalancoId)
  '
  Set objRelBalanco = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirRelBalanco(strData As String, _
                             strSaldo As String)
  
  On Error GoTo trata
  Dim objRelBalanco As datSisMed.clsRelBalanco
  Set objRelBalanco = New datSisMed.clsRelBalanco
  
  'Inserir Sub Grupo Despesa
  objRelBalanco.IncluirRelBalanco strData, _
                                  strSaldo
  '
  Set objRelBalanco = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarRelBalanco(lngRelBalancoId As Long, _
                             strData As String, _
                             strSaldo As String)
  On Error GoTo trata
  Dim objRelBalanco As datSisMed.clsRelBalanco
  Set objRelBalanco = New datSisMed.clsRelBalanco
  
  'Inserir Sub Grupo Despesa
  objRelBalanco.AlterarRelBalanco lngRelBalancoId, _
                                  strData, _
                                  strSaldo
  '
  Set objRelBalanco = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub
Public Sub ExcluirRelBalanco(lngRelBalancoId As Long)
  On Error GoTo trata
  Dim objRelBalanco As datSisMed.clsRelBalanco
  Set objRelBalanco = New datSisMed.clsRelBalanco
  
  'Excluir Sub Grupo Despesa
  objRelBalanco.ExcluirRelBalanco lngRelBalancoId

  '
  Set objRelBalanco = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


