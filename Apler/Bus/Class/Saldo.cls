VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarSaldo(lngSALDOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objSaldo As datApler.clsSaldo
  Set objSaldo = New datApler.clsSaldo
  
  'Listar Saldo
  Set SelecionarSaldo = objSaldo.SelecionarSaldo(lngSALDOID)
  '
  Set objSaldo = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirSaldo(strPercentual As String, _
                        strDescricao As String, _
                        lngParceiroId As Long)
  
  On Error GoTo trata
  Dim objSaldo As datApler.clsSaldo
  Set objSaldo = New datApler.clsSaldo
  
  'Inserir Saldo
  objSaldo.IncluirSaldo strPercentual, _
                        strDescricao, _
                        lngParceiroId
  '
  Set objSaldo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarSaldo(strPercentual As String, _
                        strDescricao As String, _
                        lngSALDOID As Long)
  On Error GoTo trata
  Dim objSaldo As datApler.clsSaldo
  Set objSaldo = New datApler.clsSaldo
  
  'Alterar Saldo
  objSaldo.AlterarSaldo strPercentual, _
                        strDescricao, _
                        lngSALDOID

  '
  Set objSaldo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub
Public Sub ExcluirSaldo(lngSALDOID As Long)
  On Error GoTo trata
  Dim objSaldo As datApler.clsSaldo
  Set objSaldo = New datApler.clsSaldo
  
  'Excluir Saldo
  objSaldo.ExcluirSaldo lngSALDOID

  '
  Set objSaldo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


