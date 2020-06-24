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
  Dim objSaldo As datSisContas.clsSaldo
  Set objSaldo = New datSisContas.clsSaldo
  
  'Listar Saldo
  Set SelecionarSaldo = objSaldo.SelecionarSaldo(lngSALDOID)
  '
  Set objSaldo = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirSaldo(strPercentual As String, _
                        strDescricao As String)
  
  On Error GoTo trata
  Dim objSaldo As datSisContas.clsSaldo
  Set objSaldo = New datSisContas.clsSaldo
  
  'Inserir Saldo
  objSaldo.IncluirSaldo strPercentual, _
                        strDescricao
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
  Dim objSaldo As datSisContas.clsSaldo
  Set objSaldo = New datSisContas.clsSaldo
  
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
  Dim objSaldo As datSisContas.clsSaldo
  Set objSaldo = New datSisContas.clsSaldo
  
  'Excluir Saldo
  objSaldo.ExcluirSaldo lngSALDOID

  '
  Set objSaldo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


