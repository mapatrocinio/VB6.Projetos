VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormaPgto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub ExcluirFormaPgto(lngFormaPgtoId As Long)
  On Error GoTo trata
  Dim objFormaPgto As datSisContas.clsFormaPgto
  Set objFormaPgto = New datSisContas.clsFormaPgto
  
  'Excluir Forma de Pagamento
  objFormaPgto.ExcluirFormaPgto lngFormaPgtoId

  '
  Set objFormaPgto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarFormaPgto(lngFormaPgtoId As Long, _
                            strFormaPgto As String)
  On Error GoTo trata
  Dim objFormaPgto As datSisContas.clsFormaPgto
  Set objFormaPgto = New datSisContas.clsFormaPgto
  
  'Inserir Forma de Pagamento
  objFormaPgto.AlterarFormaPgto lngFormaPgtoId, _
                                strFormaPgto

  '
  Set objFormaPgto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub IncluirFormaPgto(strFormaPgto As String)
  
  On Error GoTo trata
  Dim objFormaPgto As datSisContas.clsFormaPgto
  Set objFormaPgto = New datSisContas.clsFormaPgto
  
  'Inserir Forma de pagamento
  objFormaPgto.IncluirFormaPgto strFormaPgto

  '
  Set objFormaPgto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Function SelecionarFormaPgto(lngFormaPgtoId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objFormaPgto As datSisContas.clsFormaPgto
  Set objFormaPgto = New datSisContas.clsFormaPgto
  
  'Listar Forma de pagamento
  Set SelecionarFormaPgto = objFormaPgto.SelecionarFormaPgto(lngFormaPgtoId)
  '
  Set objFormaPgto = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

