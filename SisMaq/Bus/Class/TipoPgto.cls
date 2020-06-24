VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoPgto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub ExcluirTipoPgto(lngTipoPgtoId As Long)
  On Error GoTo trata
  Dim objTipoPgto As datSisMaq.clsTipoPgto
  Set objTipoPgto = New datSisMaq.clsTipoPgto
  
  'Excluir Tipo de Pagamento
  objTipoPgto.ExcluirTipoPgto lngTipoPgtoId

  '
  Set objTipoPgto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarTipoPgto(lngTipoPgtoId As Long, _
                           strTipoPgto As String)
  On Error GoTo trata
  Dim objTipoPgto As datSisMaq.clsTipoPgto
  Set objTipoPgto = New datSisMaq.clsTipoPgto
  
  'Inserir Tipo de Pagamento
  objTipoPgto.AlterarTipoPgto lngTipoPgtoId, _
                              strTipoPgto

  '
  Set objTipoPgto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub IncluirTipoPgto(strTipoPgto As String)
  
  On Error GoTo trata
  Dim objTipoPgto As datSisMaq.clsTipoPgto
  Set objTipoPgto = New datSisMaq.clsTipoPgto
  
  'Inserir Tipo de pagamento
  objTipoPgto.IncluirTipoPgto strTipoPgto

  '
  Set objTipoPgto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Function SelecionarTipoPgto(lngTipoPgtoId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objTipoPgto As datSisMaq.clsTipoPgto
  Set objTipoPgto = New datSisMaq.clsTipoPgto
  
  'Listar Tipo de pagamento
  Set SelecionarTipoPgto = objTipoPgto.SelecionarTipoPgto(lngTipoPgtoId)
  '
  Set objTipoPgto = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

