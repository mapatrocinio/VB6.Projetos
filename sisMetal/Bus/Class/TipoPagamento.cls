VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirTipoPagamento(strDescricao As String)

  On Error GoTo trata
  Dim objTipoPagamento As datSisMetal.clsTipoPagamento
  '
  Set objTipoPagamento = New datSisMetal.clsTipoPagamento
  
  'Inserir TipoPagamento
  objTipoPagamento.InserirTipoPagamento strDescricao
  '
  Set objTipoPagamento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarTipoPagamento(lngTIPOPAGAMENTOID As Long, _
                                strDescricao As String)
                         
  On Error GoTo trata
  Dim objTipoPagamento As datSisMetal.clsTipoPagamento
  Set objTipoPagamento = New datSisMetal.clsTipoPagamento
  
  'Inserir TipoPagamento
  objTipoPagamento.AlterarTipoPagamento lngTIPOPAGAMENTOID, _
                                        strDescricao

  '
  Set objTipoPagamento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarTipoPagamento(lngTIPOPAGAMENTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objTipoPagamento As datSisMetal.clsTipoPagamento
  Set objTipoPagamento = New datSisMetal.clsTipoPagamento
  
  'Listar TipoPagamento
  Set ListarTipoPagamento = objTipoPagamento.ListarTipoPagamento(lngTIPOPAGAMENTOID)
  '
  Set objTipoPagamento = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirTipoPagamento(lngTIPOPAGAMENTOID As Long)

  On Error GoTo trata
  Dim objTipoPagamento As datSisMetal.clsTipoPagamento
  Set objTipoPagamento = New datSisMetal.clsTipoPagamento
  
  'Excluir TipoPagamento
  objTipoPagamento.ExlcuirTipoPagamento lngTIPOPAGAMENTOID
  '
  Set objTipoPagamento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

