VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGRPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'''Public Function SelecionarGRPagamentoPeloNumero(strNumero As String) As ADODB.Recordset
'''  On Error GoTo trata
'''  Dim objGRPagamento As datSisMed.clsGRPagamento
'''  Set objGRPagamento = New datSisMed.clsGRPagamento
'''
'''  'Listar GRPagamento
'''  Set SelecionarGRPagamentoPeloNumero = objGRPagamento.SelecionarGRPagamentoPeloNumero(strNumero)
'''  '
'''  Set objGRPagamento = Nothing
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Function

Public Function SelecionarGRPagamentoPeloPkid(lngGRPagamentoId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objGRPagamento As datSisMed.clsGRPagamento
  Set objGRPagamento = New datSisMed.clsGRPagamento
  
  'Listar GRPagamento
  Set SelecionarGRPagamentoPeloPkid = objGRPagamento.SelecionarGRPagamentoPeloPkid(lngGRPagamentoId)
  '
  Set objGRPagamento = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirGRPagamento(lngGRPagamentoId As Long)
  On Error GoTo trata
  Dim objGRPagamento As datSisMed.clsGRPagamento
  Set objGRPagamento = New datSisMed.clsGRPagamento
  
  'Excluir GRPagamento
  objGRPagamento.ExcluirGRPagamento lngGRPagamentoId
  '
  Set objGRPagamento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirGRPagamento(lngPKID As Long, _
                              lngPRESTADORID As Long, _
                              strDataInicio As String, _
                              strDataTermino As String, _
                              strStatus As String, _
                              strStatusImpressao As String, _
                              strNomeUsu As String)
  On Error GoTo trata
  Dim objGRPagamento     As datSisMed.clsGRPagamento
  '
  Set objGRPagamento = New datSisMed.clsGRPagamento
  'Inserir GRPagamento
  objGRPagamento.InserirGRPagamento lngPKID, _
                                    lngPRESTADORID, _
                                    strDataInicio, _
                                    strDataTermino, _
                                    strStatus, _
                                    strStatusImpressao, _
                                    strNomeUsu
  '
  Set objGRPagamento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarGRPagamento(lngPKID As Long, _
                              lngPRESTADORID As Long, _
                              strDataInicio As String, _
                              strDataTermino As String)

  On Error GoTo trata
  Dim objGRPagamento     As datSisMed.clsGRPagamento
  '
  Set objGRPagamento = New datSisMed.clsGRPagamento
  'Atualizar GRPagamento
  objGRPagamento.AlterarGRPagamento lngPKID, _
                                    lngPRESTADORID, _
                                    strDataInicio, _
                                    strDataTermino
  '
  Set objGRPagamento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub



