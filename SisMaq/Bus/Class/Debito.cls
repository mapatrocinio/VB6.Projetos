VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarDebitoPeloPkid(lngDEBITOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objDebito As datSisMaq.clsDebito
  Set objDebito = New datSisMaq.clsDebito
  
  'Listar Debito
  Set SelecionarDebitoPeloPkid = objDebito.SelecionarDebitoPeloPkid(lngDEBITOID)
  '
  Set objDebito = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirDebito(lngDEBITOID As Long)
  On Error GoTo trata
  Dim objDebito As datSisMaq.clsDebito
  Set objDebito = New datSisMaq.clsDebito
  
  'Excluir Debito
  objDebito.ExcluirDebito lngDEBITOID
  '
  Set objDebito = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirDebito(lngMAQUINAID As Long, _
                         lngBOLETOATENDID As Long, _
                         strNumero As String, _
                         strMedicao As String, _
                         strValorPago As String, _
                         strCoeficiente As String, _
                         strData As String, _
                         lngTURNOID As Long, _
                         strCredito As String)
  On Error GoTo trata
  Dim objDebito     As datSisMaq.clsDebito
  '
  Set objDebito = New datSisMaq.clsDebito
  'Inserir Debito
  objDebito.InserirDebito lngMAQUINAID, _
                          lngBOLETOATENDID, _
                          strNumero, _
                          strMedicao, _
                          strValorPago, _
                          strCoeficiente, _
                          strData, _
                          lngTURNOID, _
                          strCredito
  '
  Set objDebito = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarDebito(lngPKID As Long, _
                         strMedicao As String, _
                         strValorPago As String)

  On Error GoTo trata
  Dim objDebito     As datSisMaq.clsDebito
  '
  Set objDebito = New datSisMaq.clsDebito
  'Atualizar Debito
  objDebito.AlterarDebito lngPKID, _
                          strMedicao, _
                          strValorPago
  '
  Set objDebito = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

