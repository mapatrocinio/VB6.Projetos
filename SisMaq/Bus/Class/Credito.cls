VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarCreditoPeloPkid(lngCREDITOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objCredito As datSisMaq.clsCredito
  Set objCredito = New datSisMaq.clsCredito
  
  'Listar Credito
  Set SelecionarCreditoPeloPkid = objCredito.SelecionarCreditoPeloPkid(lngCREDITOID)
  '
  Set objCredito = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirCredito(lngCREDITOID As Long)
  On Error GoTo trata
  Dim objCredito As datSisMaq.clsCredito
  Set objCredito = New datSisMaq.clsCredito
  
  'Excluir Credito
  objCredito.ExcluirCredito lngCREDITOID
  '
  Set objCredito = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirCredito(lngMAQUINAID As Long, _
                         lngBOLETOARRECID As Long, _
                         strNumero As String, _
                         strMedicao As String, _
                         strValorPago As String, _
                         strCoeficiente As String, _
                         strData As String, _
                         lngTURNOID As Long, _
                         strCredito As String)
  On Error GoTo trata
  Dim objCredito     As datSisMaq.clsCredito
  '
  Set objCredito = New datSisMaq.clsCredito
  'Inserir Credito
  objCredito.InserirCredito lngMAQUINAID, _
                          lngBOLETOARRECID, _
                          strNumero, _
                          strMedicao, _
                          strValorPago, _
                          strCoeficiente, _
                          strData, _
                          lngTURNOID, _
                          strCredito
                          
                          
  '
  Set objCredito = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarCredito(lngPKID As Long, _
                         strMedicao As String, _
                         strValorPago As String)

  On Error GoTo trata
  Dim objCredito     As datSisMaq.clsCredito
  '
  Set objCredito = New datSisMaq.clsCredito
  'Atualizar Credito
  objCredito.AlterarCredito lngPKID, _
                          strMedicao, _
                          strValorPago
  '
  Set objCredito = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

