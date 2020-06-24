VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCartaoDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirCartaoDebito(strNome As String, _
                               strNumero As String, _
                               strPercTxAdm As String, _
                               strConvenio As String)

  On Error GoTo trata
  Dim objCartaoDebito As datSisMed.clsCartaoDebito
  '
  Set objCartaoDebito = New datSisMed.clsCartaoDebito
  
  'Inserir CartaoDebito
  objCartaoDebito.InserirCartaoDebito strNome, _
                                      strNumero, _
                                      strPercTxAdm, _
                                      strConvenio
  '
  Set objCartaoDebito = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarCartaoDebito(lngCARTAODEBITOID As Long, _
                               strNome As String, _
                               strNumero As String, _
                               strPercTxAdm As String, _
                               strConvenio As String)
                         
  On Error GoTo trata
  Dim objCartaoDebito As datSisMed.clsCartaoDebito
  Set objCartaoDebito = New datSisMed.clsCartaoDebito
  
  'Inserir CartaoDebito
  objCartaoDebito.AlterarCartaoDebito lngCARTAODEBITOID, _
                                      strNome, _
                                      strNumero, _
                                      strPercTxAdm, _
                                      strConvenio

  '
  Set objCartaoDebito = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarCartaoDebito(lngCARTAODEBITOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objCartaoDebito As datSisMed.clsCartaoDebito
  Set objCartaoDebito = New datSisMed.clsCartaoDebito
  
  'Listar CartaoDebito
  Set ListarCartaoDebito = objCartaoDebito.ListarCartaoDebito(lngCARTAODEBITOID)
  '
  Set objCartaoDebito = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirCartaoDebito(lngCARTAODEBITOID As Long)

  On Error GoTo trata
  Dim objCartaoDebito As datSisMed.clsCartaoDebito
  Set objCartaoDebito = New datSisMed.clsCartaoDebito
  
  'Excluir CartaoDebito
  objCartaoDebito.ExlcuirCartaoDebito lngCARTAODEBITOID
  '
  Set objCartaoDebito = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

