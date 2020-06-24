VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLeituraFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarLeituraFechaPeloPkid(lngLEITURAFECHAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objLeituraFecha As datSisMaq.clsLeituraFecha
  Set objLeituraFecha = New datSisMaq.clsLeituraFecha
  
  'Listar LeituraFecha
  Set SelecionarLeituraFechaPeloPkid = objLeituraFecha.SelecionarLeituraFechaPeloPkid(lngLEITURAFECHAID)
  '
  Set objLeituraFecha = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirLeituraFecha(lngLEITURAFECHAID As Long)
  On Error GoTo trata
  Dim objLeituraFecha As datSisMaq.clsLeituraFecha
  Set objLeituraFecha = New datSisMaq.clsLeituraFecha
  
  'Excluir LeituraFecha
  objLeituraFecha.ExcluirLeituraFecha lngLEITURAFECHAID
  '
  Set objLeituraFecha = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirLeituraFecha(lngLEITURAFECHAID As Long, _
                               lngCAIXAID As Long, _
                               lngPERIODOID As Long, _
                               strData As String, _
                               strStatus As String)
  On Error GoTo trata
  Dim objLeituraFecha     As datSisMaq.clsLeituraFecha
  '
  Set objLeituraFecha = New datSisMaq.clsLeituraFecha
  'Inserir LeituraFecha
  objLeituraFecha.InserirLeituraFecha lngLEITURAFECHAID, _
                                      lngCAIXAID, _
                                      lngPERIODOID, _
                                      strData, _
                                      strStatus
  '
  Set objLeituraFecha = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

