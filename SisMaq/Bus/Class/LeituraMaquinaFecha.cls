VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLeituraMaquinaFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarLeituraMaquinaFechaPeloPkid(lngLEITURAMAQUINAFECHAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objLeituraMaquinaFecha As datSisMaq.clsLeituraMaquinaFecha
  Set objLeituraMaquinaFecha = New datSisMaq.clsLeituraMaquinaFecha
  
  'Listar LeituraMaquinaFecha
  Set SelecionarLeituraMaquinaFechaPeloPkid = objLeituraMaquinaFecha.SelecionarLeituraMaquinaFechaPeloPkid(lngLEITURAMAQUINAFECHAID)
  '
  Set objLeituraMaquinaFecha = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirLeituraMaquinaFecha(lngLEITURAMAQUINAFECHAID As Long)
  On Error GoTo trata
  Dim objLeituraMaquinaFecha As datSisMaq.clsLeituraMaquinaFecha
  Set objLeituraMaquinaFecha = New datSisMaq.clsLeituraMaquinaFecha
  
  'Excluir LeituraMaquinaFecha
  objLeituraMaquinaFecha.ExcluirLeituraMaquinaFecha lngLEITURAMAQUINAFECHAID
  '
  Set objLeituraMaquinaFecha = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirLeituraMaquinaFecha(lngLEITURAFECHAID As Long, _
                                      lngMAQUINAID As Long, _
                                      strMedicaoEntrada As String, _
                                      strMedicaoSaida As String, _
                                      strData As String, _
                                      strUsuario As String)
  On Error GoTo trata
  Dim objLeituraMaquinaFecha     As datSisMaq.clsLeituraMaquinaFecha
  Dim lngPKID               As Long
  '
  Set objLeituraMaquinaFecha = New datSisMaq.clsLeituraMaquinaFecha
  'Inserir LeituraMaquinaFecha
  objLeituraMaquinaFecha.InserirLeituraMaquinaFecha lngPKID, _
                                                    lngLEITURAFECHAID, _
                                                    lngMAQUINAID, _
                                                    strMedicaoEntrada, _
                                                    strMedicaoSaida, _
                                                    strUsuario
  Set objLeituraMaquinaFecha = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

