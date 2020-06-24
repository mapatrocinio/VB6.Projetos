VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLeituraMaquina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarLeituraMaquinaPeloPkid(lngLEITURAMAQUINAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objLeituraMaquina As datSisMaq.clsLeituraMaquina
  Set objLeituraMaquina = New datSisMaq.clsLeituraMaquina
  
  'Listar LeituraMaquina
  Set SelecionarLeituraMaquinaPeloPkid = objLeituraMaquina.SelecionarLeituraMaquinaPeloPkid(lngLEITURAMAQUINAID)
  '
  Set objLeituraMaquina = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirLeituraMaquina(lngLEITURAMAQUINAID As Long)
  On Error GoTo trata
  Dim objLeituraMaquina As datSisMaq.clsLeituraMaquina
  Set objLeituraMaquina = New datSisMaq.clsLeituraMaquina
  
  'Excluir LeituraMaquina
  objLeituraMaquina.ExcluirLeituraMaquina lngLEITURAMAQUINAID
  '
  Set objLeituraMaquina = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirLeituraMaquina(lngLEITURAID As Long, _
                                 lngMAQUINAID As Long, _
                                 strMedicaoEntrada As String, _
                                 strMedicaoSaida As String, _
                                 strData As String)
  On Error GoTo trata
  Dim objLeituraMaquina     As datSisMaq.clsLeituraMaquina
  Dim lngPKID               As Long
  '
  Set objLeituraMaquina = New datSisMaq.clsLeituraMaquina
  'Inserir LeituraMaquina
  objLeituraMaquina.InserirLeituraMaquina lngPKID, _
                                          lngLEITURAID, _
                                          lngMAQUINAID, _
                                          strMedicaoEntrada, _
                                          strMedicaoSaida
  '
  '------------------------------------------------
  'Verificar inclusão de Erro na leitura
  If lngPKID <> 0 Then
    'Obter recordset com dados da medição anterior
    objLeituraMaquina.TratarLeituraMaquinaStatus lngPKID, _
                                                 lngMAQUINAID, _
                                                 strData, _
                                                 strMedicaoEntrada, _
                                                 strMedicaoSaida



  End If
  '------------------------------------------------
  Set objLeituraMaquina = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

