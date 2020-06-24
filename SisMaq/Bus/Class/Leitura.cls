VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLeitura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarLeituraPeloPkid(lngLEITURAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objLeitura As datSisMaq.clsLeitura
  Set objLeitura = New datSisMaq.clsLeitura
  
  'Listar Leitura
  Set SelecionarLeituraPeloPkid = objLeitura.SelecionarLeituraPeloPkid(lngLEITURAID)
  '
  Set objLeitura = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirLeitura(lngLEITURAID As Long)
  On Error GoTo trata
  Dim objLeitura As datSisMaq.clsLeitura
  Set objLeitura = New datSisMaq.clsLeitura
  
  'Excluir Leitura
  objLeitura.ExcluirLeitura lngLEITURAID
  '
  Set objLeitura = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirLeitura(lngLEITURAID As Long, _
                          lngLEITURISTAID As Long, _
                          strData As String)
  On Error GoTo trata
  Dim objLeitura     As datSisMaq.clsLeitura
  '
  Set objLeitura = New datSisMaq.clsLeitura
  'Inserir Leitura
  objLeitura.InserirLeitura lngLEITURAID, _
                            lngLEITURISTAID, _
                            strData
  '
  Set objLeitura = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

