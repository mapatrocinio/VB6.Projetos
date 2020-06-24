VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAtende"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarAtendePeloPkid(lngAtendeId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objAtende As datSisMed.clsAtende
  Set objAtende = New datSisMed.clsAtende
  
  'Listar Atende
  Set SelecionarAtendePeloPkid = objAtende.SelecionarAtendePeloPkid(lngAtendeId)
  '
  Set objAtende = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirAtende(lngAtendeId As Long)
  On Error GoTo trata
  Dim objAtende As datSisMed.clsAtende
  Set objAtende = New datSisMed.clsAtende
  
  'Excluir Atende
  objAtende.ExcluirAtende lngAtendeId
  '
  Set objAtende = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirAtende(lngSalaId As Long, _
                         lngProntuarioId As Long, _
                         lngDiasDaSemanaId As Long, _
                         strDataInicio As String, _
                         strDataFim As String, _
                         strStatus As String)
  On Error GoTo trata
  Dim objAtende     As datSisMed.clsAtende
  '
  Set objAtende = New datSisMed.clsAtende
  'Inserir Atende
  objAtende.InserirAtende lngSalaId, _
                         lngProntuarioId, _
                         lngDiasDaSemanaId, _
                         strDataInicio, _
                         strDataFim, _
                         strStatus
  '
  Set objAtende = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarAtende(lngPKID As Long, _
                         lngSalaId As Long, _
                         lngProntuarioId As Long, _
                         lngDiasDaSemanaId As Long, _
                         strDataInicio As String, _
                         strDataFim As String, _
                         strStatus As String)

  On Error GoTo trata
  Dim objAtende     As datSisMed.clsAtende
  '
  Set objAtende = New datSisMed.clsAtende
  'Atualizar Atende
  objAtende.AlterarAtende lngPKID, _
                         lngSalaId, _
                         lngProntuarioId, _
                         lngDiasDaSemanaId, _
                         strDataInicio, _
                         strDataFim, _
                         strStatus
  '
  Set objAtende = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

