VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServicoFim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Public Function SelecionarServicoFimPeloPkid(lngSERVICOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objServicoFim As datElite.clsServicoFim
  Set objServicoFim = New datElite.clsServicoFim
  
  'Listar ServicoFim
  Set SelecionarServicoFimPeloPkid = objServicoFim.SelecionarServicoFimPeloPkid(lngSERVICOID)
  '
  Set objServicoFim = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Sub AlterarServicoFim(lngPKID As Long, _
                          strDataHora As String, _
                          strObservacao As String, _
                          strStatus As String)

  On Error GoTo trata
  Dim objServicoFim     As datElite.clsServicoFim
  '
  Set objServicoFim = New datElite.clsServicoFim
  'Atualizar ServicoFim
  objServicoFim.AlterarServicoFim lngPKID, _
                                  strDataHora, _
                                  strObservacao, _
                                  strStatus
  '
  Set objServicoFim = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

