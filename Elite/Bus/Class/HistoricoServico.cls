VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistoricoServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Public Function SelecionarHistoricoServicoPeloPkid(lngHISTORICOSERVICOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objHistoricoServico As datElite.clsHistoricoServico
  Set objHistoricoServico = New datElite.clsHistoricoServico
  
  'Listar HistoricoServico
  Set SelecionarHistoricoServicoPeloPkid = objHistoricoServico.SelecionarHistoricoServicoPeloPkid(lngHISTORICOSERVICOID)
  '
  Set objHistoricoServico = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirHistoricoServico(lngHISTORICOSERVICOID As Long)
  On Error GoTo trata
  Dim objHistoricoServico As datElite.clsHistoricoServico
  Set objHistoricoServico = New datElite.clsHistoricoServico
  
  'Excluir HistoricoServico
  objHistoricoServico.ExcluirHistoricoServico lngHISTORICOSERVICOID
  '
  Set objHistoricoServico = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirHistoricoServico(lngPKID As Long, _
                                   strDataHora As String, _
                                   lngPACOTESERVICOID As Long, _
                                   strObservacao As String)
  On Error GoTo trata
  Dim objHistoricoServico     As datElite.clsHistoricoServico
  Dim strStatus As String
  '
  'Status Inicial
  strStatus = "I"
  Set objHistoricoServico = New datElite.clsHistoricoServico
  'Inserir HistoricoServico
  objHistoricoServico.InserirHistoricoServico lngPKID, _
                                              strDataHora, _
                                              lngPACOTESERVICOID, _
                                              strObservacao
  '
  Set objHistoricoServico = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarHistoricoServico(lngPKID As Long, _
                                   strDataHora As String, _
                                   lngPACOTESERVICOID As Long, _
                                   strObservacao As String)

  On Error GoTo trata
  Dim objHistoricoServico     As datElite.clsHistoricoServico
  '
  Set objHistoricoServico = New datElite.clsHistoricoServico
  'Atualizar HistoricoServico
  objHistoricoServico.AlterarHistoricoServico lngPKID, _
                                              strDataHora, _
                                              lngPACOTESERVICOID, _
                                              strObservacao
  '
  Set objHistoricoServico = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

