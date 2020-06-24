VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Public Function SelecionarPacoteServicoPeloPkid(lngPACOTESERVICOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objServico As datElite.clsServico
  Set objServico = New datElite.clsServico
  
  'Listar Servico
  Set SelecionarPacoteServicoPeloPkid = objServico.SelecionarPacoteServicoPeloPkid(lngPACOTESERVICOID)
  '
  Set objServico = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarServicoPeloPkid(lngSERVICOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objServico As datElite.clsServico
  Set objServico = New datElite.clsServico
  
  'Listar Servico
  Set SelecionarServicoPeloPkid = objServico.SelecionarServicoPeloPkid(lngSERVICOID)
  '
  Set objServico = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirServico(lngSERVICOID As Long)
  On Error GoTo trata
  Dim objServico As datElite.clsServico
  Set objServico = New datElite.clsServico
  
  'Excluir Servico
  objServico.ExcluirServico lngSERVICOID
  '
  Set objServico = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirServico(lngPKID As Long, _
                          strDataHora As String, _
                          lngAgenciaCNPJId As Long, _
                          lngORIGEMID As Long, _
                          lngDESTINOID As Long, _
                          strSolicitante As String, _
                          strPassageiro As String, _
                          strReserva As String, _
                          strQtdPassageiro As String, _
                          strQtdCrianca As String, _
                          strVoo As String, _
                          strTerminal As String, _
                          strValor As String, _
                          strObservacao As String)
  On Error GoTo trata
  Dim objServico     As datElite.clsServico
  Dim strStatus As String
  '
  'Status Inicial
  strStatus = "I"
  Set objServico = New datElite.clsServico
  'Inserir Servico
  objServico.InserirServico lngPKID, _
                            strDataHora, _
                            lngAgenciaCNPJId, _
                            lngORIGEMID, _
                            lngDESTINOID, _
                            strSolicitante, _
                            strPassageiro, _
                            strReserva, _
                            strQtdPassageiro, _
                            strQtdCrianca, _
                            strVoo, _
                            strTerminal, _
                            strValor, _
                            strObservacao, _
                            strStatus
  '
  Set objServico = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarVeiculo(lngPACOTESERVICOID As Long, _
                          lngVEICULOID As Long)

  On Error GoTo trata
  Dim objServico     As datElite.clsServico
  '
  Set objServico = New datElite.clsServico
  'Atualizar Servico
  objServico.AlterarVeiculo lngPACOTESERVICOID, _
                            lngVEICULOID
  '
  Set objServico = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
                          
Public Sub AlterarServico(lngPKID As Long, _
                          strDataHora As String, _
                          lngAgenciaCNPJId As Long, _
                          lngORIGEMID As Long, _
                          lngDESTINOID As Long, _
                          strSolicitante As String, _
                          strPassageiro As String, _
                          strReserva As String, _
                          strQtdPassageiro As String, _
                          strQtdCrianca As String, _
                          strVoo As String, _
                          strTerminal As String, _
                          strValor As String, _
                          strObservacao As String, _
                          strStatus As String)

  On Error GoTo trata
  Dim objServico     As datElite.clsServico
  '
  Set objServico = New datElite.clsServico
  'Atualizar Servico
  objServico.AlterarServico lngPKID, _
                            strDataHora, _
                            lngAgenciaCNPJId, _
                            lngORIGEMID, _
                            lngDESTINOID, _
                            strSolicitante, _
                            strPassageiro, _
                            strReserva, _
                            strQtdPassageiro, _
                            strQtdCrianca, _
                            strVoo, _
                            strTerminal, _
                            strValor, _
                            strObservacao, _
                            strStatus
  '
  Set objServico = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub AlterarStatusServico(lngPKID As Long, _
                                strStatus As String)

  On Error GoTo trata
  Dim objServico     As datElite.clsServico
  '
  Set objServico = New datElite.clsServico
  'Atualizar Servico
  objServico.AlterarStatusServico lngPKID, _
                                  strStatus
  '
  Set objServico = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub DesativarServicoDoPacote(lngPACOTESERVICOID As Long)
  On Error GoTo trata
  '
  Dim objServico As datElite.clsServico
  '
  Set objServico = New datElite.clsServico
  '
  objServico.DesativarServicoDoPacote lngPACOTESERVICOID
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub DesassociarServicoAoPacote(lngPACOTESERVICOID As Long)
  On Error GoTo trata
  '
  Dim objServico As datElite.clsServico
  '
  Set objServico = New datElite.clsServico
  '
  objServico.DesassociarServicoAoPacote lngPACOTESERVICOID
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AssociarServicoAoPacote(lngPACOTEID As Long, _
                                   lngSERVICOID As Long, _
                                   lngVEICULOID As Long)
  On Error GoTo trata
  '
  Dim objServico    As datElite.clsServico
  Dim strStatus     As String
  '
  'Status INICIAL Ativo
  strStatus = "A"
  Set objServico = New datElite.clsServico
  '
  objServico.AssociarServicoAoPacote lngPACOTEID, _
                                     lngSERVICOID, _
                                     lngVEICULOID, _
                                     strStatus
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub



