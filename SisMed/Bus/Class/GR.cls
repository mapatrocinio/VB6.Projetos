VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarGRPeloPkid(lngGRID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objGR As datSisMed.clsGR
  Set objGR = New datSisMed.clsGR
  
  'Listar GR
  Set SelecionarGRPeloPkid = objGR.SelecionarGRPeloPkid(lngGRID)
  '
  Set objGR = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CapturaPrestEspec(strCodigo As String, _
                                  strDiaDaSemana As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objGR As datSisMed.clsGR
  Set objGR = New datSisMed.clsGR
  
  'Capturar PrestEspec
  Set CapturaPrestEspec = objGR.CapturaPrestEspec(strCodigo, _
                                                  strDiaDaSemana)
  '
  Set objGR = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function CapturaProntuario(strNome As String, _
                                  strCPF As String, _
                                  strDtNascimento) As ADODB.Recordset
  On Error GoTo trata
  Dim objGR As datSisMed.clsGR
  Set objGR = New datSisMed.clsGR
  
  'Capturar Prontuario
  Set CapturaProntuario = objGR.CapturaProntuario(strNome, _
                                                  strCPF, _
                                                  strDtNascimento)
  '
  Set objGR = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function



Public Sub InserirGR(ByRef lngPKID As Long, _
                     lngPRONTUARIOID As Long, _
                     lngTurnoId As Long, _
                     lngTurnoLabId As String, _
                     lngAtendeId As Long, _
                     lngESPECIALIDADEID As Long, _
                     strSequencial As String, _
                     strSequencialSenha As String, _
                     strData As String, _
                     strStatus As String, _
                     strStatusImp As String, _
                     lngFUNCIONARIOID As Long, _
                     strCortesia As String, _
                     strUsuLib As String, _
                     strDescricao As String)
  On Error GoTo trata
  Dim objGR     As datSisMed.clsGR
  '
  Set objGR = New datSisMed.clsGR
  'Inserir GR
  objGR.InserirGR lngPKID, _
                  lngPRONTUARIOID, _
                  lngTurnoId, _
                  lngTurnoLabId, _
                  lngAtendeId, _
                  lngESPECIALIDADEID, _
                  strSequencial, _
                  strSequencialSenha, _
                  strData, _
                  strStatus, _
                  strStatusImp, _
                  lngFUNCIONARIOID, _
                  strCortesia, _
                  strUsuLib, _
                  strDescricao
  '
  Set objGR = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub



Public Sub AlterarGRExpirarAtend(lngPKID As Long)
  On Error GoTo trata
  Dim objGR     As datSisMed.clsGR
  Dim strDataExpiraAtend As String
  '
  Set objGR = New datSisMed.clsGR
  'Atualizar GR
  strDataExpiraAtend = Format(Now, "DD/MM/YYYY hh:mm")
  objGR.AlterarGRExpirarAtend lngPKID, _
                              strDataExpiraAtend
  'Altera status da GR
  objGR.AlterarStatusGR lngPKID, _
                        "E", _
                        ""
  Set objGR = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarGRLiberarAtend(lngPKID As Long)
  On Error GoTo trata
  Dim objGR     As datSisMed.clsGR
  Dim strDataLiberaAtend As String
  '
  Set objGR = New datSisMed.clsGR
  'Atualizar GR
  strDataLiberaAtend = Format(Now, "DD/MM/YYYY hh:mm")
  objGR.AlterarGRLiberarAtend lngPKID, _
                              strDataLiberaAtend
  'Altera status da GR
  objGR.AlterarStatusGR lngPKID, _
                        "L", _
                        ""
  Set objGR = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarGRNaoAtendida(lngPKID As Long)
  On Error GoTo trata
  Dim objGR     As datSisMed.clsGR
  Dim strDataNaoAtendimento As String
  '
  Set objGR = New datSisMed.clsGR
  'Atualizar GR
  strDataNaoAtendimento = Format(Now, "DD/MM/YYYY hh:mm")
  objGR.AlterarGRNaoAtendida lngPKID, _
                             strDataNaoAtendimento
  'Altera status da GR
  objGR.AlterarStatusGR lngPKID, _
                        "N", _
                        ""
  Set objGR = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub



Public Sub AlterarGR(lngPKID As Long, _
                     lngPRONTUARIOID As Long, _
                     lngAtendeId As Long, _
                     lngESPECIALIDADEID As Long, _
                     strDescricao As String)
  On Error GoTo trata
  Dim objGR     As datSisMed.clsGR
  '
  Set objGR = New datSisMed.clsGR
  'Atualizar GR
  objGR.AlterarGR lngPKID, _
                  lngPRONTUARIOID, _
                  lngAtendeId, _
                  lngESPECIALIDADEID, _
                  strDescricao
  '
  Set objGR = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function CapturaProcedimentoGeral(strProcedimento As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objGR As datSisMed.clsGR
  Set objGR = New datSisMed.clsGR
  
  'Capturar Procedimento
  Set CapturaProcedimentoGeral = objGR.CapturaProcedimentoGeral(strProcedimento)
  '
  Set objGR = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CapturaProcedimento(strProcedimento As String, _
                                    lngPRESTADORID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objGR As datSisMed.clsGR
  Set objGR = New datSisMed.clsGR
  
  'Capturar Procedimento
  Set CapturaProcedimento = objGR.CapturaProcedimento(strProcedimento, _
                                                      lngPRESTADORID)
  '
  Set objGR = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function



Public Sub InserirGRPROCEDIMENTO(lngGRID As Long, _
                                 lngPROCEDIMENTOID As Long, _
                                 strQuantidade As String, _
                                 strValor As String, _
                                 strValorCort As String)
  On Error GoTo trata
  Dim objGR     As datSisMed.clsGR
  '
  Set objGR = New datSisMed.clsGR
  'Inserir GRPROCEDIMENTO
  objGR.InserirGRPROCEDIMENTO lngGRID, _
                              lngPROCEDIMENTOID, _
                              strQuantidade, _
                              strValor, _
                              strValorCort
  '
  Set objGR = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub



Public Sub ExcluirGRPROCEDIMENTO(lngGRPROCEDIMENTOID As Long)
  On Error GoTo trata
  Dim objGR As datSisMed.clsGR
  Set objGR = New datSisMed.clsGR
  
  'Excluir GRPROCEDIMENTO
  objGR.ExcluirGRPROCEDIMENTO lngGRPROCEDIMENTOID
  '
  Set objGR = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function SelecionarGRTotalProc(lngGRID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objGR As datSisMed.clsGR
  Set objGR = New datSisMed.clsGR
  
  'Listar GR
  Set SelecionarGRTotalProc = objGR.SelecionarGRTotalProc(lngGRID)
  '
  Set objGR = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub AlterarTurnoGR(lngGRID As Long, _
                          lngTurnoId As Long, _
                          strFuncionarioId As String)
  On Error GoTo trata
  Dim objGR     As datSisMed.clsGR
  '
  Set objGR = New datSisMed.clsGR
  'Atualizar Turno GR
  objGR.AlterarTurnoGR lngGRID, _
                       lngTurnoId, _
                       strFuncionarioId
  '
  Set objGR = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarArquivoGR(lngGRID As Long, _
                            strStatus As String)
  On Error GoTo trata
  Dim objGR     As datSisMed.clsGR
  '
  Set objGR = New datSisMed.clsGR
  'Alterar Arquivo GR
  objGR.AlterarArquivoGR lngGRID, _
                         strStatus
  '
  Set objGR = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarStatusGR(lngGRID As Long, _
                           strStatus As String, _
                           strStatusImpressao As String, _
                           Optional lngTurnoId As Long, _
                           Optional gsNomeUsuLib As String)
  On Error GoTo trata
  Dim objGR     As datSisMed.clsGR
  '
  Set objGR = New datSisMed.clsGR
  'Atualizar Status GR
  objGR.AlterarStatusGR lngGRID, _
                        strStatus, _
                        strStatusImpressao, _
                        lngTurnoId, _
                        gsNomeUsuLib
  '
  Set objGR = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function CapturaGRTurnoCorrente(strTURNOID As String, _
                                       lngPRONTUARIOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objGR As datSisMed.clsGR
  Set objGR = New datSisMed.clsGR
  
  'Capturar GR do Turno
  Set CapturaGRTurnoCorrente = objGR.CapturaGRTurnoCorrente(strTURNOID, _
                                                            lngPRONTUARIOID)
  '
  Set objGR = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CapturaGRTurnoCorrenteTODOS(strTURNOID As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objGR As datSisMed.clsGR
  Set objGR = New datSisMed.clsGR
  
  'Capturar GR do Turno
  Set CapturaGRTurnoCorrenteTODOS = objGR.CapturaGRTurnoCorrenteTODOS(strTURNOID)
  '
  Set objGR = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CapturaGRTurnoCorrenteARQ(strTURNOID As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objGR As datSisMed.clsGR
  Set objGR = New datSisMed.clsGR
  
  'Capturar GR do Turno
  Set CapturaGRTurnoCorrenteARQ = objGR.CapturaGRTurnoCorrenteARQ(strTURNOID)
  '
  Set objGR = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CapturaGRTurnoCorrentePRE(lngPRESTADORID As Long, _
                                          strDataIniAtual As String, _
                                          strDataFimAtual As String, _
                                          strNivel As String, _
                                          strPrestador As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objGR As datSisMed.clsGR
  Set objGR = New datSisMed.clsGR
  
  'Capturar GR do Turno
  Set CapturaGRTurnoCorrentePRE = objGR.CapturaGRTurnoCorrentePRE(lngPRESTADORID, _
                                                                  strDataIniAtual, _
                                                                  strDataFimAtual, _
                                                                  strNivel, _
                                                                  strPrestador)
  '
  Set objGR = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function CapturaGRTurnoCorrenteFIN(strDataIniAtual As String, _
                                          strDataFimAtual As String, _
                                          intMaxDiasAtend As Integer) As ADODB.Recordset
  On Error GoTo trata
  Dim objGR As datSisMed.clsGR
  Set objGR = New datSisMed.clsGR
  
  'Capturar GR do Turno
  Set CapturaGRTurnoCorrenteFIN = objGR.CapturaGRTurnoCorrenteFIN(strDataIniAtual, _
                                                                  strDataFimAtual, _
                                                                  intMaxDiasAtend)
  '
  Set objGR = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


