VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTabResp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirTabResp(ByVal intTipo As Integer, _
                          ByVal lngTabRespId As Long)
  Dim objTabResp      As datSisContas.clsTabResp
  '
  On Error GoTo trata
  Set objTabResp = New datSisContas.clsTabResp
  '
  objTabResp.ExcluirTabResp intTipo, _
                            lngTabRespId
  '
  Set objTabResp = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirTabResp(intTipo As Integer, _
                          lngResLocId As Long, _
                          lngResponsabilidadeId As Long, _
                          strTipo As String, _
                          lngGrupoCardapioId As String, _
                          lngQtdResponsabilidade As String, _
                          strCardapioId As String, _
                          strTurno As String, _
                          strTipoDiretoriaId As String)

  On Error GoTo trata
  Dim objTabResp As datSisContas.clsTabResp
  '
  Set objTabResp = New datSisContas.clsTabResp
  
  'Inserir TabResp
  objTabResp.InserirTabResp intTipo, _
                            lngResLocId, _
                            lngResponsabilidadeId, _
                            strTipo, _
                            lngGrupoCardapioId, _
                            lngQtdResponsabilidade, _
                            strCardapioId, _
                            strTurno, _
                            strTipoDiretoriaId
  '
  Set objTabResp = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarTabResp(intTipo As Integer, _
                          lngTabRespId As Long, _
                          lngResponsabilidadeId As Long, _
                          strTipo As String, _
                          lngGrupoCardapioId As String, _
                          lngQtdResponsabilidade As String, _
                          strCardapioId As String, _
                          strTurno As String, _
                          strTipoDiretoriaId As String)
                         
  Dim objTabResp As datSisContas.clsTabResp
  Set objTabResp = New datSisContas.clsTabResp
  
  'Alterar TabResp
  objTabResp.AlterarTabResp intTipo, _
                            lngTabRespId, _
                            lngResponsabilidadeId, _
                            strTipo, _
                            lngGrupoCardapioId, _
                            lngQtdResponsabilidade, _
                            strCardapioId, _
                            strTurno, _
                            strTipoDiretoriaId
  '
  Set objTabResp = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function SelecionarTabResp(intTipo As Integer, _
                                  lngTabRespId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objTabResp As datSisContas.clsTabResp
  Set objTabResp = New datSisContas.clsTabResp
  
  'Selecionar TabResp
  Set SelecionarTabResp = objTabResp.SelecionarTabResp(intTipo, _
                                                       lngTabRespId)
  '
  Set objTabResp = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
