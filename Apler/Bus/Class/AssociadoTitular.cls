VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAssociadoTitular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarTitularPeloPkid(lngTitularId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objTitular As datApler.clsAssociadoTitular
  Set objTitular = New datApler.clsAssociadoTitular
  
  'Listar Titular
  Set SelecionarTitularPeloPkid = objTitular.SelecionarTitularPeloPkid(lngTitularId)
  '
  Set objTitular = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub InserirTitular(lngPKID As Long, _
                          lngORIGEMID As Long, _
                          lngCaptadorId As Long, _
                          lngEMPRESAID As Long, _
                          strMatricula As String, _
                          strNumeroContrato As String, _
                          strDataIniContrato As String, _
                          strDataFimContrato As String, _
                          strEndRuaCom As String, _
                          strEndNumeroCom As String, _
                          strEndComplCom As String, _
                          strEndCepCom As String, _
                          strEndBairroCom As String, _
                          strEndCidadeCom As String, _
                          strEndEstadoCom As String, _
                          strTelefoneCom1 As String, _
                          strTelefoneCom2 As String, _
                          strEndRuaCob As String, _
                          strEndNumeroCob As String, _
                          strEndComplCob As String, _
                          strEndCepCob As String, _
                          strEndBairroCob As String, _
                          strEndCidadeCob As String, _
                          strEndEstadoCob As String)
  On Error GoTo trata
  Dim objTitular     As datApler.clsAssociadoTitular
  '
  Set objTitular = New datApler.clsAssociadoTitular
  'Inserir Titular
  objTitular.InserirTitular lngPKID, _
                            lngORIGEMID, _
                            lngCaptadorId, _
                            lngEMPRESAID, _
                            strMatricula, _
                            strNumeroContrato, _
                            strDataIniContrato, _
                            strDataFimContrato, _
                            strEndRuaCom, _
                            strEndNumeroCom, _
                            strEndComplCom, _
                            strEndCepCom, _
                            strEndBairroCom, _
                            strEndCidadeCom, _
                            strEndEstadoCom, _
                            strTelefoneCom1, _
                            strTelefoneCom2, _
                            strEndRuaCob, _
                            strEndNumeroCob, _
                            strEndComplCob, _
                            strEndCepCob, _
                            strEndBairroCob, _
                            strEndCidadeCob, _
                            strEndEstadoCob
  '
  Set objTitular = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarTitular(lngPKID As Long, _
                          lngORIGEMID As Long, _
                          lngCaptadorId As Long, _
                          lngEMPRESAID As Long, _
                          strMatricula As String, _
                          strNumeroContrato As String, _
                          strDataIniContrato As String, _
                          strDataFimContrato As String, _
                          strEndRuaCom As String, _
                          strEndNumeroCom As String, _
                          strEndComplCom As String, _
                          strEndCepCom As String, _
                          strEndBairroCom As String, _
                          strEndCidadeCom As String, _
                          strEndEstadoCom As String, _
                          strTelefoneCom1 As String, _
                          strTelefoneCom2 As String, _
                          strEndRuaCob As String, _
                          strEndNumeroCob As String, _
                          strEndComplCob As String, _
                          strEndCepCob As String, _
                          strEndBairroCob As String, _
                          strEndCidadeCob As String, _
                          strEndEstadoCob As String)
  On Error GoTo trata
  Dim objTitular     As datApler.clsAssociadoTitular
  '
  Set objTitular = New datApler.clsAssociadoTitular
  'Atualizar Titular
  objTitular.AlterarTitular lngPKID, _
                            lngORIGEMID, _
                            lngCaptadorId, _
                            lngEMPRESAID, _
                            strMatricula, _
                            strNumeroContrato, _
                            strDataIniContrato, _
                            strDataFimContrato, _
                            strEndRuaCom, _
                            strEndNumeroCom, _
                            strEndComplCom, _
                            strEndCepCom, _
                            strEndBairroCom, _
                            strEndCidadeCom, _
                            strEndEstadoCom, _
                            strTelefoneCom1, _
                            strTelefoneCom2, _
                            strEndRuaCob, _
                            strEndNumeroCob, _
                            strEndComplCob, _
                            strEndCepCob, _
                            strEndBairroCob, _
                            strEndCidadeCob, _
                            strEndEstadoCob
  '
  Set objTitular = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

