VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTpConvenioVrPlano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub AssociarTpConvenioAoVrPlano(lngVALORPLANOID As Long, _
                                       lngTIPOCONVENIOID As Long)
  On Error GoTo trata
  '
  Dim strSql                  As String
  Dim objRs                   As ADODB.Recordset
  Dim objTpConvenioVrPlano    As datApler.clsTpConvenioVrPlano
  '
  Set objTpConvenioVrPlano = New datApler.clsTpConvenioVrPlano
  '
  objTpConvenioVrPlano.InserirTpConvenioVrPlano lngVALORPLANOID, _
                                                lngTIPOCONVENIOID
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub DesassociarTpConvenioDoVrPlano(lngVALORPLANOID As Long, _
                                          lngTIPOCONVENIOID As Long)
  On Error GoTo trata
  '
  Dim strSql                  As String
  Dim objRs                   As ADODB.Recordset
  Dim objTpConvenioVrPlano    As datApler.clsTpConvenioVrPlano
  '
  Set objTpConvenioVrPlano = New datApler.clsTpConvenioVrPlano
  '
  objTpConvenioVrPlano.ExcluirTpConvenioVrPlano lngVALORPLANOID, _
                                                lngTIPOCONVENIOID
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

