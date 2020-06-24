VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaixaArrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub FecharCaixaArrec(lngTURNOFECHAID As Long, _
                            strValorFecha As String, _
                            lngTURNOCAIXAARRECID As Long)
  On Error GoTo trata
  Dim objCaixaArrec     As datSisMaq.clsCaixaArrec
  '
  Set objCaixaArrec = New datSisMaq.clsCaixaArrec
  'Inserir CaixaArrec
  objCaixaArrec.FecharCaixaArrec lngTURNOFECHAID, _
                                 strValorFecha, _
                                 lngTURNOCAIXAARRECID
  '
  Set objCaixaArrec = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirCaixaArrec(ByRef lngPKID As Long, _
                             lngARRECADADORID As Long, _
                             lngTURNOENTRADAID As Long)
  On Error GoTo trata
  Dim objCaixaArrec     As datSisMaq.clsCaixaArrec
  '
  Set objCaixaArrec = New datSisMaq.clsCaixaArrec
  'Inserir CaixaArrec
  objCaixaArrec.InserirCaixaArrec lngPKID, _
                                  lngARRECADADORID, _
                                  lngTURNOENTRADAID
  '
  Set objCaixaArrec = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


