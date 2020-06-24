VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaixaAtend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub FecharCaixaAtend(lngTURNOFECHAID As Long, _
                            strValorFecha As String, _
                            lngTURNOCAIXAATENDID As Long)

  On Error GoTo trata
  Dim objCaixaAtend     As datSisMaq.clsCaixaAtend
  '
  Set objCaixaAtend = New datSisMaq.clsCaixaAtend
  'Fechar CaixaAtend
  objCaixaAtend.FecharCaixaAtend lngTURNOFECHAID, _
                                 strValorFecha, _
                                 lngTURNOCAIXAATENDID
  '
  Set objCaixaAtend = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
                            
                            
Public Sub InserirCaixaAtend(ByRef lngPKID As Long, _
                             lngATENDENTEID As Long, _
                             lngTURNOENTRADAID As Long)
  On Error GoTo trata
  Dim objCaixaAtend     As datSisMaq.clsCaixaAtend
  '
  Set objCaixaAtend = New datSisMaq.clsCaixaAtend
  'Inserir CaixaAtend
  objCaixaAtend.InserirCaixaAtend lngPKID, _
                                  lngATENDENTEID, _
                                  lngTURNOENTRADAID
  '
  Set objCaixaAtend = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


