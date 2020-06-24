VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItemOSFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirItemOSFinalItem(lngOSFINALID As Long, _
                                  lngOSID As Long, _
                                  lngCORID As Long, _
                                  lngANODIZACAO_ITEMID As Long, _
                                  lngITEM_OS_FINALID As Long, _
                                  lngLINHAID As Long, _
                                  strQuantidadeANOD As String)
  On Error GoTo trata
  Dim objItemOSFinal   As datSisMetal.clsItemOSFinal
  '
  Set objItemOSFinal = New datSisMetal.clsItemOSFinal
  'Inserir ItemOSFinal
  objItemOSFinal.InserirItemOSFinalItem lngOSFINALID, _
                                        lngOSID, _
                                        lngCORID, _
                                        lngANODIZACAO_ITEMID, _
                                        lngITEM_OS_FINALID, _
                                        lngLINHAID, _
                                        strQuantidadeANOD
  '
  Set objItemOSFinal = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

