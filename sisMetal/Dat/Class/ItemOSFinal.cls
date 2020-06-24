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
  Dim strSql As String
  '
  strSql = "EXEC SP_INSERE_ITEM_OS_FINAL " & Formata_Dados(lngOSFINALID, tpDados_Longo) & _
      "," & Formata_Dados(lngOSID, tpDados_Longo) & _
      "," & Formata_Dados(lngCORID, tpDados_Longo) & _
      "," & Formata_Dados(lngANODIZACAO_ITEMID, tpDados_Longo) & _
      "," & Formata_Dados(lngITEM_OS_FINALID, tpDados_Longo) & _
      "," & Formata_Dados(lngLINHAID, tpDados_Longo) & _
      "," & Formata_Dados(strQuantidadeANOD, tpDados_Longo)

  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemOS.InserirItemOSFinalItem]", _
            Err.Description
End Sub

