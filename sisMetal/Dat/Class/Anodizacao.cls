VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAnodizacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirAnodizacaoItem(lngOSID As Long, _
                                 lngITEM_OSID As Long, _
                                 lngLINHAID As Long, _
                                 strQuantidade As String, _
                                 strQuantidadeBRA As String, _
                                 strQuantidadeBRI As String, _
                                 strQuantidadeBRO As String, _
                                 strQuantidadeNAT As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "EXEC SP_INSERE_ITEM_ANOD " & Formata_Dados(lngOSID, tpDados_Longo) & _
      "," & Formata_Dados(lngITEM_OSID, tpDados_Longo) & _
      "," & Formata_Dados(lngLINHAID, tpDados_Longo) & _
      "," & Formata_Dados(strQuantidade, tpDados_Longo) & _
      "," & Formata_Dados(strQuantidadeBRA, tpDados_Longo) & _
      "," & Formata_Dados(strQuantidadeBRI, tpDados_Longo) & _
      "," & Formata_Dados(strQuantidadeBRO, tpDados_Longo) & _
      "," & Formata_Dados(strQuantidadeNAT, tpDados_Longo)

  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemOS.InserirAnodizacaoItem]", _
            Err.Description
End Sub

