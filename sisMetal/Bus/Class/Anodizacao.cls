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
  Dim objAnodizacao   As datSisMetal.clsAnodizacao
  '
  Set objAnodizacao = New datSisMetal.clsAnodizacao
  'Inserir Anodizacao
  objAnodizacao.InserirAnodizacaoItem lngOSID, _
                                      lngITEM_OSID, _
                                      lngLINHAID, _
                                      strQuantidade, _
                                      strQuantidadeBRA, _
                                      strQuantidadeBRI, _
                                      strQuantidadeBRO, _
                                      strQuantidadeNAT
  '
  Set objAnodizacao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

