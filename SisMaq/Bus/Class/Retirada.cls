VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRetirada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarRetiradaPeloPkid(lngRETIRADAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objRetirada As datSisMaq.clsRetirada
  Set objRetirada = New datSisMaq.clsRetirada
  
  'Listar Retirada
  Set SelecionarRetiradaPeloPkid = objRetirada.SelecionarRetiradaPeloPkid(lngRETIRADAID)
  '
  Set objRetirada = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirRetirada(lngRETIRADAID As Long)
  On Error GoTo trata
  Dim objRetirada As datSisMaq.clsRetirada
  Set objRetirada = New datSisMaq.clsRetirada
  
  'Excluir Retirada
  objRetirada.ExcluirRetirada lngRETIRADAID
  '
  Set objRetirada = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirRetirada(lngTURNOID As Long, _
                          strValor As String, _
                          strStatus As String, _
                          strData As String, _
                          lngGERENTEID As Long)
  On Error GoTo trata
  Dim objRetirada     As datSisMaq.clsRetirada
  '
  Set objRetirada = New datSisMaq.clsRetirada
  'Inserir Retirada
  objRetirada.InserirRetirada lngTURNOID, _
                              strValor, _
                              strStatus, _
                              strData, _
                              lngGERENTEID
  '
  Set objRetirada = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarRetirada(lngPKID As Long, _
                          strValor As String)

  On Error GoTo trata
  Dim objRetirada     As datSisMaq.clsRetirada
  '
  Set objRetirada = New datSisMaq.clsRetirada
  'Atualizar Retirada
  objRetirada.AlterarRetirada lngPKID, _
                            strValor
  '
  Set objRetirada = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

