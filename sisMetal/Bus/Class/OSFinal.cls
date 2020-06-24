VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOSFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirOSFinal(ByVal lngOSFinalID As Long)
  On Error GoTo trata
  '
  Dim objOSFinal       As datSisMetal.clsOSFinal
  '
  Set objOSFinal = New datSisMetal.clsOSFinal
  '
  'Excluir OSFinal
  objOSFinal.ExcluirOSFinal lngOSFinalID
  '
  Set objOSFinal = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirOSFinal(lngOSFinalID As Long, _
                          lngCORID As Long, _
                          strComplemento As String, _
                          lngOSID As Long, _
                          strPesoBruto As String, _
                          strPesoLiquido As String, _
                          strValorMetal As String, _
                          strValor As String, _
                          strValorFrete As String, _
                          strOperador As String, _
                          strConferente As String)
  On Error GoTo trata
  Dim objOSFinal           As datSisMetal.clsOSFinal
  Dim strData         As String
  '
  Set objOSFinal = New datSisMetal.clsOSFinal
  '
  strData = Format(Now, "DD/MM/YYYY hh:mm")
  'Inserir OSFinal
  objOSFinal.InserirOSFinal lngOSFinalID, _
                            lngCORID, _
                            strComplemento, _
                            lngOSID, _
                            strData, _
                            strPesoBruto, _
                            strPesoLiquido, _
                            strValorMetal, _
                            strValor, _
                            strValorFrete, _
                            strOperador, _
                            strConferente
  '
  Set objOSFinal = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarOSFinal(lngOSFinalID As Long, _
                          lngCORID As Long, _
                          strComplemento As String, _
                          strPesoBruto As String, _
                          strPesoLiquido As String, _
                          strValorMetal As String, _
                          strValor As String, _
                          strValorFrete As String, _
                          strOperador As String, _
                          strConferente As String)
  On Error GoTo trata
  Dim objOSFinal As datSisMetal.clsOSFinal
  '
  Set objOSFinal = New datSisMetal.clsOSFinal
  '
  'Alterar OSFinal
  objOSFinal.AlterarOSFinal lngOSFinalID, _
                            lngCORID, _
                            strComplemento, _
                            strPesoBruto, _
                            strPesoLiquido, _
                            strValorMetal, _
                            strValor, _
                            strValorFrete, _
                            strOperador, _
                            strConferente
  '
  Set objOSFinal = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarOSFinal(lngOSFinalID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objOSFinal As datSisMetal.clsOSFinal
  Set objOSFinal = New datSisMetal.clsOSFinal

  'Listar OSFinal
  Set ListarOSFinal = objOSFinal.ListarOSFinal(lngOSFinalID)
  '
  Set objOSFinal = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
