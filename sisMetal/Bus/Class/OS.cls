VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirOS(ByVal lngOSID As Long)
  On Error GoTo trata
  '
  Dim objOS       As datSisMetal.clsOS
  '
  Set objOS = New datSisMetal.clsOS
  '
  'Excluir OS
  objOS.ExcluirOS lngOSID
  '
  Set objOS = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirOS(lngOSID As Long, _
                     lngFORNECEDORID As Long, _
                     lngCORID As Long, _
                     lngFABRICAID As Long, _
                     strNumero As String, _
                     strNF As String, _
                     strPesoBruto As String, _
                     strPesoLiquido As String, _
                     strValorMetal As String, _
                     strValor As String, _
                     strOperador As String, _
                     strConferente As String)
  On Error GoTo trata
  Dim objOS           As datSisMetal.clsOS
  Dim strData         As String
  '
  Set objOS = New datSisMetal.clsOS
  '
  strData = Format(Now, "DD/MM/YYYY hh:mm")
  'Inserir OS
  objOS.InserirOS lngOSID, _
                  lngFORNECEDORID, _
                  lngCORID, _
                  lngFABRICAID, _
                  strNumero, _
                  strNF, _
                  strData, _
                  strPesoBruto, _
                  strPesoLiquido, _
                  strValorMetal, _
                  strValor, _
                  strOperador, _
                  strConferente
  '
  Set objOS = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarOS(lngOSID As Long, _
                     lngFORNECEDORID As Long, _
                     lngCORID As Long, _
                     lngFABRICAID As Long, _
                     strNumero As String, _
                     strNF As String, _
                     strPesoBruto As String, _
                     strPesoLiquido As String, _
                     strValorMetal As String, _
                     strValor As String, _
                     strOperador As String, _
                     strConferente As String)
  On Error GoTo trata
  Dim objOS As datSisMetal.clsOS
  '
  Set objOS = New datSisMetal.clsOS
  '
  'Alterar OS
  objOS.AlterarOS lngOSID, _
                  lngFORNECEDORID, _
                  lngCORID, _
                  lngFABRICAID, _
                  strNumero, _
                  strNF, _
                  strPesoBruto, _
                  strPesoLiquido, _
                  strValorMetal, _
                  strValor, _
                  strOperador, _
                  strConferente
  '
  Set objOS = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarOS(lngOSID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objOS As datSisMetal.clsOS
  Set objOS = New datSisMetal.clsOS

  'Listar OS
  Set ListarOS = objOS.ListarOS(lngOSID)
  '
  Set objOS = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
