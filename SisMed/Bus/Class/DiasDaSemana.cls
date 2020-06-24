VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiasDaSemana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirDiaDaSemana(strCodigo As String, _
                              strDiaDaSemana As String)

  On Error GoTo trata
  Dim objDiasDaSemana As datSisMed.clsDiasDaSemana
  '
  Set objDiasDaSemana = New datSisMed.clsDiasDaSemana
  
  'Inserir Dia Da Semana
  objDiasDaSemana.InserirDiaDaSemana strCodigo, _
                                     strDiaDaSemana
  '
  Set objDiasDaSemana = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarDiaDaSemana(lngDIADASEMANAID As Long, _
                              strCodigo As String, _
                              strDiaDaSemana As String)
                         
  On Error GoTo trata
  Dim objDiasDaSemana As datSisMed.clsDiasDaSemana
  Set objDiasDaSemana = New datSisMed.clsDiasDaSemana
  
  'Inserir Dia da Semana
  objDiasDaSemana.AlterarDiaDaSemana lngDIADASEMANAID, _
                                     strCodigo, _
                                     strDiaDaSemana

  '
  Set objDiasDaSemana = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarDiaDaSemana(lngDIADASEMANAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objDiasDaSemana As datSisMed.clsDiasDaSemana
  Set objDiasDaSemana = New datSisMed.clsDiasDaSemana
  
  'Listar Dia Da Semana
  Set ListarDiaDaSemana = objDiasDaSemana.ListarDiaDaSemana(lngDIADASEMANAID)
  '
  Set objDiasDaSemana = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirDiaDaSemana(lngDIADASEMANAID As Long)

  On Error GoTo trata
  Dim objDiasDaSemana As datSisMed.clsDiasDaSemana
  Set objDiasDaSemana = New datSisMed.clsDiasDaSemana
  
  'Excluir Dia da Semana
  objDiasDaSemana.ExlcuirDiaDaSemana lngDIADASEMANAID
  '
  Set objDiasDaSemana = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

