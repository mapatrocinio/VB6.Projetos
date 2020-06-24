VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEntregaDireta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirEntregaDireta(ByVal lngENTREGADIRETAID As Long)
  On Error GoTo trata
  '
  Dim objEntregaDireta       As datSisMetal.clsEntregaDireta
  Dim strData         As String
  '
  strData = Format(Now, "DD/MM/YYYY hh:mm")
  Set objEntregaDireta = New datSisMetal.clsEntregaDireta
  '
  'Excluir EntregaDireta
  objEntregaDireta.ExcluirEntregaDireta lngENTREGADIRETAID, _
                                        strData
  '
  '
  Set objEntregaDireta = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirEntregaDireta(lngENTREGADIRETAID As Long, _
                                lngFORNECEDORID As Long)
  On Error GoTo trata
  Dim objEntregaDireta       As datSisMetal.clsEntregaDireta
  Dim strAno          As String
  Dim strNumeroOS     As String
  Dim strData         As String
  '
  Set objEntregaDireta = New datSisMetal.clsEntregaDireta
  '
  strData = Format(Now, "DD/MM/YYYY hh:mm")
  strAno = Format(Now, "YYYY")
  'Capturar Sequencial do EntregaDireta
  objEntregaDireta.CapturarSequencialEntregaDireta strAno, _
                                                   lngFORNECEDORID, _
                                                   strNumeroOS
  'Inserir EntregaDireta
  objEntregaDireta.InserirEntregaDireta lngENTREGADIRETAID, _
                                        lngFORNECEDORID, _
                                        strAno, _
                                        strNumeroOS, _
                                        strData
  '
  Set objEntregaDireta = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarEntregaDireta(lngENTREGADIRETAID As Long, _
                                lngFORNECEDORID As Long)
  On Error GoTo trata
  Dim objEntregaDireta As datSisMetal.clsEntregaDireta
  '
  Set objEntregaDireta = New datSisMetal.clsEntregaDireta
  '
  'Alterar EntregaDireta
  objEntregaDireta.AlterarEntregaDireta lngENTREGADIRETAID, _
                          lngFORNECEDORID
  '
  Set objEntregaDireta = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarEntregaDireta(lngENTREGADIRETAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objEntregaDireta As datSisMetal.clsEntregaDireta
  Set objEntregaDireta = New datSisMetal.clsEntregaDireta

  'Listar EntregaDireta
  Set ListarEntregaDireta = objEntregaDireta.ListarEntregaDireta(lngENTREGADIRETAID)
  '
  Set objEntregaDireta = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
