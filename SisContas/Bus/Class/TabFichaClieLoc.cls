VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTabFichaClieLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarTabFichaClieLoc(lngLocacaoId As Long, _
                                          strTipoCliente As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objTabFichaClieLoc As datSisContas.clsTabFichaClieLoc
  Set objTabFichaClieLoc = New datSisContas.clsTabFichaClieLoc
  
  'Tab FichaClie Loc
  Set SelecionarTabFichaClieLoc = objTabFichaClieLoc.SelecionarTabFichaClieLoc(lngLocacaoId, _
                                                                               strTipoCliente)
  '
  Set objTabFichaClieLoc = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Sub ExcluirTabFichaClieLoc(lngTabFichaClieLocId As Long)

  On Error GoTo trata
  Dim objTabFichaClieLoc As datSisContas.clsTabFichaClieLoc
  Set objTabFichaClieLoc = New datSisContas.clsTabFichaClieLoc
  
  'Excluir TabFichaClieLoc
  objTabFichaClieLoc.ExcluirTabFichaClieLoc lngTabFichaClieLocId
  '
  Set objTabFichaClieLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarTabFichaClieLoc(lngTabFichaClieLocId As Long, _
                                  lngFichaClienteId As Long)

  On Error GoTo trata
  Dim objTabFichaClieLoc As datSisContas.clsTabFichaClieLoc
  Set objTabFichaClieLoc = New datSisContas.clsTabFichaClieLoc
  
  'Alterar TabFichaClieLoc
  objTabFichaClieLoc.AlterarTabFichaClieLoc lngTabFichaClieLocId, _
                                            lngFichaClienteId
  '
  Set objTabFichaClieLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub IncluirTabFichaClieLoc(lngLocacoId As Long, _
                                  lngFichaClienteId As Long, _
                                  strTipoCliente As String)

  On Error GoTo trata
  Dim objTabFichaClieLoc As datSisContas.clsTabFichaClieLoc
  Set objTabFichaClieLoc = New datSisContas.clsTabFichaClieLoc
  
  'Inserir TabFichaClieLoc
  objTabFichaClieLoc.IncluirTabFichaClieLoc lngLocacoId, _
                                            lngFichaClienteId, _
                                            strTipoCliente
  '
  Set objTabFichaClieLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


