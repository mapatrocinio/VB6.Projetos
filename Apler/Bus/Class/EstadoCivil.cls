VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEstadoCivil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ListarEstadoCivilPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objEstadoCivil As datApler.clsEstadoCivil
  Set objEstadoCivil = New datApler.clsEstadoCivil
  
  'Listar EstadoCivil pela descrição
  Set ListarEstadoCivilPelaDesc = objEstadoCivil.ListarEstadoCivilPelaDesc(strDescricao)
  '
  Set objEstadoCivil = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub InserirEstadoCivil(strDescricao As String)

  On Error GoTo trata
  Dim objEstadoCivil As datApler.clsEstadoCivil
  '
  Set objEstadoCivil = New datApler.clsEstadoCivil
  
  'Inserir EstadoCivil
  objEstadoCivil.InserirEstadoCivil strDescricao
  '
  Set objEstadoCivil = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarEstadoCivil(lngESTADOCIVILID As Long, _
                              strDescricao As String)
                         
  On Error GoTo trata
  Dim objEstadoCivil As datApler.clsEstadoCivil
  Set objEstadoCivil = New datApler.clsEstadoCivil
  
  'Inserir EstadoCivil
  objEstadoCivil.AlterarEstadoCivil lngESTADOCIVILID, _
                                    strDescricao

  '
  Set objEstadoCivil = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarEstadoCivil(lngESTADOCIVILID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objEstadoCivil As datApler.clsEstadoCivil
  Set objEstadoCivil = New datApler.clsEstadoCivil
  
  'Listar EstadoCivil
  Set ListarEstadoCivil = objEstadoCivil.ListarEstadoCivil(lngESTADOCIVILID)
  '
  Set objEstadoCivil = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirEstadoCivil(lngESTADOCIVILID As Long)

  On Error GoTo trata
  Dim objEstadoCivil As datApler.clsEstadoCivil
  Set objEstadoCivil = New datApler.clsEstadoCivil
  
  'Excluir EstadoCivil
  objEstadoCivil.ExlcuirEstadoCivil lngESTADOCIVILID
  '
  Set objEstadoCivil = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

