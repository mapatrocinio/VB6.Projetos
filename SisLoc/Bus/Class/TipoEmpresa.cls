VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ListarTipoEmpresaPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objTipoEmpresa As datSisLoc.clsTipoEmpresa
  Set objTipoEmpresa = New datSisLoc.clsTipoEmpresa
  
  'Listar TipoEmpresa pela descrição
  Set ListarTipoEmpresaPelaDesc = objTipoEmpresa.ListarTipoEmpresaPelaDesc(strDescricao)
  '
  Set objTipoEmpresa = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub InserirTipoEmpresa(strDescricao As String)

  On Error GoTo trata
  Dim objTipoEmpresa As datSisLoc.clsTipoEmpresa
  '
  Set objTipoEmpresa = New datSisLoc.clsTipoEmpresa
  
  'Inserir TipoEmpresa
  objTipoEmpresa.InserirTipoEmpresa strDescricao
  '
  Set objTipoEmpresa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarTipoEmpresa(lngTIPOEMPRESAID As Long, _
                               strDescricao As String)
                         
  On Error GoTo trata
  Dim objTipoEmpresa As datSisLoc.clsTipoEmpresa
  Set objTipoEmpresa = New datSisLoc.clsTipoEmpresa
  
  'Inserir TipoEmpresa
  objTipoEmpresa.AlterarTipoEmpresa lngTIPOEMPRESAID, _
                                      strDescricao

  '
  Set objTipoEmpresa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarTipoEmpresa(lngTIPOEMPRESAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objTipoEmpresa As datSisLoc.clsTipoEmpresa
  Set objTipoEmpresa = New datSisLoc.clsTipoEmpresa
  
  'Listar TipoEmpresa
  Set ListarTipoEmpresa = objTipoEmpresa.ListarTipoEmpresa(lngTIPOEMPRESAID)
  '
  Set objTipoEmpresa = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirTipoEmpresa(lngTIPOEMPRESAID As Long)

  On Error GoTo trata
  Dim objTipoEmpresa As datSisLoc.clsTipoEmpresa
  Set objTipoEmpresa = New datSisLoc.clsTipoEmpresa
  
  'Excluir TipoEmpresa
  objTipoEmpresa.ExlcuirTipoEmpresa lngTIPOEMPRESAID
  '
  Set objTipoEmpresa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

