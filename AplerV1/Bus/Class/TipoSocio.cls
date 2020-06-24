VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ListarTipoSocioPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objTipoSocio As datApler.clsTipoSocio
  Set objTipoSocio = New datApler.clsTipoSocio
  
  'Listar TipoSocio pela descrição
  Set ListarTipoSocioPelaDesc = objTipoSocio.ListarTipoSocioPelaDesc(strDescricao)
  '
  Set objTipoSocio = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub InserirTipoSocio(strDescricao As String, _
                            strTipo As String)

  On Error GoTo trata
  Dim objTipoSocio As datApler.clsTipoSocio
  '
  Set objTipoSocio = New datApler.clsTipoSocio
  
  'Inserir TipoSocio
  objTipoSocio.InserirTipoSocio strDescricao, _
                                strTipo
  '
  Set objTipoSocio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarTipoSocio(lngTIPOSOCIOID As Long, _
                            strDescricao As String, _
                            strTipo As String)
                         
  On Error GoTo trata
  Dim objTipoSocio As datApler.clsTipoSocio
  Set objTipoSocio = New datApler.clsTipoSocio
  
  'Inserir TipoSocio
  objTipoSocio.AlterarTipoSocio lngTIPOSOCIOID, _
                                strDescricao, _
                                strTipo

  '
  Set objTipoSocio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarTipoSocio(lngTIPOSOCIOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objTipoSocio As datApler.clsTipoSocio
  Set objTipoSocio = New datApler.clsTipoSocio
  
  'Listar TipoSocio
  Set ListarTipoSocio = objTipoSocio.ListarTipoSocio(lngTIPOSOCIOID)
  '
  Set objTipoSocio = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirTipoSocio(lngTIPOSOCIOID As Long)

  On Error GoTo trata
  Dim objTipoSocio As datApler.clsTipoSocio
  Set objTipoSocio = New datApler.clsTipoSocio
  
  'Excluir TipoSocio
  objTipoSocio.ExlcuirTipoSocio lngTIPOSOCIOID
  '
  Set objTipoSocio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

