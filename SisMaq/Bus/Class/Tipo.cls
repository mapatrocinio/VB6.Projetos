VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirTipo(strTipo As String, _
                       strDescricao As String)

  On Error GoTo trata
  Dim objTipo As datSisMaq.clsTipo
  '
  Set objTipo = New datSisMaq.clsTipo
  
  'Inserir Tipo
  objTipo.InserirTipo strTipo, _
                      strDescricao
  '
  Set objTipo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarTipo(lngTIPOID As Long, _
                       strTipo As String, _
                       strDescricao As String)
                         
  On Error GoTo trata
  Dim objTipo As datSisMaq.clsTipo
  Set objTipo = New datSisMaq.clsTipo
  
  'Inserir Tipo
  objTipo.AlterarTipo lngTIPOID, _
                      strTipo, _
                      strDescricao

  '
  Set objTipo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarTipo(lngTIPOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objTipo As datSisMaq.clsTipo
  Set objTipo = New datSisMaq.clsTipo
  
  'Listar Tipo
  Set ListarTipo = objTipo.ListarTipo(lngTIPOID)
  '
  Set objTipo = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirTipo(lngTIPOID As Long)

  On Error GoTo trata
  Dim objTipo As datSisMaq.clsTipo
  Set objTipo = New datSisMaq.clsTipo
  
  'Excluir Tipo
  objTipo.ExlcuirTipo lngTIPOID
  '
  Set objTipo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

