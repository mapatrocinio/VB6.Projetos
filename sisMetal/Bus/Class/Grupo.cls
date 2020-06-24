VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function SelecionarGrupo(lngGRUPOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objGrupo As datSisMetal.clsGrupo
  Set objGrupo = New datSisMetal.clsGrupo
  
  'Listar Grupo
  Set SelecionarGrupo = objGrupo.SelecionarGrupo(lngGRUPOID)
  '
  Set objGrupo = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirGrupo(strNome As String, _
                        strReferencia As String)
  
  On Error GoTo trata
  Dim objGrupo As datSisMetal.clsGrupo
  Set objGrupo = New datSisMetal.clsGrupo
  
  'Inserir Grupo
  objGrupo.IncluirGrupo strNome, _
                        strReferencia
  '
  Set objGrupo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarGrupo(strNome As String, _
                        strReferencia As String, _
                        lngGRUPOID As Long)
  On Error GoTo trata
  Dim objGrupo As datSisMetal.clsGrupo
  Set objGrupo = New datSisMetal.clsGrupo
  
  'Alterar Grupo
  objGrupo.AlterarGrupo strNome, _
                        strReferencia, _
                        lngGRUPOID

  '
  Set objGrupo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub
Public Sub ExcluirGrupo(lngGRUPOID As Long)
  On Error GoTo trata
  Dim objGrupo As datSisMetal.clsGrupo
  Set objGrupo = New datSisMetal.clsGrupo
  
  'Excluir Grupo
  objGrupo.ExcluirGrupo lngGRUPOID

  '
  Set objGrupo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


