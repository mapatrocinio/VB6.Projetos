VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEspecialidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirEspecialidade(strEspecialidade As String)

  On Error GoTo trata
  Dim objEspecialidade As datSisMed.clsEspecialidade
  '
  Set objEspecialidade = New datSisMed.clsEspecialidade
  
  'Inserir Especialidade
  objEspecialidade.InserirEspecialidade strEspecialidade
  '
  Set objEspecialidade = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarEspecialidade(lngESPECIALIDADEID As Long, _
                         strEspecialidade As String)
                         
  On Error GoTo trata
  Dim objEspecialidade As datSisMed.clsEspecialidade
  Set objEspecialidade = New datSisMed.clsEspecialidade
  
  'Inserir Especialidade
  objEspecialidade.AlterarEspecialidade lngESPECIALIDADEID, _
                          strEspecialidade

  '
  Set objEspecialidade = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarEspecialidade(lngESPECIALIDADEID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objEspecialidade As datSisMed.clsEspecialidade
  Set objEspecialidade = New datSisMed.clsEspecialidade
  
  'Listar Especialidade
  Set ListarEspecialidade = objEspecialidade.ListarEspecialidade(lngESPECIALIDADEID)
  '
  Set objEspecialidade = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirEspecialidade(lngESPECIALIDADEID As Long)

  On Error GoTo trata
  Dim objEspecialidade As datSisMed.clsEspecialidade
  Set objEspecialidade = New datSisMed.clsEspecialidade
  
  'Excluir Especialidade
  objEspecialidade.ExlcuirEspecialidade lngESPECIALIDADEID
  '
  Set objEspecialidade = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

