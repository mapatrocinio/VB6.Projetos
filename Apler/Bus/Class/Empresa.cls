VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ListarEmpresaPeloNome(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objEmpresa As datApler.clsEmpresa
  Set objEmpresa = New datApler.clsEmpresa
  
  'Listar Empresa pela descrição
  Set ListarEmpresaPeloNome = objEmpresa.ListarEmpresaPeloNome(strNome)
  '
  Set objEmpresa = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub InserirEmpresa(strNome As String, _
                          strCodigo As String)

  On Error GoTo trata
  Dim objEmpresa As datApler.clsEmpresa
  '
  Set objEmpresa = New datApler.clsEmpresa
  
  'Inserir Empresa
  objEmpresa.InserirEmpresa strNome, _
                            strCodigo
  '
  Set objEmpresa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarEmpresa(lngEMPRESAID As Long, _
                          strNome As String, _
                          strCodigo As String)
                         
  On Error GoTo trata
  Dim objEmpresa As datApler.clsEmpresa
  Set objEmpresa = New datApler.clsEmpresa
  
  'Inserir Empresa
  objEmpresa.AlterarEmpresa lngEMPRESAID, _
                            strNome, _
                            strCodigo

  '
  Set objEmpresa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarEmpresa(lngEMPRESAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objEmpresa As datApler.clsEmpresa
  Set objEmpresa = New datApler.clsEmpresa
  
  'Listar Empresa
  Set ListarEmpresa = objEmpresa.ListarEmpresa(lngEMPRESAID)
  '
  Set objEmpresa = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirEmpresa(lngEMPRESAID As Long)

  On Error GoTo trata
  Dim objEmpresa As datApler.clsEmpresa
  Set objEmpresa = New datApler.clsEmpresa
  
  'Excluir Empresa
  objEmpresa.ExlcuirEmpresa lngEMPRESAID
  '
  Set objEmpresa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

