VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFuncionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarFuncionarioPeloPkid(lngFUNCIONARIOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objFuncionario As datSisMed.clsFuncionario
  Set objFuncionario = New datSisMed.clsFuncionario
  
  'Listar Funcionario
  Set SelecionarFuncionarioPeloPkid = objFuncionario.SelecionarFuncionarioPeloPkid(lngFUNCIONARIOID)
  '
  Set objFuncionario = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function



Public Sub InserirFuncionario(lngPRONTUARIOID As Long, _
                              strUsuario As String, _
                              strNivel As String, _
                              strSenha As String, _
                              strExcluido As String)
  On Error GoTo trata
  Dim objFuncionario     As datSisMed.clsFuncionario
  '
  Set objFuncionario = New datSisMed.clsFuncionario
  'Inserir Funcionario
  objFuncionario.InserirFuncionario lngPRONTUARIOID, _
                              strUsuario, _
                              strNivel, _
                              strSenha, _
                              strExcluido
  '
  Set objFuncionario = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarFuncionario(lngPRONTUARIOID As Long, _
                              strUsuario As String, _
                              strNivel As String, _
                              strSenha As String, _
                              strExcluido As String)
  On Error GoTo trata
  Dim objFuncionario     As datSisMed.clsFuncionario
  '
  Set objFuncionario = New datSisMed.clsFuncionario
  'Atualizar Funcionario
  objFuncionario.AlterarFuncionario lngPRONTUARIOID, _
                              strUsuario, _
                              strNivel, _
                              strSenha, _
                              strExcluido
  '
  Set objFuncionario = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

