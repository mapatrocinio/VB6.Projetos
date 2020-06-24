VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirUsuario(strUsuario As String, _
                          strNivel As String)

  On Error GoTo trata
  Dim objUsuario As datSisMed.clsUsuario
  '
  Set objUsuario = New datSisMed.clsUsuario
  
  'Inserir Usuario
  objUsuario.InserirUsuario strUsuario, _
                            strNivel
  '
  Set objUsuario = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarSenha(lngCONTROLEACESSOID As Long, _
                        strSenha As String)
                         
  On Error GoTo trata
  Dim objUsuario As datSisMed.clsUsuario
  Set objUsuario = New datSisMed.clsUsuario
  
  'Alterar Senha
  objUsuario.AlterarSenha lngCONTROLEACESSOID, _
                          strSenha
  '
  Set objUsuario = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarUsuario(lngCONTROLEACESSOID As Long, _
                          strUsuario As String, _
                          strNivel As String)
                         
  On Error GoTo trata
  Dim objUsuario As datSisMed.clsUsuario
  Set objUsuario = New datSisMed.clsUsuario
  
  'Inserir Usuario
  objUsuario.AlterarUsuario lngCONTROLEACESSOID, _
                            strUsuario, _
                            strNivel

  '
  Set objUsuario = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarUsuario(lngCONTROLEACESSOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objUsuario As datSisMed.clsUsuario
  Set objUsuario = New datSisMed.clsUsuario
  
  'Listar Usuario
  Set ListarUsuario = objUsuario.ListarUsuario(lngCONTROLEACESSOID)
  '
  Set objUsuario = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirUsuario(lngCONTROLEACESSOID As Long)

  On Error GoTo trata
  Dim objUsuario As datSisMed.clsUsuario
  Set objUsuario = New datSisMed.clsUsuario
  
  'Excluir Usuario
  objUsuario.ExlcuirUsuario lngCONTROLEACESSOID
  '
  Set objUsuario = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
