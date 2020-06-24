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

Public Function SelecionarFuncionarioPeloPkid(lngFuncionarioId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objFuncionario As datSisMaq.clsFuncionario
  Set objFuncionario = New datSisMaq.clsFuncionario
  
  'Listar Funcionario
  Set SelecionarFuncionarioPeloPkid = objFuncionario.SelecionarFuncionarioPeloPkid(lngFuncionarioId)
  '
  Set objFuncionario = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub InserirGerente(lngPRONTUARIOID As Long)
  On Error GoTo trata
  Dim objFuncionario     As datSisMaq.clsFuncionario
  '
  Set objFuncionario = New datSisMaq.clsFuncionario
  'Inserir Gerente
  objFuncionario.InserirGerente lngPRONTUARIOID
  '
  Set objFuncionario = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirAtendente(lngPRONTUARIOID As Long)
  On Error GoTo trata
  Dim objFuncionario     As datSisMaq.clsFuncionario
  '
  Set objFuncionario = New datSisMaq.clsFuncionario
  'Inserir Atendente
  objFuncionario.InserirAtendente lngPRONTUARIOID
  '
  Set objFuncionario = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirArrecadador(lngPRONTUARIOID As Long)
  On Error GoTo trata
  Dim objFuncionario     As datSisMaq.clsFuncionario
  '
  Set objFuncionario = New datSisMaq.clsFuncionario
  'Inserir Arrecadador
  objFuncionario.InserirArrecadador lngPRONTUARIOID
  '
  Set objFuncionario = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirDono(lngPRONTUARIOID As Long)
  On Error GoTo trata
  Dim objFuncionario     As datSisMaq.clsFuncionario
  '
  Set objFuncionario = New datSisMaq.clsFuncionario
  'Inserir Dono
  objFuncionario.InserirDono lngPRONTUARIOID
  '
  Set objFuncionario = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirLeiturista(lngPRONTUARIOID As Long)
  On Error GoTo trata
  Dim objFuncionario     As datSisMaq.clsFuncionario
  '
  Set objFuncionario = New datSisMaq.clsFuncionario
  'Inserir Leiturista
  objFuncionario.InserirLeiturista lngPRONTUARIOID
  '
  Set objFuncionario = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirCaixa(lngPESSOAID As Long)
  On Error GoTo trata
  Dim objFuncionario     As datSisMaq.clsFuncionario
  '
  Set objFuncionario = New datSisMaq.clsFuncionario
  'Inserir Caixa
  objFuncionario.InserirCaixa lngPESSOAID
  '
  Set objFuncionario = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirFuncionario(lngPRONTUARIOID As Long, _
                              strUsuario As String, _
                              strNivel As String, _
                              strSenha As String, _
                              strExcluido As String)
  On Error GoTo trata
  Dim objFuncionario     As datSisMaq.clsFuncionario
  '
  Set objFuncionario = New datSisMaq.clsFuncionario
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
  Dim objFuncionario     As datSisMaq.clsFuncionario
  '
  Set objFuncionario = New datSisMaq.clsFuncionario
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

