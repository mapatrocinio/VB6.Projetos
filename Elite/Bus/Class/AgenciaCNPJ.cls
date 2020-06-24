VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAgenciaCNPJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarAgenciaCNPJPeloPkid(lngAgenciaCNPJId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objAgenciaCNPJ As datElite.clsAgenciaCNPJ
  Set objAgenciaCNPJ = New datElite.clsAgenciaCNPJ
  
  'Listar AgenciaCNPJ
  Set SelecionarAgenciaCNPJPeloPkid = objAgenciaCNPJ.SelecionarAgenciaCNPJPeloPkid(lngAgenciaCNPJId)
  '
  Set objAgenciaCNPJ = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirAgenciaCNPJ(lngAgenciaCNPJId As Long)
  On Error GoTo trata
  Dim objAgenciaCNPJ As datElite.clsAgenciaCNPJ
  Set objAgenciaCNPJ = New datElite.clsAgenciaCNPJ
  
  'Excluir AgenciaCNPJ
  objAgenciaCNPJ.ExcluirAgenciaCNPJ lngAgenciaCNPJId
  '
  Set objAgenciaCNPJ = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirAgenciaCNPJ(lngAGENCIAID As Long, _
                              strCNPJ As String)
  On Error GoTo trata
  Dim objAgenciaCNPJ     As datElite.clsAgenciaCNPJ
  '
  Set objAgenciaCNPJ = New datElite.clsAgenciaCNPJ
  'Inserir AgenciaCNPJ
  objAgenciaCNPJ.InserirAgenciaCNPJ lngAGENCIAID, _
                                    strCNPJ
  '
  Set objAgenciaCNPJ = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarAgenciaCNPJ(lngPKID As Long, _
                              strCNPJ As String)

  On Error GoTo trata
  Dim objAgenciaCNPJ     As datElite.clsAgenciaCNPJ
  '
  Set objAgenciaCNPJ = New datElite.clsAgenciaCNPJ
  'Atualizar AgenciaCNPJ
  objAgenciaCNPJ.AlterarAgenciaCNPJ lngPKID, _
                                    strCNPJ
  '
  Set objAgenciaCNPJ = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

