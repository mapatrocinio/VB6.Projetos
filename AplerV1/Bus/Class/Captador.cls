VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaptador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarCaptadorPeloPkid(lngCaptadorId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objCaptador As datApler.clsCaptador
  Set objCaptador = New datApler.clsCaptador
  
  'Listar Captador
  Set SelecionarCaptadorPeloPkid = objCaptador.SelecionarCaptadorPeloPkid(lngCaptadorId)
  '
  Set objCaptador = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirCaptador(lngCaptadorId As Long)
  On Error GoTo trata
  Dim objCaptador As datApler.clsCaptador
  Set objCaptador = New datApler.clsCaptador
  
  'Excluir Captador
  objCaptador.ExcluirCaptador lngCaptadorId
  '
  Set objCaptador = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirCaptador(strNome As String)
  On Error GoTo trata
  Dim objCaptador     As datApler.clsCaptador
  '
  Set objCaptador = New datApler.clsCaptador
  'Inserir Captador
  objCaptador.InserirCaptador strNome
  '
  Set objCaptador = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarCaptador(lngPKID As Long, _
                           strNome As String, _
                           strStatus As String)

  On Error GoTo trata
  Dim objCaptador     As datApler.clsCaptador
  '
  Set objCaptador = New datApler.clsCaptador
  'Atualizar Captador
  objCaptador.AlterarCaptador lngPKID, _
                              strNome, _
                              strStatus
  '
  Set objCaptador = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

