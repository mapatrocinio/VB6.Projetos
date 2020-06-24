VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarSalaPeloNumero(strNumero As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objSala As datSisMed.clsSala
  Set objSala = New datSisMed.clsSala
  
  'Listar Sala
  Set SelecionarSalaPeloNumero = objSala.SelecionarSalaPeloNumero(strNumero)
  '
  Set objSala = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarSalaPeloPkid(lngSalaId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objSala As datSisMed.clsSala
  Set objSala = New datSisMed.clsSala
  
  'Listar Sala
  Set SelecionarSalaPeloPkid = objSala.SelecionarSalaPeloPkid(lngSalaId)
  '
  Set objSala = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirSala(lngSalaId As Long)
  On Error GoTo trata
  Dim objSala As datSisMed.clsSala
  Set objSala = New datSisMed.clsSala
  
  'Excluir Sala
  objSala.ExcluirSala lngSalaId
  '
  Set objSala = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirSala(lngPREDIOID As Long, _
                       strNumero As String, _
                       strAndar As String, _
                       strTelefone As String, _
                       strStatus As String)
  On Error GoTo trata
  Dim objSala     As datSisMed.clsSala
  '
  Set objSala = New datSisMed.clsSala
  'Inserir Sala
  objSala.InserirSala lngPREDIOID, _
                       strNumero, _
                       strAndar, _
                       strTelefone, _
                       strStatus
  '
  Set objSala = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarSala(lngPKID As Long, _
                       lngPREDIOID As Long, _
                       strNumero As String, _
                       strAndar As String, _
                       strTelefone As String, _
                       strStatus As String)

  On Error GoTo trata
  Dim objSala     As datSisMed.clsSala
  '
  Set objSala = New datSisMed.clsSala
  'Atualizar Sala
  objSala.AlterarSala lngPKID, _
                       lngPREDIOID, _
                       strNumero, _
                       strAndar, _
                       strTelefone, _
                       strStatus
  '
  Set objSala = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

