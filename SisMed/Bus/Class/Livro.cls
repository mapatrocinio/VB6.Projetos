VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLivro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarLivro(lngLIVROID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objLivro As datSisMed.clsLivro
  Set objLivro = New datSisMed.clsLivro
  
  'Listar Livro
  Set SelecionarLivro = objLivro.SelecionarLivro(lngLIVROID)
  '
  Set objLivro = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirLivro(lngBancoId As Long, _
                        strConta As String, _
                        strAgencia As String, _
                        strNumeroLivro As String)
  
  On Error GoTo trata
  Dim objLivro As datSisMed.clsLivro
  Set objLivro = New datSisMed.clsLivro
  
  'Inserir Livro
  objLivro.IncluirLivro lngBancoId, _
                        strConta, _
                        strAgencia, _
                        strNumeroLivro
  '
  Set objLivro = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarLivro(lngLIVROID As Long, _
                        lngBancoId As Long, _
                        strConta As String, _
                        strAgencia As String, _
                        strNumeroLivro As String)
  On Error GoTo trata
  Dim objLivro As datSisMed.clsLivro
  Set objLivro = New datSisMed.clsLivro
  
  'Inserir Livro
  objLivro.AlterarLivro lngLIVROID, _
                        lngBancoId, _
                        strConta, _
                        strAgencia, _
                        strNumeroLivro
  '
  Set objLivro = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub
Public Sub ExcluirLivro(lngLIVROID As Long)
  On Error GoTo trata
  Dim objLivro As datSisMed.clsLivro
  Set objLivro = New datSisMed.clsLivro
  
  'Excluir Livro
  objLivro.ExcluirLivro lngLIVROID

  '
  Set objLivro = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


