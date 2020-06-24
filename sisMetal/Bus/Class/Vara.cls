VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarVara(lngVARAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objVara As datSisMetal.clsVara
  Set objVara = New datSisMetal.clsVara
  
  'Listar Vara
  Set SelecionarVara = objVara.SelecionarVara(lngVARAID)
  '
  Set objVara = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirVara(strValor As String, _
                       strNome As String)
  
  On Error GoTo trata
  Dim objVara As datSisMetal.clsVara
  Set objVara = New datSisMetal.clsVara
  
  'Inserir Vara
  objVara.IncluirVara strValor, _
                      strNome
  '
  Set objVara = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarVara(strValor As String, _
                       strNome As String, _
                       lngVARAID As Long)
  On Error GoTo trata
  Dim objVara As datSisMetal.clsVara
  Set objVara = New datSisMetal.clsVara
  
  'Alterar Vara
  objVara.AlterarVara strValor, _
                      strNome, _
                      lngVARAID

  '
  Set objVara = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub
Public Sub ExcluirVara(lngVARAID As Long)
  On Error GoTo trata
  Dim objVara As datSisMetal.clsVara
  Set objVara = New datSisMetal.clsVara
  
  'Excluir Vara
  objVara.ExcluirVara lngVARAID

  '
  Set objVara = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


