VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLinha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function SelecionarLinha(lngTIPOLINHAIDID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objLinha As datSisMetal.clsLinha
  Set objLinha = New datSisMetal.clsLinha
  
  'Listar Linha
  Set SelecionarLinha = objLinha.SelecionarLinha(lngTIPOLINHAIDID)
  '
  Set objLinha = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirLinha(strNome As String)
  
  On Error GoTo trata
  Dim objLinha As datSisMetal.clsLinha
  Set objLinha = New datSisMetal.clsLinha
  
  'Inserir Linha
  objLinha.IncluirLinha strNome
  '
  Set objLinha = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarLinha(strNome As String, _
                        lngTIPOLINHAIDID As Long)
  On Error GoTo trata
  Dim objLinha As datSisMetal.clsLinha
  Set objLinha = New datSisMetal.clsLinha
  
  'Alterar Linha
  objLinha.AlterarLinha strNome, _
                        lngTIPOLINHAIDID

  '
  Set objLinha = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub
Public Sub ExcluirLinha(lngTIPOLINHAIDID As Long)
  On Error GoTo trata
  Dim objLinha As datSisMetal.clsLinha
  Set objLinha = New datSisMetal.clsLinha
  
  'Excluir Linha
  objLinha.ExcluirLinha lngTIPOLINHAIDID

  '
  Set objLinha = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


