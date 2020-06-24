VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarCor(lngCORID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objCor As datSisMetal.clsCor
  Set objCor = New datSisMetal.clsCor
  
  'Listar Cor
  Set SelecionarCor = objCor.SelecionarCor(lngCORID)
  '
  Set objCor = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirCor(strSigla As String, _
                      strNome As String)
  
  On Error GoTo trata
  Dim objCor As datSisMetal.clsCor
  Set objCor = New datSisMetal.clsCor
  
  'Inserir Cor
  objCor.IncluirCor strSigla, _
                    strNome
  '
  Set objCor = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarCor(strSigla As String, _
                      strNome As String, _
                      lngCORID As Long)
  On Error GoTo trata
  Dim objCor As datSisMetal.clsCor
  Set objCor = New datSisMetal.clsCor
  
  'Alterar Cor
  objCor.AlterarCor strSigla, _
                    strNome, _
                    lngCORID

  '
  Set objCor = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub
Public Sub ExcluirCor(lngCORID As Long)
  On Error GoTo trata
  Dim objCor As datSisMetal.clsCor
  Set objCor = New datSisMetal.clsCor
  
  'Excluir Cor
  objCor.ExcluirCor lngCORID

  '
  Set objCor = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


