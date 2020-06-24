VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmbalagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarEmbalagem(lngEMBALAGEMID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objEmbalagem As datSisMetal.clsEmbalagem
  Set objEmbalagem = New datSisMetal.clsEmbalagem
  
  'Listar Embalagem
  Set SelecionarEmbalagem = objEmbalagem.SelecionarEmbalagem(lngEMBALAGEMID)
  '
  Set objEmbalagem = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirEmbalagem(strSigla As String, _
                            strNome As String)
  
  On Error GoTo trata
  Dim objEmbalagem As datSisMetal.clsEmbalagem
  Set objEmbalagem = New datSisMetal.clsEmbalagem
  
  'Inserir Embalagem
  objEmbalagem.IncluirEmbalagem strSigla, _
                                strNome
  '
  Set objEmbalagem = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarEmbalagem(strSigla As String, _
                            strNome As String, _
                            lngEMBALAGEMID As Long)
  On Error GoTo trata
  Dim objEmbalagem As datSisMetal.clsEmbalagem
  Set objEmbalagem = New datSisMetal.clsEmbalagem
  
  'Alterar Embalagem
  objEmbalagem.AlterarEmbalagem strSigla, _
                                strNome, _
                                lngEMBALAGEMID

  '
  Set objEmbalagem = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub
Public Sub ExcluirEmbalagem(lngEMBALAGEMID As Long)
  On Error GoTo trata
  Dim objEmbalagem As datSisMetal.clsEmbalagem
  Set objEmbalagem = New datSisMetal.clsEmbalagem
  
  'Excluir Embalagem
  objEmbalagem.ExcluirEmbalagem lngEMBALAGEMID

  '
  Set objEmbalagem = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


