VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLinhaPerfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function CapturaItemLinha(ByVal strCodigo As String) As ADODB.Recordset
  '
  On Error GoTo trata
    
  Dim objLinhaPerfil As datSisMetal.clsLinhaPerfil
  '
  Set objLinhaPerfil = New datSisMetal.clsLinhaPerfil
  
  Set CapturaItemLinha = objLinhaPerfil.CapturaItemLinha(strCodigo)
  '
  Set objLinhaPerfil = Nothing
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarLinha(lngLINHAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objLinhaPerfil As datSisMetal.clsLinhaPerfil
  Set objLinhaPerfil = New datSisMetal.clsLinhaPerfil
  
  'Listar Linha
  Set SelecionarLinha = objLinhaPerfil.SelecionarLinha(lngLINHAID)
  '
  Set objLinhaPerfil = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirLinha(lngTIPOLINHAID As Long, _
                        strCodigo As String, _
                        lngVARAID As Long, _
                        strPesoVara As String, _
                        strNomeProduto As String, _
                        strLargura As String, _
                        strAba As String, _
                        strEspessura As String)
  
  On Error GoTo trata
  Dim objLinhaPerfil As datSisMetal.clsLinhaPerfil
  Set objLinhaPerfil = New datSisMetal.clsLinhaPerfil
  
  'Inserir Linha
  objLinhaPerfil.IncluirLinha lngTIPOLINHAID, _
                        strCodigo, _
                        lngVARAID, _
                        strPesoVara, _
                        strNomeProduto, _
                        strLargura, _
                        strAba, _
                        strEspessura
  '
  Set objLinhaPerfil = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarLinha(lngTIPOLINHAID As Long, _
                        strCodigo As String, _
                        lngVARAID As Long, _
                        strPesoVara As String, _
                        lngLINHAID As Long, _
                        strNomeProduto As String, _
                        strLargura As String, _
                        strAba As String, _
                        strEspessura As String)
  On Error GoTo trata
  Dim objLinhaPerfil As datSisMetal.clsLinhaPerfil
  Set objLinhaPerfil = New datSisMetal.clsLinhaPerfil
  
  'Alterar Linha
  objLinhaPerfil.AlterarLinha lngTIPOLINHAID, _
                        strCodigo, _
                        lngVARAID, _
                        strPesoVara, _
                        lngLINHAID, _
                        strNomeProduto, _
                        strLargura, _
                        strAba, _
                        strEspessura

  '
  Set objLinhaPerfil = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub
Public Sub ExcluirLinha(lngLINHAID As Long)
  On Error GoTo trata
  Dim objLinhaPerfil As datSisMetal.clsLinhaPerfil
  Set objLinhaPerfil = New datSisMetal.clsLinhaPerfil
  
  'Excluir Linha
  objLinhaPerfil.ExcluirLinha lngLINHAID

  '
  Set objLinhaPerfil = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


