VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarSeriePeloNumero(strNumero As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objSerie As datSisMaq.clsSerie
  Set objSerie = New datSisMaq.clsSerie
  
  'Listar Serie
  Set SelecionarSeriePeloNumero = objSerie.SelecionarSeriePeloNumero(strNumero)
  '
  Set objSerie = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarSeriePeloPkid(lngSERIEID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objSerie As datSisMaq.clsSerie
  Set objSerie = New datSisMaq.clsSerie
  
  'Listar Serie
  Set SelecionarSeriePeloPkid = objSerie.SelecionarSeriePeloPkid(lngSERIEID)
  '
  Set objSerie = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirSerie(lngSERIEID As Long)
  On Error GoTo trata
  Dim objSerie As datSisMaq.clsSerie
  Set objSerie = New datSisMaq.clsSerie
  
  'Excluir Serie
  objSerie.ExcluirSerie lngSERIEID
  '
  Set objSerie = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirSerie(ByRef lngPKID As Long, _
                        lngDONOID As Long, _
                        strNumero As String, _
                        strPercDono As String, _
                        strPercCasa As String, _
                        strCoeficiente As String, _
                        strStatus As String)
  On Error GoTo trata
  Dim objSerie     As datSisMaq.clsSerie
  '
  Set objSerie = New datSisMaq.clsSerie
  'Inserir Serie
  objSerie.InserirSerie lngPKID, _
                        lngDONOID, _
                        strNumero, _
                        strPercDono, _
                        strPercCasa, _
                        strCoeficiente, _
                        strStatus
  '
  Set objSerie = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarSerie(lngPKID As Long, _
                        lngDONOID As Long, _
                        strNumero As String, _
                        strPercDono As String, _
                        strPercCasa As String, _
                        strCoeficiente As String, _
                        strStatus As String, _
                        strUsuario As String)

  On Error GoTo trata
  Dim objSerie     As datSisMaq.clsSerie
  '
  Set objSerie = New datSisMaq.clsSerie
  'Atualizar Serie
  objSerie.AlterarSerie lngPKID, _
                        lngDONOID, _
                        strNumero, _
                        strPercDono, _
                        strPercCasa, _
                        strCoeficiente, _
                        strStatus, _
                        strUsuario
  '
  Set objSerie = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

