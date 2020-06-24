VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVeiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirVeiculo(lngVEICULOID As Long, _
                         strPlaca As String, _
                         lngMODELOID As Long, _
                         strAno As String, _
                         strObservacao As String)

  On Error GoTo trata
  Dim objVeiculo  As datElite.clsVeiculo
  '
  Set objVeiculo = New datElite.clsVeiculo
  
  'Inserir Família de produtos
  objVeiculo.InserirVeiculo lngVEICULOID, _
                            strPlaca, _
                            lngMODELOID, _
                            strAno, _
                            strObservacao
  '
  Set objVeiculo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarVeiculo(lngVEICULOID As Long, _
                         strPlaca As String, _
                         lngMODELOID As Long, _
                         strAno As String, _
                         strObservacao As String)
                         
  On Error GoTo trata
  Dim objVeiculo As datElite.clsVeiculo
  Set objVeiculo = New datElite.clsVeiculo
  
  'Alterar Família de produtos
  objVeiculo.AlterarVeiculo lngVEICULOID, _
                            strPlaca, _
                            lngMODELOID, _
                            strAno, _
                            strObservacao

  '
  Set objVeiculo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarVeiculo(lngVEICULOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objVeiculo As datElite.clsVeiculo
  Set objVeiculo = New datElite.clsVeiculo
  
  'Listar Família de produtos
  Set ListarVeiculo = objVeiculo.ListarVeiculo(lngVEICULOID)
  '
  Set objVeiculo = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirVeiculo(ByVal lngVEICULOID As Long)
  Dim objVeiculo    As datElite.clsVeiculo
  '
  On Error GoTo trata
  Set objVeiculo = New datElite.clsVeiculo
  '
  objVeiculo.ExcluirVeiculo lngVEICULOID
  '
  Set objVeiculo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


