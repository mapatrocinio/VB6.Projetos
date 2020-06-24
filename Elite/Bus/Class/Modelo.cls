VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirModelo(strNome As String, _
                         lngMARCAID As Long)

  On Error GoTo trata
  Dim objModelo  As datElite.clsModelo
  '
  Set objModelo = New datElite.clsModelo
  
  'Inserir Família de produtos
  objModelo.InserirModelo strNome, _
                          lngMARCAID
  '
  Set objModelo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarModelo(lngMODELOID As Long, _
                         strNome As String, _
                         lngMARCAID As Long)
                         
  On Error GoTo trata
  Dim objModelo As datElite.clsModelo
  Set objModelo = New datElite.clsModelo
  
  'Alterar Família de produtos
  objModelo.AlterarModelo lngMODELOID, _
                          strNome, _
                          lngMARCAID

  '
  Set objModelo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarModelo(lngMODELOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objModelo As datElite.clsModelo
  Set objModelo = New datElite.clsModelo
  
  'Listar Família de produtos
  Set ListarModelo = objModelo.ListarModelo(lngMODELOID)
  '
  Set objModelo = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirModelo(ByVal lngMODELOID As Long)
  Dim objModelo    As datElite.clsModelo
  '
  On Error GoTo trata
  Set objModelo = New datElite.clsModelo
  '
  objModelo.ExcluirModelo lngMODELOID
  '
  Set objModelo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


