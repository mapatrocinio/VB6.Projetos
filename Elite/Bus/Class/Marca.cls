VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMarca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirMarca(strNome As String)

  On Error GoTo trata
  Dim clsMarca  As datElite.clsMarca
  '
  Set clsMarca = New datElite.clsMarca
  
  'Inserir Documento de entrada
  clsMarca.InserirMarca strNome
  '
  Set clsMarca = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarMarca(lngMARCAID As Long, _
                        strNome As String)
                         
  On Error GoTo trata
  Dim clsMarca As datElite.clsMarca
  Set clsMarca = New datElite.clsMarca
  
  'Inserir Documento de Saida
  clsMarca.AlterarMarca lngMARCAID, _
                        strNome

  '
  Set clsMarca = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarMarca(lngMARCAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsMarca As datElite.clsMarca
  Set clsMarca = New datElite.clsMarca
  
  'Listar Documento de entrada
  Set ListarMarca = clsMarca.ListarMarca(lngMARCAID)
  '
  Set clsMarca = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirMARCA(ByVal lngMARCAID As Long)
  Dim clsMarca    As datElite.clsMarca
  '
  On Error GoTo trata
  Set clsMarca = New datElite.clsMarca
  '
  clsMarca.ExcluirMARCA lngMARCAID
  '
  Set clsMarca = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
