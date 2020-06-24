VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrigem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirOrigem(strNome As String, _
                         strIndicador As String)

  On Error GoTo trata
  Dim clsOrigem  As datElite.clsOrigem
  '
  Set clsOrigem = New datElite.clsOrigem
  
  'Inserir Origem
  clsOrigem.InserirOrigem strNome, _
                          Left(strIndicador, 1)
  '
  Set clsOrigem = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarOrigem(lngORIGEMID As Long, _
                        strNome As String, _
                        strIndicador As String)
                         
  On Error GoTo trata
  Dim clsOrigem As datElite.clsOrigem
  '
  Set clsOrigem = New datElite.clsOrigem
  'Inserir Documento de Saida
  clsOrigem.AlterarOrigem lngORIGEMID, _
                        strNome, _
                        Left(strIndicador, 1)

  '
  Set clsOrigem = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarOrigem(lngORIGEMID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsOrigem As datElite.clsOrigem
  Set clsOrigem = New datElite.clsOrigem
  
  'Listar Documento de entrada
  Set ListarOrigem = clsOrigem.ListarOrigem(lngORIGEMID)
  '
  Set clsOrigem = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirORIGEM(ByVal lngORIGEMID As Long)
  Dim clsOrigem    As datElite.clsOrigem
  '
  On Error GoTo trata
  Set clsOrigem = New datElite.clsOrigem
  '
  clsOrigem.ExcluirORIGEM lngORIGEMID
  '
  Set clsOrigem = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
