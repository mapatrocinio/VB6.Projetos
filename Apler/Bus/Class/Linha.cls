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

Public Function ListarLinhaPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objLinha As datApler.clsLinha
  Set objLinha = New datApler.clsLinha
  
  'Listar Linha pela descrição
  Set ListarLinhaPelaDesc = objLinha.ListarLinhaPelaDesc(strDescricao)
  '
  Set objLinha = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub InserirLinha(strDescricao As String)

  On Error GoTo trata
  Dim objLinha As datApler.clsLinha
  '
  Set objLinha = New datApler.clsLinha
  
  'Inserir Linha
  objLinha.InserirLinha strDescricao
  '
  Set objLinha = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarLinha(lngLINHAID As Long, _
                        strDescricao As String)
                         
  On Error GoTo trata
  Dim objLinha As datApler.clsLinha
  Set objLinha = New datApler.clsLinha
  
  'Inserir Linha
  objLinha.AlterarLinha lngLINHAID, _
                        strDescricao

  '
  Set objLinha = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarLinha(lngLINHAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objLinha As datApler.clsLinha
  Set objLinha = New datApler.clsLinha
  
  'Listar Linha
  Set ListarLinha = objLinha.ListarLinha(lngLINHAID)
  '
  Set objLinha = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirLinha(lngLINHAID As Long)

  On Error GoTo trata
  Dim objLinha As datApler.clsLinha
  Set objLinha = New datApler.clsLinha
  
  'Excluir Linha
  objLinha.ExlcuirLinha lngLINHAID
  '
  Set objLinha = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

