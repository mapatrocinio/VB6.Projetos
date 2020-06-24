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

Public Function ListarOrigemPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objOrigem As datApler.clsOrigem
  Set objOrigem = New datApler.clsOrigem
  
  'Listar Origem pela descrição
  Set ListarOrigemPelaDesc = objOrigem.ListarOrigemPelaDesc(strDescricao)
  '
  Set objOrigem = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub InserirOrigem(strDescricao As String)

  On Error GoTo trata
  Dim objOrigem As datApler.clsOrigem
  '
  Set objOrigem = New datApler.clsOrigem
  
  'Inserir Origem
  objOrigem.InserirOrigem strDescricao
  '
  Set objOrigem = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarOrigem(lngORIGEMID As Long, _
                         strDescricao As String)
                         
  On Error GoTo trata
  Dim objOrigem As datApler.clsOrigem
  Set objOrigem = New datApler.clsOrigem
  
  'Inserir Origem
  objOrigem.AlterarOrigem lngORIGEMID, _
                          strDescricao

  '
  Set objOrigem = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarOrigem(lngORIGEMID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objOrigem As datApler.clsOrigem
  Set objOrigem = New datApler.clsOrigem
  
  'Listar Origem
  Set ListarOrigem = objOrigem.ListarOrigem(lngORIGEMID)
  '
  Set objOrigem = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirOrigem(lngORIGEMID As Long)

  On Error GoTo trata
  Dim objOrigem As datApler.clsOrigem
  Set objOrigem = New datApler.clsOrigem
  
  'Excluir Origem
  objOrigem.ExlcuirOrigem lngORIGEMID
  '
  Set objOrigem = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

