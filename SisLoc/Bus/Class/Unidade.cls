VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUnidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ListarUnidadePelaDesc(strUnidade As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objUnidade As datSisLoc.clsUnidade
  Set objUnidade = New datSisLoc.clsUnidade
  
  'Listar Unidade pela descrição
  Set ListarUnidadePelaDesc = objUnidade.ListarUnidadePelaDesc(strUnidade)
  '
  Set objUnidade = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub InserirUnidade(strUnidade As String)

  On Error GoTo trata
  Dim objUnidade As datSisLoc.clsUnidade
  '
  Set objUnidade = New datSisLoc.clsUnidade
  
  'Inserir Unidade
  objUnidade.InserirUnidade strUnidade
  '
  Set objUnidade = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarUnidade(lngUNIDADEID As Long, _
                            strUnidade As String)
                         
  On Error GoTo trata
  Dim objUnidade As datSisLoc.clsUnidade
  Set objUnidade = New datSisLoc.clsUnidade
  
  'Inserir Unidade
  objUnidade.AlterarUnidade lngUNIDADEID, _
                                strUnidade

  '
  Set objUnidade = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarUnidade(lngUNIDADEID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objUnidade As datSisLoc.clsUnidade
  Set objUnidade = New datSisLoc.clsUnidade
  
  'Listar Unidade
  Set ListarUnidade = objUnidade.ListarUnidade(lngUNIDADEID)
  '
  Set objUnidade = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirUnidade(lngUNIDADEID As Long)

  On Error GoTo trata
  Dim objUnidade As datSisLoc.clsUnidade
  Set objUnidade = New datSisLoc.clsUnidade
  
  'Excluir Unidade
  objUnidade.ExlcuirUnidade lngUNIDADEID
  '
  Set objUnidade = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub



