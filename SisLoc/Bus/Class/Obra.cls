VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function CapturaObra(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objObra As datSisLoc.clsObra
  Set objObra = New datSisLoc.clsObra
  
  'Capturar Obra
  Set CapturaObra = objObra.CapturaObra(strDescricao)
  '
  Set objObra = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarObraPeloPkid(lngObraId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objObra As datSisLoc.clsObra
  Set objObra = New datSisLoc.clsObra
  
  'Listar Obra
  Set SelecionarObraPeloPkid = objObra.SelecionarObraPeloPkid(lngObraId)
  '
  Set objObra = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirObra(lngObraId As Long)
  On Error GoTo trata
  Dim objObra As datSisLoc.clsObra
  Set objObra = New datSisLoc.clsObra
  
  'Excluir Obra
  objObra.ExcluirObra lngObraId
  '
  Set objObra = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirObra(lngContratoId As Long, _
                       strDescricao As String)
  On Error GoTo trata
  Dim objObra     As datSisLoc.clsObra
  '
  Set objObra = New datSisLoc.clsObra
  'Inserir Obra
  objObra.InserirObra lngContratoId, _
                      strDescricao
  '
  Set objObra = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarObra(lngPKID As Long, _
                       strDescricao As String)
  On Error GoTo trata
  Dim objObra     As datSisLoc.clsObra
  '
  Set objObra = New datSisLoc.clsObra
  'Atualizar Obra
  objObra.AlterarObra lngPKID, _
                      strDescricao
  '
  Set objObra = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

