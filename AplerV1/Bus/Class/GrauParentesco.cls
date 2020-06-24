VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrauParentesco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ListarGrauParentescoPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objGrauParentesco As datApler.clsGrauParentesco
  Set objGrauParentesco = New datApler.clsGrauParentesco
  
  'Listar GrauParentesco pela descrição
  Set ListarGrauParentescoPelaDesc = objGrauParentesco.ListarGrauParentescoPelaDesc(strDescricao)
  '
  Set objGrauParentesco = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub InserirGrauParentesco(strDescricao As String)

  On Error GoTo trata
  Dim objGrauParentesco As datApler.clsGrauParentesco
  '
  Set objGrauParentesco = New datApler.clsGrauParentesco
  
  'Inserir GrauParentesco
  objGrauParentesco.InserirGrauParentesco strDescricao
  '
  Set objGrauParentesco = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarGrauParentesco(lngGRAUPARENTESCOID As Long, _
                                 strDescricao As String)
                         
  On Error GoTo trata
  Dim objGrauParentesco As datApler.clsGrauParentesco
  Set objGrauParentesco = New datApler.clsGrauParentesco
  
  'Inserir GrauParentesco
  objGrauParentesco.AlterarGrauParentesco lngGRAUPARENTESCOID, _
                                          strDescricao

  '
  Set objGrauParentesco = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarGrauParentesco(lngGRAUPARENTESCOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objGrauParentesco As datApler.clsGrauParentesco
  Set objGrauParentesco = New datApler.clsGrauParentesco
  
  'Listar GrauParentesco
  Set ListarGrauParentesco = objGrauParentesco.ListarGrauParentesco(lngGRAUPARENTESCOID)
  '
  Set objGrauParentesco = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirGrauParentesco(lngGRAUPARENTESCOID As Long)

  On Error GoTo trata
  Dim objGrauParentesco As datApler.clsGrauParentesco
  Set objGrauParentesco = New datApler.clsGrauParentesco
  
  'Excluir GrauParentesco
  objGrauParentesco.ExlcuirGrauParentesco lngGRAUPARENTESCOID
  '
  Set objGrauParentesco = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

