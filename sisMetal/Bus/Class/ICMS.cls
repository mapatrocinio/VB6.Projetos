VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirICMS(strDescricao As String)

  On Error GoTo trata
  Dim objICMS  As datSisMetal.clsICMS
  '
  Set objICMS = New datSisMetal.clsICMS
  
  'Inserir Família de produtos
  objICMS.InserirICMS strDescricao
  '
  Set objICMS = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarICMS(lngICMSID As Long, _
                       strICMS As String)
                         
  On Error GoTo trata
  Dim objICMS As datSisMetal.clsICMS
  Set objICMS = New datSisMetal.clsICMS
  
  'Alterar Família de produtos
  objICMS.AlterarICMS lngICMSID, _
                      strICMS

  '
  Set objICMS = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarICMS(lngICMSID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objICMS As datSisMetal.clsICMS
  Set objICMS = New datSisMetal.clsICMS
  
  'Listar Família de produtos
  Set ListarICMS = objICMS.ListarICMS(lngICMSID)
  '
  Set objICMS = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirICMS(ByVal lngICMSID As Long)
  Dim objICMS    As datSisMetal.clsICMS
  '
  On Error GoTo trata
  Set objICMS = New datSisMetal.clsICMS
  '
  objICMS.ExcluirICMS lngICMSID
  '
  Set objICMS = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


