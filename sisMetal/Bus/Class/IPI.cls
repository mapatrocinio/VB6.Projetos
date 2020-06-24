VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirIPI(strIPI As String)

  On Error GoTo trata
  Dim objIPI  As datSisMetal.clsIPI
  '
  Set objIPI = New datSisMetal.clsIPI
  
  'Inserir Família de produtos
  objIPI.InserirIPI strIPI
  '
  Set objIPI = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarIPI(lngIPIID As Long, _
                      strIPI As String)
                         
  On Error GoTo trata
  Dim objIPI As datSisMetal.clsIPI
  Set objIPI = New datSisMetal.clsIPI
  
  'Alterar Família de produtos
  objIPI.AlterarIPI lngIPIID, _
                    strIPI

  '
  Set objIPI = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarIPI(lngIPIID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objIPI As datSisMetal.clsIPI
  Set objIPI = New datSisMetal.clsIPI
  
  'Listar Família de produtos
  Set ListarIPI = objIPI.ListarIPI(lngIPIID)
  '
  Set objIPI = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirIPI(ByVal lngIPIID As Long)
  Dim objIPI    As datSisMetal.clsIPI
  '
  On Error GoTo trata
  Set objIPI = New datSisMetal.clsIPI
  '
  objIPI.ExcluirIPI lngIPIID
  '
  Set objIPI = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


