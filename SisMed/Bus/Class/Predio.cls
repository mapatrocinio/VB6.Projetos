VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPredio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirPredio(strNome As String)

  On Error GoTo trata
  Dim objPredio As datSisMed.clsPredio
  '
  Set objPredio = New datSisMed.clsPredio
  
  'Inserir Predio
  objPredio.InserirPredio strNome
  '
  Set objPredio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarPredio(lngPREDIOID As Long, _
                         strNome As String)
                         
  On Error GoTo trata
  Dim objPredio As datSisMed.clsPredio
  Set objPredio = New datSisMed.clsPredio
  
  'Inserir Predio
  objPredio.AlterarPredio lngPREDIOID, _
                          strNome

  '
  Set objPredio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarPredio(lngPREDIOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objPredio As datSisMed.clsPredio
  Set objPredio = New datSisMed.clsPredio
  
  'Listar Predio
  Set ListarPredio = objPredio.ListarPredio(lngPREDIOID)
  '
  Set objPredio = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirPredio(lngPREDIOID As Long)

  On Error GoTo trata
  Dim objPredio As datSisMed.clsPredio
  Set objPredio = New datSisMed.clsPredio
  
  'Excluir Predio
  objPredio.ExlcuirPredio lngPREDIOID
  '
  Set objPredio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

