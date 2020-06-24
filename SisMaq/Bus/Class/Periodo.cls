VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirPeriodo(strPeriodo As String, _
                          strInicio As String, _
                          strTermino As String)

  On Error GoTo trata
  Dim objPeriodo As datSisMaq.clsPeriodo
  '
  Set objPeriodo = New datSisMaq.clsPeriodo
  
  'Inserir Periodo
  objPeriodo.InserirPeriodo strPeriodo, _
                            strInicio, _
                            strTermino
                            
  '
  Set objPeriodo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarPeriodo(lngPERIODOID As Long, _
                          strPeriodo As String, _
                          strInicio As String, _
                          strTermino As String)
                         
  On Error GoTo trata
  Dim objPeriodo As datSisMaq.clsPeriodo
  Set objPeriodo = New datSisMaq.clsPeriodo
  
  'Inserir Periodo
  objPeriodo.AlterarPeriodo lngPERIODOID, _
                            strPeriodo, _
                            strInicio, _
                            strTermino

  '
  Set objPeriodo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarPeriodo(lngPERIODOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objPeriodo As datSisMaq.clsPeriodo
  Set objPeriodo = New datSisMaq.clsPeriodo
  
  'Listar Periodo
  Set ListarPeriodo = objPeriodo.ListarPeriodo(lngPERIODOID)
  '
  Set objPeriodo = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirPeriodo(lngPERIODOID As Long)

  On Error GoTo trata
  Dim objPeriodo As datSisMaq.clsPeriodo
  Set objPeriodo = New datSisMaq.clsPeriodo
  
  'Excluir Periodo
  objPeriodo.ExlcuirPeriodo lngPERIODOID
  '
  Set objPeriodo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

