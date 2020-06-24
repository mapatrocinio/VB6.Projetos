VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ListarDocumentoPelaDesc(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objDocumento As datSisLoc.clsDocumento
  Set objDocumento = New datSisLoc.clsDocumento
  
  'Listar Documento pela descrição
  Set ListarDocumentoPelaDesc = objDocumento.ListarDocumentoPelaDesc(strNome)
  '
  Set objDocumento = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub InserirDocumento(strNome As String)

  On Error GoTo trata
  Dim objDocumento As datSisLoc.clsDocumento
  '
  Set objDocumento = New datSisLoc.clsDocumento
  
  'Inserir Documento
  objDocumento.InserirDocumento strNome
  '
  Set objDocumento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarDocumento(lngDOCUMENTOID As Long, _
                            strNome As String)
                         
  On Error GoTo trata
  Dim objDocumento As datSisLoc.clsDocumento
  Set objDocumento = New datSisLoc.clsDocumento
  
  'Inserir Documento
  objDocumento.AlterarDocumento lngDOCUMENTOID, _
                                strNome

  '
  Set objDocumento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarDocumento(lngDOCUMENTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objDocumento As datSisLoc.clsDocumento
  Set objDocumento = New datSisLoc.clsDocumento
  
  'Listar Documento
  Set ListarDocumento = objDocumento.ListarDocumento(lngDOCUMENTOID)
  '
  Set objDocumento = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirDocumento(lngDOCUMENTOID As Long)

  On Error GoTo trata
  Dim objDocumento As datSisLoc.clsDocumento
  Set objDocumento = New datSisLoc.clsDocumento
  
  'Excluir Documento
  objDocumento.ExlcuirDocumento lngDOCUMENTOID
  '
  Set objDocumento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

