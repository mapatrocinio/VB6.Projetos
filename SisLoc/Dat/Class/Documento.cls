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
Public Sub InserirDocumento(strNome As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO DOCUMENTO (NOME) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDocumento.InserirDocumento]", _
            Err.Description
End Sub

Public Sub AlterarDocumento(lngDOCUMENTOID As Long, _
                               strNome As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE DOCUMENTO SET NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngDOCUMENTOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDocumento.AlterarDocumento]", _
            Err.Description
End Sub

Public Function ListarDocumento(lngDOCUMENTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT DOCUMENTO.* " & _
    "FROM DOCUMENTO " & _
    "WHERE DOCUMENTO.PKID = " & _
    Formata_Dados(lngDOCUMENTOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarDocumento = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDocumento.ListarDocumento]", _
            Err.Description
End Function

Public Sub ExlcuirDocumento(lngDOCUMENTOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM DOCUMENTO WHERE PKID = " & _
    Formata_Dados(lngDOCUMENTOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDocumento.ExlcuirDocumento]", _
            Err.Description
End Sub


Public Function ListarDocumentoPelaDesc(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT DOCUMENTO.* " & _
    "FROM DOCUMENTO " & _
    "WHERE DOCUMENTO.NOME = " & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarDocumentoPelaDesc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDocumento.ListarDocumentoPelaDesc]", _
            Err.Description
End Function

