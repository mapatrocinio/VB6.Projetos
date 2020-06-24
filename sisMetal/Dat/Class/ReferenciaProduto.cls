VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReferenciaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function ListarReferenciaProdutoPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT REFERENCIAPRODUTO.* " & _
    "FROM REFERENCIAPRODUTO " & _
    "WHERE REFERENCIAPRODUTO.REFERENCIA = " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarReferenciaProdutoPelaDesc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsReferenciaProduto.ListarReferenciaProdutoPelaDesc]", _
            Err.Description
End Function

Public Sub InserirReferenciaProduto(strReferencia As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO REFERENCIAPRODUTO (REFERENCIA) VALUES (" & _
    Formata_Dados(strReferencia, tpDados_Texto, tpNulo_NaoAceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsReferenciaProduto.InserirReferenciaProduto]", _
            Err.Description
End Sub
Public Sub AlterarReferenciaProduto(lngREFERENCIAPRODUTOID As Long, _
                                    strReferencia As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE REFERENCIAPRODUTO SET REFERENCIA = " & Formata_Dados(strReferencia, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngREFERENCIAPRODUTOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsReferenciaProduto.AlterarReferenciaProduto]", _
            Err.Description
End Sub


Public Function ListarReferenciaProduto(lngREFERENCIAPRODUTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT REFERENCIAPRODUTO.* " & _
    "FROM REFERENCIAPRODUTO " & _
    "WHERE REFERENCIAPRODUTO.PKID = " & _
    Formata_Dados(lngREFERENCIAPRODUTOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarReferenciaProduto = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsReferenciaProduto.ListarReferenciaProduto]", _
            Err.Description
End Function

Public Sub ExcluirREFERENCIAPRODUTO(lngREFERENCIAPRODUTOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui DOCUMENTO DE ENTRADA
  strSql = "DELETE FROM REFERENCIAPRODUTO WHERE REFERENCIAPRODUTO.PKID = " & _
    Formata_Dados(lngREFERENCIAPRODUTOID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsReferenciaProduto.ExcluirREFERENCIAPRODUTO]", _
            Err.Description
End Sub
