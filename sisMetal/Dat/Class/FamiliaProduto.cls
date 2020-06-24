VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFamiliaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirFamiliaProduto(strDescricao As String, _
                                 lngIPIID As Long, _
                                 lngICMSID As Long)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO FAMILIAPRODUTOS (DESCRICAO, IPIID, ICMSID) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & _
    ", " & Formata_Dados(IIf(lngIPIID = 0, "", lngIPIID), tpDados_Longo) & _
    ", " & Formata_Dados(IIf(lngICMSID = 0, "", lngICMSID), tpDados_Longo) & _
    ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFamiliaProduto.InserirFamiliaProduto]", _
            Err.Description
End Sub

Public Sub AlterarFamiliaProduto(lngFAMILIAPRODUTOSID As Long, _
                                 strDescricao As String, _
                                 lngIPIID As Long, _
                                 lngICMSID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE FAMILIAPRODUTOS SET DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    " IPIID = " & Formata_Dados(IIf(lngIPIID = 0, "", lngIPIID), tpDados_Longo) & ", " & _
    " ICMSID = " & Formata_Dados(IIf(lngICMSID = 0, "", lngICMSID), tpDados_Longo) & " " & _
    " WHERE PKID = " & lngFAMILIAPRODUTOSID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFamiliaProduto.AlterarFamiliaProduto]", _
            Err.Description
End Sub

Public Function ListarFamiliaProduto(lngFAMILIAPRODUTOSID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT FAMILIAPRODUTOS.*, IPI.IPI AS IPI, ICMS.ICMS AS ICMS " & _
    "FROM FAMILIAPRODUTOS " & _
    " LEFT JOIN IPI ON IPI.PKID = FAMILIAPRODUTOS.IPIID " & _
    " LEFT JOIN ICMS ON ICMS.PKID = FAMILIAPRODUTOS.ICMSID " & _
    "WHERE FAMILIAPRODUTOS.PKID = " & _
    Formata_Dados(lngFAMILIAPRODUTOSID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarFamiliaProduto = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFamiliaProduto.ListarFamiliaProduto]", _
            Err.Description
End Function

Public Sub ExcluirFamiliaProduto(lngFAMILIAPRODUTOSID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui FAMÍLIA DE PRODUTOS
  strSql = "DELETE FROM FAMILIAPRODUTOS WHERE FAMILIAPRODUTOS.PKID = " & _
    Formata_Dados(lngFAMILIAPRODUTOSID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFamiliaProduto.ExcluirFAMILIAPRODUTO]", _
            Err.Description
End Sub


