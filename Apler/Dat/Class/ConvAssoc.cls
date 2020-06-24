VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConvAssoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirConvAssoc(lngConvAssocId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM TAB_CONVASSOC " & _
    "WHERE TAB_CONVASSOC.PKID = " & _
    Formata_Dados(lngConvAssocId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConvAssoc.ExcluirConvAssoc]", _
            Err.Description
End Function

Public Function SelecionarConvAssocPeloPkid(lngConvAssocId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TAB_CONVASSOC.*, PLANOCONVENIO.NOME AS DESCR_PLANOCONVENIO " & _
    "FROM TAB_CONVASSOC LEFT JOIN PLANOCONVENIO ON PLANOCONVENIO.PKID = TAB_CONVASSOC.PLANOCONVENIOID " & _
    "WHERE TAB_CONVASSOC.PKID = " & _
    Formata_Dados(lngConvAssocId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarConvAssocPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConvAssoc.SelecionarConvAssocPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarConvAssoc(lngPKID As Long, _
                            lngPLANOCONVENIOID As Long, _
                            strDataInicio As String, _
                            strDataTermino As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE TAB_CONVASSOC SET " & _
    "PLANOCONVENIOID = " & Formata_Dados(lngPLANOCONVENIOID, tpDados_Longo) & ", " & _
    "DATAINICIO = " & Formata_Dados(strDataInicio, tpDados_DataHora) & ", " & _
    "DATATERMINO = " & Formata_Dados(strDataTermino, tpDados_DataHora) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConvAssoc.AlterarConvAssoc]", _
            Err.Description
End Sub

Public Sub InserirConvAssoc(lngASSOCIADOID As Long, _
                            lngPLANOCONVENIOID As Long, _
                            strDataInicio As String, _
                            strDataTermino As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO TAB_CONVASSOC (ASSOCIADOID, PLANOCONVENIOID, DATAINICIO, DATATERMINO) VALUES (" & _
    Formata_Dados(lngASSOCIADOID, tpDados_Longo) & ", " & _
    Formata_Dados(lngPLANOCONVENIOID, tpDados_Longo) & ", " & _
    Formata_Dados(strDataInicio, tpDados_DataHora) & ", " & _
    Formata_Dados(strDataTermino, tpDados_DataHora) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConvAssoc.InserirConvAssoc]", _
            Err.Description
End Sub

Public Function SelecionarConvAssoc(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TAB_CONVASSOC.* " & _
    "FROM TAB_CONVASSOC " & _
    "WHERE TAB_CONVASSOC.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarConvAssoc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsConvAssoc.SelecionarConvAssoc]", _
            Err.Description
End Function

