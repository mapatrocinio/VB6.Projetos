VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrestador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarPrestadorPeloPkid(lngPRESTADORID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PRESTADOR.*, FUNCAO.FUNCAO AS DESCFUNCAO " & _
    "FROM PRESTADOR LEFT JOIN FUNCAO ON FUNCAO.PKID = PRESTADOR.FUNCAOID " & _
    "WHERE PRESTADOR.PRONTUARIOID = " & _
    Formata_Dados(lngPRESTADORID, tpDados_Longo)
  Set SelecionarPrestadorPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPrestador.SelecionarPrestadorPeloPkid]", _
            Err.Description
End Function

Public Sub AlterarPrestador(lngPRONTUARIOID As Long, _
                            lngFUNCAOID As Long, _
                            strCRM As String, _
                            blnTrabalhaComScaner As Boolean, _
                            strIndAceitaCheque As String, _
                            strExcluido As String, _
                            strDataDesativa As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE PRESTADOR SET " & _
    "CRM = " & Formata_Dados(strCRM, tpDados_Texto) & ", " & _
    "TRABCOMSCANER = " & Formata_Dados(blnTrabalhaComScaner, tpDados_Boolean) & ", " & _
    "FUNCAOID = " & Formata_Dados(lngFUNCAOID, tpDados_Longo) & ", " & _
    "INDACEITACHEQUE = " & Formata_Dados(strIndAceitaCheque, tpDados_Texto) & ", " & _
    "INDEXCLUIDO = " & Formata_Dados(strExcluido, tpDados_Texto) & ", " & _
    "DATADESATIVACAO= " & Formata_Dados(strDataDesativa, tpDados_DataHora) & " " & _
    "WHERE PRONTUARIOID = " & Formata_Dados(lngPRONTUARIOID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPrestador.AlterarPrestador]", _
            Err.Description
End Sub

Public Sub InserirPrestador(lngPRONTUARIOID As Long, _
                            lngFUNCAOID As Long, _
                            strCRM As String, _
                            blnTrabalhaComScaner As Boolean, _
                            strIndAceitaCheque As String, _
                            strExcluido As String)
  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO PRESTADOR (PRONTUARIOID, FUNCAOID, CRM, TRABCOMSCANER, INDACEITACHEQUE, INDEXCLUIDO) VALUES (" & _
    " " & Formata_Dados(lngPRONTUARIOID, tpDados_Longo) & _
    ", " & Formata_Dados(lngFUNCAOID, tpDados_Longo) & _
    ", " & Formata_Dados(strCRM, tpDados_Texto) & _
    ", " & Formata_Dados(blnTrabalhaComScaner, tpDados_Boolean) & _
    ", " & Formata_Dados(strIndAceitaCheque, tpDados_Texto) & _
    ", " & Formata_Dados(strExcluido, tpDados_Texto) & ")"
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPrestador.InserirPrestador]", _
            Err.Description
End Sub


