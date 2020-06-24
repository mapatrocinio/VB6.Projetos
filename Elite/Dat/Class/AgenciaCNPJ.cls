VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAgenciaCNPJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirAgenciaCNPJ(lngAgenciaCNPJId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM AGENCIACNPJ " & _
    "WHERE AGENCIACNPJ.PKID = " & _
    Formata_Dados(lngAgenciaCNPJId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAgenciaCNPJ.ExcluirAgenciaCNPJ]", _
            Err.Description
End Function

Public Function SelecionarAgenciaCNPJPeloPkid(lngAgenciaCNPJId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT AGENCIACNPJ.* " & _
    "FROM AGENCIACNPJ " & _
    "WHERE AGENCIACNPJ.PKID = " & _
    Formata_Dados(lngAgenciaCNPJId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarAgenciaCNPJPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAgenciaCNPJ.SelecionarAgenciaCNPJPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarAgenciaCNPJ(lngPKID As Long, _
                              strCNPJ As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE AGENCIACNPJ SET " & _
    "CNPJ = " & Formata_Dados(strCNPJ, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAgenciaCNPJ.AlterarAgenciaCNPJ]", _
            Err.Description
End Sub

Public Sub InserirAgenciaCNPJ(lngAGENCIAID As Long, _
                              strCNPJ As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO AGENCIACNPJ (AGENCIAID, CNPJ) VALUES (" & _
    Formata_Dados(lngAGENCIAID, tpDados_Longo) & ", " & _
    Formata_Dados(strCNPJ, tpDados_Texto) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAgenciaCNPJ.InserirAgenciaCNPJ]", _
            Err.Description
End Sub

Public Function SelecionarAgenciaCNPJ(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT AGENCIACNPJ.* " & _
    "FROM AGENCIACNPJ " & _
    "WHERE AGENCIACNPJ.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarAgenciaCNPJ = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAgenciaCNPJ.SelecionarAgenciaCNPJ]", _
            Err.Description
End Function

