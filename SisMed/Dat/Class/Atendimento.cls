VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAtendimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarAtendimentoPeloPkid(lngATENDIMENTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ATENDIMENTO.* " & _
    "FROM ATENDIMENTO " & _
    "WHERE ATENDIMENTO.PKID = " & _
    Formata_Dados(lngATENDIMENTOID, tpDados_Longo)
  Set SelecionarAtendimentoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAtendimento.SelecionarAtendimentoPeloPkid]", _
            Err.Description
End Function

Public Sub AlterarAtendimento(lngPKID As Long, _
                              strPathArquivo As String, _
                              strNomeArquivo As String, _
                              strDescricao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE ATENDIMENTO SET " & _
    "PATHARQUIVO = " & Formata_Dados(strPathArquivo, tpDados_Texto) & ", " & _
    "NOMEARQUIVO = " & Formata_Dados(strNomeArquivo, tpDados_Texto) & ", " & _
    "DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAtendimento.AlterarAtendimento]", _
            Err.Description
End Sub

Public Sub InserirAtendimento(lngGRID As Long, _
                              strData As String, _
                              strTrabalhaComScaner As String, _
                              strPathArquivo As String, _
                              strNomeArquivo As String, _
                              strDescricao As String, _
                              strDataImportacao As String, _
                              lngPRONTUARIOID As Long)
  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO ATENDIMENTO (GRID, DATA, INDSCANER, PATHARQUIVO, NOMEARQUIVO, DESCRICAO, DATAIMPORTACAO, PRONTUARIOID) VALUES (" & _
    " " & Formata_Dados(IIf(lngGRID = 0, "", lngGRID), tpDados_Longo) & _
    ", " & Formata_Dados(strData, tpDados_DataHora) & _
    ", " & Formata_Dados(strTrabalhaComScaner, tpDados_Texto) & _
    ", " & Formata_Dados(strPathArquivo, tpDados_Texto) & _
    ", " & Formata_Dados(strNomeArquivo, tpDados_Texto) & _
    ", " & Formata_Dados(strDescricao, tpDados_Texto) & _
    ", " & Formata_Dados(strDataImportacao, tpDados_DataHora) & _
    ", " & Formata_Dados(IIf(lngPRONTUARIOID = 0, "", lngPRONTUARIOID), tpDados_Longo) & ")"
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAtendimento.InserirAtendimento]", _
            Err.Description
End Sub


