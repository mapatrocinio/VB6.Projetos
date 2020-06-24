VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLinha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarLinha(lngTIPOLINHAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPO_LINHA.* " & _
    " FROM TIPO_LINHA " & _
    " WHERE TIPO_LINHA.PKID = " & _
    Formata_Dados(lngTIPOLINHAID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarLinha = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsLinha.SelecionarLinha]"
End Function
Public Sub ExcluirLinha(lngTIPOLINHAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM TIPO_LINHA " & _
    " WHERE PKID = " & lngTIPOLINHAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsLinha.ExcluirLinha]"
End Sub

Public Sub AlterarLinha(strNome As String, _
                        lngTIPOLINHAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE TIPO_LINHA SET " & _
    " NOME = " & Formata_Dados(strNome, tpDados_Texto) & _
    " WHERE PKID = " & Formata_Dados(lngTIPOLINHAID, tpDados_Longo)
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsLinha.AlterarLinha]"
End Sub
Public Sub IncluirLinha(strNome As String)
  On Error GoTo trata
  '
  Dim strSql          As String
  '
  strSql = "INSERT INTO TIPO_LINHA (" & _
    " NOME) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsLinha.IncluirLinha]"
End Sub

