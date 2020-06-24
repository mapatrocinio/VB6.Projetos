VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarGrupo(lngGRUPOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT GRUPO.* " & _
    " FROM GRUPO " & _
    " WHERE GRUPO.PKID = " & _
    Formata_Dados(lngGRUPOID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarGrupo = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsGrupo.SelecionarGrupo]"
End Function
Public Sub ExcluirGrupo(lngGRUPOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM GRUPO " & _
    " WHERE PKID = " & lngGRUPOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsGrupo.ExcluirGrupo]"
End Sub

Public Sub AlterarGrupo(strNome As String, _
                        strReferencia As String, _
                        lngGRUPOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE GRUPO SET " & _
    " NOME = " & Formata_Dados(strNome, tpDados_Texto) & _
    ", REFERENCIA = " & Formata_Dados(strReferencia, tpDados_Texto) & _
    " WHERE PKID = " & Formata_Dados(lngGRUPOID, tpDados_Longo)
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsGrupo.AlterarGrupo]"
End Sub
Public Sub IncluirGrupo(strNome As String, _
                        strReferencia As String)
  On Error GoTo trata
  '
  Dim strSql          As String
  '
  strSql = "INSERT INTO GRUPO (" & _
    " NOME, REFERENCIA) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto) & _
    ", " & Formata_Dados(strReferencia, tpDados_Texto) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsGrupo.IncluirGrupo]"
End Sub

