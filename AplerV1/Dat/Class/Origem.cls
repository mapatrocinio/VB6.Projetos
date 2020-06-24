VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrigem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirOrigem(strDescricao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO ORIGEM (DESCRICAO) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsOrigem.InserirOrigem]", _
            Err.Description
End Sub

Public Sub AlterarOrigem(lngORIGEMID As Long, _
                            strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE ORIGEM SET DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngORIGEMID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsOrigem.AlterarOrigem]", _
            Err.Description
End Sub

Public Function ListarOrigem(lngORIGEMID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ORIGEM.* " & _
    "FROM ORIGEM " & _
    "WHERE ORIGEM.PKID = " & _
    Formata_Dados(lngORIGEMID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarOrigem = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsOrigem.ListarOrigem]", _
            Err.Description
End Function

Public Sub ExlcuirOrigem(lngORIGEMID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM ORIGEM WHERE PKID = " & _
    Formata_Dados(lngORIGEMID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsOrigem.ExlcuirOrigem]", _
            Err.Description
End Sub


Public Function ListarOrigemPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ORIGEM.* " & _
    "FROM ORIGEM " & _
    "WHERE ORIGEM.DESCRICAO = " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarOrigemPelaDesc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsOrigem.ListarOrigemPelaDesc]", _
            Err.Description
End Function

