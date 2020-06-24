VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEspecialidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirEspecialidade(strEspecialidade As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO ESPECIALIDADE (ESPECIALIDADE) VALUES (" & _
    Formata_Dados(strEspecialidade, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEspecialidade.InserirEspecialidade]", _
            Err.Description
End Sub

Public Sub AlterarEspecialidade(lngESPECIALIDADEID As Long, _
                         strEspecialidade As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE ESPECIALIDADE SET ESPECIALIDADE = " & Formata_Dados(strEspecialidade, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngESPECIALIDADEID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEspecialidade.AlterarEspecialidade]", _
            Err.Description
End Sub

Public Function ListarEspecialidade(lngESPECIALIDADEID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ESPECIALIDADE.* " & _
    "FROM ESPECIALIDADE " & _
    "WHERE ESPECIALIDADE.PKID = " & _
    Formata_Dados(lngESPECIALIDADEID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarEspecialidade = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEspecialidade.ListarEspecialidade]", _
            Err.Description
End Function

Public Sub ExlcuirEspecialidade(lngESPECIALIDADEID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM ESPECIALIDADE WHERE PKID = " & _
    Formata_Dados(lngESPECIALIDADEID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEspecialidade.ExlcuirEspecialidade]", _
            Err.Description
End Sub
