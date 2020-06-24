VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirUsuario(strUsuario As String, _
                          strNivel As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO CONTROLEACESSO (USUARIO, NIVEL) VALUES (" & _
    Formata_Dados(strUsuario, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strNivel, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsUsuario.InserirUsuario]", _
            Err.Description
End Sub
Public Sub AlterarSenha(lngCONTROLEACESSOID As Long, _
                        strSenha As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE CONTROLEACESSO SET SENHA = " & Formata_Dados(strSenha, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngCONTROLEACESSOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsUsuario.AlterarSenha]", _
            Err.Description
End Sub

Public Sub AlterarUsuario(lngCONTROLEACESSOID As Long, _
                          strUsuario As String, _
                          strNivel As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE CONTROLEACESSO SET USUARIO = " & Formata_Dados(strUsuario, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "NIVEL = " & Formata_Dados(strNivel, tpDados_Texto, tpNulo_Aceita) & " " & _
    "WHERE PKID = " & lngCONTROLEACESSOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsUsuario.AlterarUsuario]", _
            Err.Description
End Sub

Public Function ListarUsuario(lngCONTROLEACESSOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  strSql = "SELECT CONTROLEACESSO.*, case CONTROLEACESSO.NIVEL when 'ADM' then 'ADMINISTRADOR' when 'DIR' then 'DIRETOR' when 'FIN' then 'FINANCEIRO' when 'GER' then 'GERENTE' when 'POR' then 'PORTARIA' when 'REC' then 'RECEPÇÃO' when 'EST' then 'ESTOQUISTA' else '' end AS DESCNIVEL "
  
  strSql = strSql & " FROM CONTROLEACESSO " & _
    "WHERE CONTROLEACESSO.PKID = " & _
    Formata_Dados(lngCONTROLEACESSOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarUsuario = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsUsuario.ListarUsuario]", _
            Err.Description
End Function

Public Sub ExlcuirUsuario(lngCONTROLEACESSOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM CONTROLEACESSO WHERE PKID = " & _
    Formata_Dados(lngCONTROLEACESSOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsUsuario.ListarUsuario]", _
            Err.Description
End Sub
