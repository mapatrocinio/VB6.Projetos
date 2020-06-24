VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaptador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit

Public Function ExcluirCaptador(lngCaptadorId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM CAPTADOR " & _
    "WHERE CAPTADOR.PKID = " & _
    Formata_Dados(lngCaptadorId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCaptador.ExcluirCaptador]", _
            Err.Description
End Function

Public Function SelecionarCaptadorPeloPkid(lngCaptadorId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CAPTADOR.* " & _
    "FROM CAPTADOR " & _
    "WHERE CAPTADOR.PKID = " & _
    Formata_Dados(lngCaptadorId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarCaptadorPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCaptador.SelecionarCaptadorPeloPkid]", _
            Err.Description
End Function

Public Sub AlterarCaptador(lngPKID As Long, _
                           strNome As String, _
                           strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE CAPTADOR SET " & _
    "NOME = " & Formata_Dados(strNome, tpDados_Texto) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCaptador.AlterarCaptador]", _
            Err.Description
End Sub

Public Sub InserirCaptador(strNome As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO CAPTADOR (NOME, STATUS) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto) & ", " & _
    Formata_Dados("A", tpDados_Texto) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCaptador.InserirCaptador]", _
            Err.Description
End Sub

Public Function SelecionarCaptador(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CAPTADOR.* " & _
    "FROM CAPTADOR " & _
    "WHERE CAPTADOR.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarCaptador = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCaptador.SelecionarCaptador]", _
            Err.Description
End Function



