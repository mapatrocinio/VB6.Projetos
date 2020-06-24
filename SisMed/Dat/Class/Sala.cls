VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public Function ExcluirSala(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM SALA " & _
    "WHERE SALA.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSala.ExcluirSala]", _
            Err.Description
End Function


Public Function SelecionarSalaPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT SALA.*, PREDIO.NOME AS NOME_PREDIO " & _
    "FROM SALA LEFT JOIN PREDIO ON PREDIO.PKID = SALA.PREDIOID " & _
    "WHERE SALA.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarSalaPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSala.SelecionarSalaPeloPkid]", _
            Err.Description
End Function

Public Function SelecionarSalaPeloNumero(strNumero As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT SALA.*, PREDIO.NOME AS NOME_PREDIO " & _
    "FROM SALA LEFT JOIN PREDIO ON PREDIO.PKID = SALA.PREDIOID " & _
    "WHERE SALA.NUMERO = " & _
    Formata_Dados(strNumero, tpDados_Texto)
  Set SelecionarSalaPeloNumero = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSala.SelecionarSalaPeloNome]", _
            Err.Description
End Function


Public Sub AlterarSala(lngPKID As Long, _
                       lngPREDIOID As Long, _
                       strNumero As String, _
                       strAndar As String, _
                       strTelefone As String, _
                       strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE SALA SET " & _
    "PREDIOID = " & Formata_Dados(IIf(lngPREDIOID = 0, "", lngPREDIOID), tpDados_Longo) & ", " & _
    "NUMERO = " & Formata_Dados(strNumero, tpDados_Texto) & ", " & _
    "ANDAR = " & Formata_Dados(strAndar, tpDados_Texto) & ", " & _
    "TELEFONE = " & Formata_Dados(strTelefone, tpDados_Texto) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSala.AlterarSala]", _
            Err.Description
End Sub

Public Sub InserirSala(lngPREDIOID As Long, _
                       strNumero As String, _
                       strAndar As String, _
                       strTelefone As String, _
                       strStatus As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO SALA (PREDIOID, NUMERO, ANDAR, TELEFONE, STATUS) VALUES (" & _
    Formata_Dados(IIf(lngPREDIOID = 0, "", lngPREDIOID), tpDados_Longo) & ", " & _
    Formata_Dados(strNumero, tpDados_Texto) & ", " & _
    Formata_Dados(strAndar, tpDados_Texto) & ", " & _
    Formata_Dados(strTelefone, tpDados_Texto) & ", " & _
    Formata_Dados(strStatus, tpDados_Texto) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSala.InserirSala]", _
            Err.Description
End Sub

'''Public Function SelecionarSala(lngPKID As Long) As ADODB.Recordset
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "SELECT SALA.*, CONTROLEACESSO.USUARIO " & _
'''    "FROM SALA LEFT JOIN CONTROLEACESSO ON CONTROLEACESSO.PKID = SALA.CONTROLEACESSOID " & _
'''    "WHERE SALA.PKID = " & _
'''    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
'''  Set SelecionarSala = RunSPReturnRS(strSql)
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsSala.SelecionarSala]", _
'''            Err.Description
'''End Function



