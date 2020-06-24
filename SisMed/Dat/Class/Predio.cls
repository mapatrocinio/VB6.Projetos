VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPredio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirPredio(strNome As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO PREDIO (NOME) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPredio.InserirPredio]", _
            Err.Description
End Sub

Public Sub AlterarPredio(lngPREDIOID As Long, _
                         strNome As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE PREDIO SET NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngPREDIOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPredio.AlterarPredio]", _
            Err.Description
End Sub

Public Function ListarPredio(lngPREDIOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PREDIO.* " & _
    "FROM PREDIO " & _
    "WHERE PREDIO.PKID = " & _
    Formata_Dados(lngPREDIOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarPredio = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPredio.ListarPredio]", _
            Err.Description
End Function

Public Sub ExlcuirPredio(lngPREDIOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM PREDIO WHERE PKID = " & _
    Formata_Dados(lngPREDIOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPredio.ExlcuirPredio]", _
            Err.Description
End Sub
