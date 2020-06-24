VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLivro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarLivro(lngLivroId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT LIVRO.*, BANCO.NOME FROM LIVRO " & _
    " LEFT JOIN BANCO ON BANCO.PKID = LIVRO.BANCOID " & _
    " WHERE LIVRO.PKID = " & _
    Formata_Dados(lngLivroId, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarLivro = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsLivro.SelecionarLivro]"
End Function
Public Sub ExcluirLivro(lngLivroId As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM LIVRO " & _
    " WHERE PKID = " & lngLivroId
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsLivro.ExcluirLivro]"
End Sub

Public Sub AlterarLivro(lngLivroId As Long, _
                        lngBANCOID As Long, _
                        strConta As String, _
                        strAgencia As String, _
                        strNumeroLivro As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE LIVRO SET " & _
    " BANCOID = " & Formata_Dados(lngBANCOID, tpDados_Longo, tpNulo_Aceita) & _
    ", CONTA = " & Formata_Dados(strConta, tpDados_Texto, tpNulo_Aceita) & _
    ", AGENCIA = " & Formata_Dados(strAgencia, tpDados_Texto, tpNulo_Aceita) & _
    ", NUMEROLIVRO = " & Formata_Dados(strNumeroLivro, tpDados_Texto, tpNulo_Aceita) & _
    " WHERE PKID = " & lngLivroId
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsLivro.AlterarLivro]"
End Sub
Public Sub IncluirLivro(lngBANCOID As Long, _
                        strConta As String, _
                        strAgencia As String, _
                        strNumeroLivro As String, _
                        lngParceiroId As Long)
  On Error GoTo trata
  '
  Dim strSql          As String
  '
  strSql = "INSERT INTO LIVRO (" & _
    " BANCOID, CONTA, AGENCIA, NUMEROLIVRO, PARCEIROID) VALUES (" & _
    Formata_Dados(lngBANCOID, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strConta, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strAgencia, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strNumeroLivro, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(lngParceiroId, tpDados_Longo) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsLivro.IncluirLivro]"
End Sub

