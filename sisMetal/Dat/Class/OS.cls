VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirOS(lngOSID As Long)
  On Error GoTo trata
  
  '
  Dim strSql As String

'  strSql = "UPDATE OS " & _
'    " SET CANCELADO = " & Formata_Dados(strStatus, tpDados_Texto) & _
'    ", DATA_CANCELAMENTO = " & Formata_Dados(strData, tpDados_DataHora) & _
'    " WHERE PKID = " & _
'    Formata_Dados(lngOSID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsOS.ExcluirOS]", _
            Err.Description
End Sub

Public Function ListarOS(lngOSID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT OS.*, FORNECEDOR.NOME AS NOME_FORNECEDOR, FABRICA.NOME AS NOME_FABRICA, COR.NOME AS NOME_COR " & _
    " FROM OS LEFT JOIN LOJA AS FORNECEDOR ON OS.FORNECEDORID = FORNECEDOR.PKID " & _
    " LEFT JOIN LOJA AS FABRICA ON OS.FABRICAID = FABRICA.PKID " & _
    " LEFT JOIN COR ON COR.PKID = OS.CORID " & _
    " WHERE OS.PKID = " & _
    Formata_Dados(lngOSID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarOS = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsOS.ListarOS]", _
            Err.Description
End Function


Public Sub InserirOS(lngOSID As Long, _
                     lngFORNECEDORID As Long, _
                     lngCORID As Long, _
                     lngFABRICAID As Long, _
                     strNumero As String, _
                     strNF As String, _
                     strData As String, _
                     strPesoBruto As String, _
                     strPesoLiquido As String, _
                     strValorMetal As String, _
                     strValor As String, _
                     strOperador As String, _
                     strConferente As String)
  On Error GoTo trata
  '
  Dim objRs As ADODB.Recordset
  Dim strSql As String
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO OS (NUMERO, NF, DATA, FABRICAID, CORID, FORNECEDORID, " & _
      " PESOBRUTO, PESOLIQUIDO, VALORMETAL, VALOR, OPERADOR, CONFERENTE, STATUS) VALUES (" & _
    Formata_Dados(strNumero, tpDados_Longo) & ", " & _
    Formata_Dados(strNF, tpDados_Texto) & ", " & _
    Formata_Dados(strData, tpDados_DataHora) & ", " & _
    Formata_Dados(IIf(lngFABRICAID = 0, "", lngFABRICAID), tpDados_Longo) & ", " & _
    Formata_Dados(IIf(lngCORID = 0, "", lngCORID), tpDados_Longo) & ", " & _
    Formata_Dados(IIf(lngFORNECEDORID = 0, "", lngFORNECEDORID), tpDados_Longo) & ", " & _
    Formata_Dados(strPesoBruto, tpDados_Moeda) & ", " & _
    Formata_Dados(strPesoLiquido, tpDados_Moeda) & ", " & _
    Formata_Dados(strValorMetal, tpDados_Moeda) & ", " & _
    Formata_Dados(strValor, tpDados_Moeda) & ", " & _
    Formata_Dados(strOperador, tpDados_Texto) & ", " & _
    Formata_Dados(strConferente, tpDados_Texto) & ", " & _
    Formata_Dados("I", tpDados_Texto) & ");"
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngOSID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsOS.InserirOS]", _
            Err.Description
End Sub

Public Sub AlterarOS(lngOSID As Long, _
                     lngFORNECEDORID As Long, _
                     lngCORID As Long, _
                     lngFABRICAID As Long, _
                     strNumero As String, _
                     strNF As String, _
                     strPesoBruto As String, _
                     strPesoLiquido As String, _
                     strValorMetal As String, _
                     strValor As String, _
                     strOperador As String, _
                     strConferente As String)
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE OS SET " & _
    "FORNECEDORID = " & Formata_Dados(IIf(lngFORNECEDORID = 0, "", lngFORNECEDORID), tpDados_Longo) & _
    ", CORID = " & Formata_Dados(IIf(lngCORID = 0, "", lngCORID), tpDados_Longo) & _
    ", FABRICAID = " & Formata_Dados(IIf(lngFABRICAID = 0, "", lngFABRICAID), tpDados_Longo) & _
    ", NUMERO = " & Formata_Dados(strNumero, tpDados_Longo) & _
    ", NF = " & Formata_Dados(strNF, tpDados_Texto) & _
    ", PESOBRUTO = " & Formata_Dados(strPesoBruto, tpDados_Moeda) & _
    ", PESOLIQUIDO = " & Formata_Dados(strPesoLiquido, tpDados_Moeda) & _
    ", VALORMETAL = " & Formata_Dados(strValorMetal, tpDados_Moeda) & _
    ", VALOR = " & Formata_Dados(strValor, tpDados_Moeda) & _
    ", OPERADOR = " & Formata_Dados(strOperador, tpDados_Texto) & _
    ", CONFERENTE = " & Formata_Dados(strConferente, tpDados_Texto) & _
    " WHERE OS.PKID = " & Formata_Dados(lngOSID, tpDados_Longo)
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsOS.AlterarOS]", _
            Err.Description
End Sub

