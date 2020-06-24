VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOSFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirOSFinal(lngOSFinalID As Long)
  On Error GoTo trata
  
  '
  Dim strSql As String

'  strSql = "UPDATE OSFinal " & _
'    " SET CANCELADO = " & Formata_Dados(strStatus, tpDados_Texto) & _
'    ", DATA_CANCELAMENTO = " & Formata_Dados(strData, tpDados_DataHora) & _
'    " WHERE PKID = " & _
'    Formata_Dados(lngOSFinalID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsOSFinal.ExcluirOSFinal]", _
            Err.Description
End Sub

Public Function ListarOSFinal(lngOSFinalID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT OS_FINAL.*, COR.NOME AS NOME_COR " & _
    " FROM OS_FINAL LEFT JOIN COR ON COR.PKID = OS_FINAL.CORID " & _
    " WHERE OS_FINAL.PKID = " & _
    Formata_Dados(lngOSFinalID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarOSFinal = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsOSFinal.ListarOSFinal]", _
            Err.Description
End Function


Public Sub InserirOSFinal(lngOSFinalID As Long, _
                          lngCORID As Long, _
                          strComplemento As String, _
                          lngOSID As Long, _
                          strData As String, _
                          strPesoBruto As String, _
                          strPesoLiquido As String, _
                          strValorMetal As String, _
                          strValor As String, _
                          strValorFrete As String, _
                          strOperador As String, _
                          strConferente As String)
  On Error GoTo trata
  '
  Dim objRs As ADODB.Recordset
  Dim strSql As String
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO OS_FINAL (COMPLEMENTO, OSID, DATA, CORID, " & _
      " PESOBRUTO, PESOLIQUIDO, VALORMETAL, VALOR, VALORFRETE, OPERADOR, CONFERENTE) VALUES (" & _
    Formata_Dados(strComplemento, tpDados_Texto) & ", " & _
    Formata_Dados(lngOSID, tpDados_Longo) & ", " & _
    Formata_Dados(strData, tpDados_DataHora) & ", " & _
    Formata_Dados(IIf(lngCORID = 0, "", lngCORID), tpDados_Longo) & ", " & _
    Formata_Dados(strPesoBruto, tpDados_Moeda) & ", " & _
    Formata_Dados(strPesoLiquido, tpDados_Moeda) & ", " & _
    Formata_Dados(strValorMetal, tpDados_Moeda) & ", " & _
    Formata_Dados(strValor, tpDados_Moeda) & ", " & _
    Formata_Dados(strValorFrete, tpDados_Moeda) & ", " & _
    Formata_Dados(strOperador, tpDados_Texto) & ", " & _
    Formata_Dados(strConferente, tpDados_Texto) & ");"
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngOSFinalID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsOSFinal.InserirOSFinal]", _
            Err.Description
End Sub

Public Sub AlterarOSFinal(lngOSFinalID As Long, _
                          lngCORID As Long, _
                          strComplemento As String, _
                          strPesoBruto As String, _
                          strPesoLiquido As String, _
                          strValorMetal As String, _
                          strValor As String, _
                          strValorFrete As String, _
                          strOperador As String, _
                          strConferente As String)
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE OS_FINAL SET " & _
    " CORID = " & Formata_Dados(IIf(lngCORID = 0, "", lngCORID), tpDados_Longo) & _
    ", COMPLEMENTO = " & Formata_Dados(strComplemento, tpDados_Texto) & _
    ", PESOBRUTO = " & Formata_Dados(strPesoBruto, tpDados_Moeda) & _
    ", PESOLIQUIDO = " & Formata_Dados(strPesoLiquido, tpDados_Moeda) & _
    ", VALORMETAL = " & Formata_Dados(strValorMetal, tpDados_Moeda) & _
    ", VALOR = " & Formata_Dados(strValor, tpDados_Moeda) & _
    ", VALORFRETE = " & Formata_Dados(strValorFrete, tpDados_Moeda) & _
    ", OPERADOR = " & Formata_Dados(strOperador, tpDados_Texto) & _
    ", CONFERENTE = " & Formata_Dados(strConferente, tpDados_Texto) & _
    " WHERE OS_FINAL.PKID = " & Formata_Dados(lngOSFinalID, tpDados_Longo)
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsOSFinal.AlterarOSFinal]", _
            Err.Description
End Sub

