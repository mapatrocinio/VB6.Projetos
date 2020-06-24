VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public Sub TratarStatus(lngPACOTEID As Long, _
                        strStatus As String, _
                        strDataFimPacote As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE PACOTE SET " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & ", " & _
    "DATATERMINO = " & Formata_Dados(strDataFimPacote, tpDados_DataHora) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPACOTEID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPacote.TratarStatus]", _
            Err.Description
End Sub


Public Function ExcluirPacote(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  'Delete PACOTEPACOTE
  strSql = "DELETE FROM PACOTEPACOTE " & _
    "WHERE PACOTEPACOTE.PACOTEID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  'Delete PACOTE
  strSql = "DELETE FROM PACOTE " & _
    "WHERE PACOTE.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPacote.ExcluirPacote]", _
            Err.Description
End Function


Public Function SelecionarPacotePeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PACOTE.*, " & _
    " PESSOA.NOME AS DESC_MOTORISTA " & _
    " FROM PACOTE " & _
    " LEFT JOIN MOTORISTA ON PACOTE.MOTORISTAID = MOTORISTA.PESSOAID " & _
    " LEFT JOIN PESSOA ON PESSOA.PKID = MOTORISTA.PESSOAID " & _
    "WHERE PACOTE.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarPacotePeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPacote.SelecionarPacotePeloPkid]", _
            Err.Description
End Function

Public Sub AlterarPacote(lngPKID As Long, _
                          strDataHoraInicio As String, _
                          strDataHoraTermino As String, _
                          lngMOTORISTAID As Long, _
                          strValor As String, _
                          strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE PACOTE SET " & _
    "DATAINICIO = " & Formata_Dados(strDataHoraInicio, tpDados_DataHora) & ", " & _
    "DATATERMINO = " & Formata_Dados(strDataHoraTermino, tpDados_DataHora) & ", " & _
    "MOTORISTAID = " & Formata_Dados(lngMOTORISTAID, tpDados_Longo) & ", " & _
    "VALOR = " & Formata_Dados(strValor, tpDados_Moeda) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPacote.AlterarPacote]", _
            Err.Description
End Sub

Public Sub InserirPacote(lngPKID As Long, _
                         strDataHoraInicio As String, _
                         strDataHoraTermino As String, _
                         lngMOTORISTAID As Long, _
                         strValor As String, _
                         strStatus As String)

  On Error GoTo trata
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO PACOTE (DATAINICIO, DATATERMINO, MOTORISTAID, " & _
    "VALOR, STATUS) VALUES (" & _
    Formata_Dados(strDataHoraInicio, tpDados_DataHora) & ", " & _
    Formata_Dados(strDataHoraTermino, tpDados_DataHora) & ", " & _
    Formata_Dados(lngMOTORISTAID, tpDados_Longo) & ", " & _
    Formata_Dados(strValor, tpDados_Moeda) & ", " & _
    Formata_Dados(strStatus, tpDados_Texto) & ");"
  '
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngPKID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPacote.InserirPacote]", _
            Err.Description
End Sub

