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

Public Function SelecionarPacotePeloPkid(lngSERVICOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objPacote As datElite.clsPacote
  Set objPacote = New datElite.clsPacote
  
  'Listar Pacote
  Set SelecionarPacotePeloPkid = objPacote.SelecionarPacotePeloPkid(lngSERVICOID)
  '
  Set objPacote = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Sub TratarStatus(lngPACOTEID As Long)
  On Error GoTo trata
  Dim objGeral                  As datElite.clsGeral
  Dim objPacote                 As datElite.clsPacote
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim strStatus                 As String
  Dim strDataFimPacote          As String
  
  '
  'Verifica se ainda possui serviços em aberto para pacote
  Set objGeral = New datElite.clsGeral
  strSql = "SELECT SERVICO.PKID " & _
           "FROM PACOTESERVICO " & _
           " INNER JOIN SERVICO ON SERVICO.PKID = PACOTESERVICO.SERVICOID " & _
           " WHERE PACOTESERVICO.PACOTEID = " & Formata_Dados(lngPACOTEID, tpDados_Longo) & _
           " AND PACOTESERVICO.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
           " AND SERVICO.STATUS = " & Formata_Dados("I", tpDados_Texto)

  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    strStatus = "I"
  Else
    strStatus = "C"
  End If
  objRs.Close
  Set objRs = Nothing
  'Verifica data final do pacote
  strDataFimPacote = ""
  If strStatus = "C" Then
    'Obtém ultima data do serviço para fechar o pacote
    strSql = "SELECT MAX(SERVICO.DATAFINAL) AS DATAFINAL " & _
             "FROM PACOTESERVICO " & _
             " INNER JOIN SERVICO ON SERVICO.PKID = PACOTESERVICO.SERVICOID " & _
             " WHERE PACOTESERVICO.PACOTEID = " & Formata_Dados(lngPACOTEID, tpDados_Longo) & _
             " AND PACOTESERVICO.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
             " AND SERVICO.STATUS = " & Formata_Dados("F", tpDados_Texto)
    
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      strDataFimPacote = Format(objRs.Fields("DATAFINAL").Value, "DD/MM/YYYY hh:mm")
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  
  Set objGeral = Nothing
  '
  Set objPacote = New datElite.clsPacote
  'Atualiza status do pacote
  objPacote.TratarStatus lngPACOTEID, _
                         strStatus, _
                         strDataFimPacote
  '
  Set objPacote = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ExcluirPacote(lngPACOTEID As Long)
  On Error GoTo trata
  Dim objPacote As datElite.clsPacote
  Set objPacote = New datElite.clsPacote
  
  'Excluir Pacote
  objPacote.ExcluirPacote lngPACOTEID
  '
  Set objPacote = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirPacote(lngPKID As Long, _
                         strDataHoraInicio As String, _
                         strDataHoraTermino As String, _
                         lngMOTORISTAID As Long, _
                         strValor As String)

  On Error GoTo trata
  Dim objPacote     As datElite.clsPacote
  Dim strStatus As String
  '
  'Status Inicial
  strStatus = "I"
  Set objPacote = New datElite.clsPacote
  'Inserir Pacote
  objPacote.InserirPacote lngPKID, _
                          strDataHoraInicio, _
                          strDataHoraTermino, _
                          lngMOTORISTAID, _
                          strValor, _
                          strStatus

  '
  Set objPacote = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarPacote(lngPKID As Long, _
                         strDataHoraInicio As String, _
                         strDataHoraTermino As String, _
                         lngMOTORISTAID As Long, _
                         strValor As String, _
                         strStatus As String)

  On Error GoTo trata
  Dim objPacote     As datElite.clsPacote
  '
  Set objPacote = New datElite.clsPacote
  'Atualizar Pacote
  objPacote.AlterarPacote lngPKID, _
                          strDataHoraInicio, _
                          strDataHoraTermino, _
                          lngMOTORISTAID, _
                          strValor, _
                          strStatus
  '
  Set objPacote = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

