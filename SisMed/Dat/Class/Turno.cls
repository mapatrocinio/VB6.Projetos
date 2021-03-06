VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Sub InserirTurno(lngFuncionarioId As Long, _
                        strPeriodoId As String, _
                        strData As String, _
                        blnStatus As Boolean, _
                        strSequencial As String, _
                        strDiaDaSemana As String, _
                        strVrCaixaInicial As String, _
                        strSequencialFecha As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO TURNO (ProntuarioId, PeriodoId, Data, Status, SEQUENCIALTURNO, DIASDASEMANAID, VRCAIXAINICIAL, SEQUENCIALFECHA) VALUES (" & _
    Formata_Dados(lngFuncionarioId, tpDados_Longo, tpNulo_Aceita) & ", " & _
    Formata_Dados(strPeriodoId, tpDados_Longo, tpNulo_Aceita) & ", " & _
    Formata_Dados(strData, tpDados_DataHora, tpNulo_Aceita) & ", " & _
    Formata_Dados(blnStatus, tpDados_Boolean, tpNulo_Aceita) & ", " & _
    Formata_Dados(strSequencial, tpDados_Longo, tpNulo_Aceita) & ", " & _
    Formata_Dados(strDiaDaSemana, tpDados_Longo, tpNulo_Aceita) & ", " & _
    Formata_Dados(strVrCaixaInicial, tpDados_Moeda, tpNulo_Aceita) & ", " & _
    Formata_Dados(strSequencialFecha, tpDados_Longo, tpNulo_Aceita) & ");"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTurno.InserirTurno]", _
            Err.Description
End Sub

Public Sub FecharTurno(lngTurnoId As Long, _
                       blnStatus As Boolean, _
                       strDataFechamento As String)

  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE TURNO SET " & _
    " STATUS = " & Formata_Dados(blnStatus, tpDados_Boolean, tpNulo_Aceita) & _
    ", DTFECHAMENTO = " & Formata_Dados(strDataFechamento, tpDados_DataHora, tpNulo_Aceita) & _
    " WHERE PKID = " & Formata_Dados(lngTurnoId, tpDados_Longo, tpNulo_Aceita)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTurno.FecharTurno]", _
            Err.Description
End Sub

'Retorna 0 - para C�digo de Erro
Public Function RetornaCodTurnoCorrente() As Long
  '
  On Error GoTo trata
  
  Dim objRs As ADODB.Recordset
  Dim strSql  As String
  Dim lngRet As Long
  '
  strSql = "Select * from Turno Where Status = " & Formata_Dados(True, tpDados_Boolean)
  Set objRs = RunSPReturnRS(strSql)
  '
  If objRs.EOF Then
    'N�o h� turno corrente cadastrado
    lngRet = 0
  ElseIf objRs.RecordCount > 1 Then
    'h� mais de um turno corrente cadastrado
    lngRet = 0
  Else
    lngRet = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  RetornaCodTurnoCorrente = lngRet
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTurno.RetornaCodTurnoCorrente]", _
            Err.Description
End Function

Public Function CapturaTurno(ByVal lngTurnoId As Long) As ADODB.Recordset
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "SELECT * FROM TURNO WHERE PKID = " & Formata_Dados(lngTurnoId, tpDados_Longo, tpNulo_NaoAceita)
  Set CapturaTurno = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTurno.CapturaTurno]", _
            Err.Description
End Function

Public Function ListarTurno() As ADODB.Recordset
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "Select right('0' + convert(varchar(2), day(DATA)),2) + '/' + " & _
    "right('0' + convert(varchar(2), month(DATA)),2) + '/' + " & _
    "convert(varchar(4), year(DATA)) + ' ' + " & _
    "right('0' + convert(varchar(2), datepart(hh,DATA)),2) + ':' + " & _
    "right('0' + convert(varchar(2), datepart(n,DATA)),2) + " & _
    "' / ' + convert(varchar(10), DIASDASEMANA.DIADASEMANA) + ' - Per�odo ' + convert(varchar(10), Periodo) + ' de ' + inicio + ' as ' + termino AS DESCTURNO, TURNO.PKID  " & _
      "FROM (PERIODO INNER JOIN TURNO ON PERIODO.PKID = TURNO.PERIODOID) INNER JOIN DIASDASEMANA ON TURNO.DIADASEMANA = DIASDASEMANA.CODIGO " & _
      " ORDER BY DATA DESC"
        
  Set ListarTurno = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTurno.ListarTurno]", _
            Err.Description
End Function


Public Function ListarTurnoPorUnidade(strSqlWhere As String) As ADODB.Recordset
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "Select right('0' + convert(varchar(2), day(DATA)),2) + '/' + " & _
    "right('0' + convert(varchar(2), month(DATA)),2) + '/' + " & _
    "convert(varchar(4), year(DATA)) + ' ' + " & _
    "right('0' + convert(varchar(2), datepart(hh,DATA)),2) + ':' + " & _
    "right('0' + convert(varchar(2), datepart(n,DATA)),2) + " & _
    "' / ' + convert(varchar(10), DIASDASEMANA.DIADASEMANA) + ' - Per�odo ' + convert(varchar(10), Periodo) + ' de ' + inicio + ' as ' + termino AS DESCTURNO, TURNO.PKID  " & _
      "FROM PERIODO INNER JOIN TURNO ON PERIODO.PKID = TURNO.PERIODOID INNER JOIN DIASDASEMANA ON TURNO.DIADASEMANA = DIASDASEMANA.CODIGO "
  If strSqlWhere <> "" Then
    strSql = strSql & " WHERE " & strSqlWhere
  End If
  strSql = strSql & " ORDER BY DATA DESC"
        
  Set ListarTurnoPorUnidade = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTurno.ListarTurnoPorUnidade]", _
            Err.Description
End Function




