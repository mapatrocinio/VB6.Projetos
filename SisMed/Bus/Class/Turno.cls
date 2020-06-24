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

Public Function ListarTurnoPorUnidade(strSqlWhere As String) As ADODB.Recordset
  On Error GoTo trata
  Dim clsTurno As datSisMed.clsTurno
  Set clsTurno = New datSisMed.clsTurno
  
  'Listar Turnos por unidades
  Set ListarTurnoPorUnidade = clsTurno.ListarTurnoPorUnidade(strSqlWhere)
  '
  Set clsTurno = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ListarTurno() As ADODB.Recordset
  On Error GoTo trata
  Dim clsTurno As datSisMed.clsTurno
  Set clsTurno = New datSisMed.clsTurno
  
  'Listar Turnos
  Set ListarTurno = clsTurno.ListarTurno
  '
  Set clsTurno = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Sub InserirTurno(lngFuncionarioId As Long, _
                        strPeriodoId As String, _
                        strData As String, _
                        blnStatus As Boolean, _
                        strSequencial As String, _
                        strDiaDaSemana As String, _
                        strVrCaixaInicial As String, _
                        strSequencialFecha As String)

  On Error GoTo trata
  Dim objTurno As datSisMed.clsTurno
  '
  Set objTurno = New datSisMed.clsTurno
  'Inserir Turno - Abertura

  objTurno.InserirTurno lngFuncionarioId, _
                        strPeriodoId, _
                        strData, _
                        blnStatus, _
                        strSequencial, _
                        strDiaDaSemana, _
                        strVrCaixaInicial, _
                        strSequencialFecha
  '
  Set objTurno = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub FecharTurno(lngTurnoId As Long, _
                       blnStatus As Boolean, _
                       strDataFechamento As String)

  On Error GoTo trata
  Dim objTurno As datSisMed.clsTurno
  '
  Set objTurno = New datSisMed.clsTurno
  
  'Alterar Turno - Fechamento

  objTurno.FecharTurno lngTurnoId, _
                       blnStatus, _
                       strDataFechamento
  Set objTurno = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


