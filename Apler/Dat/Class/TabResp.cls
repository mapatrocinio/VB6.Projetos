VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTabResp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'intTipo Assume
' 1 - Reserva
' 2 - Entrada
Public Sub ExcluirTabResp(intTipo As Integer, _
                          lngTabRespId As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  If intTipo = 1 Then
    'Reserva
    strSql = "DELETE FROM TAB_RESPRESERVA WHERE PKID = " & _
      Formata_Dados(lngTabRespId, tpDados_Longo, tpNulo_NaoAceita)
  Else
    'Locacao
    strSql = "DELETE FROM TAB_RESPLOCACAO WHERE PKID = " & _
      Formata_Dados(lngTabRespId, tpDados_Longo, tpNulo_NaoAceita)
  
  End If
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTabResp.ExcluirTabResp]", _
            Err.Description
End Sub


Public Sub InserirTabResp(intTipo As Integer, _
                          lngResLocId As Long, _
                          lngResponsabilidadeId As Long, _
                          strTipo As String, _
                          lngGrupoCardapioId As String, _
                          lngQtdResponsabilidade As String, _
                          strCardapioId As String, _
                          strTurno As String, _
                          strTipoDiretoriaId As String)

  On Error GoTo trata
  '
  Dim strSql As String
  If intTipo = 1 Then
    'Rserva
    strSql = "INSERT INTO TAB_RESPRESERVA (RESERVAID, RESPONSABILIDADEID, TIPO, GRUPOCARDAPIOID, QTDRESPONSABILIDADE, CARDAPIOID, TURNO, TIPO_DIRETORIA) VALUES (" & _
      Formata_Dados(lngResLocId, tpDados_Longo, tpNulo_Aceita) & ", " & _
      Formata_Dados(lngResponsabilidadeId, tpDados_Longo, tpNulo_Aceita) & ", " & _
      Formata_Dados(strTipo, tpDados_Texto, tpNulo_Aceita) & ", " & _
      Formata_Dados(lngGrupoCardapioId, tpDados_Longo, tpNulo_Aceita) & ", " & _
      Formata_Dados(lngQtdResponsabilidade, tpDados_Longo, tpNulo_Aceita) & ", " & _
      Formata_Dados(strCardapioId, tpDados_Longo, tpNulo_Aceita) & ", " & _
      Formata_Dados(strTurno, tpDados_Texto, tpNulo_Aceita) & ", " & _
      Formata_Dados(strTipoDiretoriaId, tpDados_Texto, tpNulo_Aceita) & " " & _
      ")"
  Else
    'Locação
    strSql = "INSERT INTO TAB_RESPLOCACAO (LOCACAOID, RESPONSABILIDADEID, TIPO, GRUPOCARDAPIOID, QTDRESPONSABILIDADE, CARDAPIOID, TURNO, TIPO_DIRETORIA) VALUES (" & _
      Formata_Dados(lngResLocId, tpDados_Longo, tpNulo_Aceita) & ", " & _
      Formata_Dados(lngResponsabilidadeId, tpDados_Longo, tpNulo_Aceita) & ", " & _
      Formata_Dados(strTipo, tpDados_Texto, tpNulo_Aceita) & ", " & _
      Formata_Dados(lngGrupoCardapioId, tpDados_Longo, tpNulo_Aceita) & ", " & _
      Formata_Dados(lngQtdResponsabilidade, tpDados_Longo, tpNulo_Aceita) & ", " & _
      Formata_Dados(strCardapioId, tpDados_Longo, tpNulo_Aceita) & ", " & _
      Formata_Dados(strTurno, tpDados_Texto, tpNulo_Aceita) & ", " & _
      Formata_Dados(strTipoDiretoriaId, tpDados_Texto, tpNulo_Aceita) & " " & _
      ")"
  End If
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTabResp.InserirTabResp]", _
            Err.Description
End Sub

Public Sub AlterarTabResp(intTipo As Integer, _
                          lngTabRespId As Long, _
                          lngResponsabilidadeId As Long, _
                          strTipo As String, _
                          lngGrupoCardapioId As String, _
                          lngQtdResponsabilidade As String, _
                          strCardapioId As String, _
                          strTurno As String, _
                          strTipoDiretoriaId As String)
  On Error GoTo trata
  '
  Dim strSql As String
  If intTipo = 1 Then
    'Reserva
    strSql = "UPDATE TAB_RESPRESERVA SET " & _
      "TIPO = " & Formata_Dados(strTipo, tpDados_Texto, tpNulo_Aceita) & ", " & _
      "GRUPOCARDAPIOID = " & Formata_Dados(lngGrupoCardapioId, tpDados_Longo, tpNulo_Aceita) & ", " & _
      "RESPONSABILIDADEID = " & Formata_Dados(lngResponsabilidadeId, tpDados_Longo, tpNulo_Aceita) & ", " & _
      "QTDRESPONSABILIDADE = " & Formata_Dados(lngQtdResponsabilidade, tpDados_Longo, tpNulo_Aceita) & ", " & _
      "CARDAPIOID = " & Formata_Dados(strCardapioId, tpDados_Longo, tpNulo_Aceita) & ", " & _
      "TURNO = " & Formata_Dados(strTurno, tpDados_Texto, tpNulo_Aceita) & ", " & _
      "TIPO_DIRETORIA = " & Formata_Dados(strTipoDiretoriaId, tpDados_Texto, tpNulo_Aceita) & " " & _
      "WHERE PKID = " & lngTabRespId
  Else
    'Locacao
    strSql = "UPDATE TAB_RESPLOCACAO SET " & _
      "TIPO = " & Formata_Dados(strTipo, tpDados_Texto, tpNulo_Aceita) & ", " & _
      "GRUPOCARDAPIOID = " & Formata_Dados(lngGrupoCardapioId, tpDados_Longo, tpNulo_Aceita) & ", " & _
      "RESPONSABILIDADEID = " & Formata_Dados(lngResponsabilidadeId, tpDados_Longo, tpNulo_Aceita) & ", " & _
      "QTDRESPONSABILIDADE = " & Formata_Dados(lngQtdResponsabilidade, tpDados_Longo, tpNulo_Aceita) & ", " & _
      "CARDAPIOID = " & Formata_Dados(strCardapioId, tpDados_Longo, tpNulo_Aceita) & ", " & _
      "TURNO = " & Formata_Dados(strTurno, tpDados_Texto, tpNulo_Aceita) & ", " & _
      "TIPO_DIRETORIA = " & Formata_Dados(strTipoDiretoriaId, tpDados_Texto, tpNulo_Aceita) & " " & _
      "WHERE PKID = " & lngTabRespId
  End If
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTabResp.AlterarTabResp]", _
            Err.Description
End Sub

Public Function SelecionarTabResp(intTipo As Integer, _
                                  lngTabRespId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  If intTipo = 1 Then
    'Reserva
    strSql = "SELECT TAB_RESPRESERVA.*, RESPONSABILIDADE.DESCRICAO AS DESCRESPONSABILIDADE, TIPO.DESCRICAO AS DESCTIPO, TIPO_DIRETORIA.DESCRICAO AS DESCTIPO_DIRETORIA, GRUPOCARDAPIO.NOME AS DESCGRUPOCARDAPIO, CARDAPIO.DESCRICAO AS DESCCARDAPIO, case TAB_RESPRESERVA.QTDRESPONSABILIDADE when 0 then 'Responsável por tudo' when -1 then 'Responsável por nada' else convert(varchar(50), TAB_RESPRESERVA.QTDRESPONSABILIDADE) end AS QTDRESPONSABILIDADE, case TAB_RESPRESERVA.TURNO when 'P' then 'Por Turno' when 'T' then 'Total' else convert(varchar(50), TAB_RESPRESERVA.TURNO) end AS TURNO " & _
      "FROM ((((TAB_RESPRESERVA LEFT JOIN RESPONSABILIDADE ON RESPONSABILIDADE.PKID = TAB_RESPRESERVA.RESPONSABILIDADEID) " & _
      "LEFT JOIN TIPO ON TIPO.TIPO = TAB_RESPRESERVA.TIPO) " & _
      "LEFT JOIN TIPO_DIRETORIA ON TIPO_DIRETORIA.TIPO_DIRETORIA = TAB_RESPRESERVA.TIPO_DIRETORIA) " & _
      "LEFT JOIN GRUPOCARDAPIO ON GRUPOCARDAPIO.PKID = TAB_RESPRESERVA.GRUPOCARDAPIOID) " & _
      "LEFT JOIN CARDAPIO ON CARDAPIO.PKID = TAB_RESPRESERVA.CARDAPIOID " & _
      "WHERE TAB_RESPRESERVA.PKID = " & _
        Formata_Dados(lngTabRespId, tpDados_Longo, tpNulo_NaoAceita)
  Else
    'Locacao
    strSql = "SELECT TAB_RESPLOCACAO.*, RESPONSABILIDADE.DESCRICAO AS DESCRESPONSABILIDADE, TIPO.DESCRICAO AS DESCTIPO, TIPO_DIRETORIA.DESCRICAO AS DESCTIPO_DIRETORIA, GRUPOCARDAPIO.NOME AS DESCGRUPOCARDAPIO, CARDAPIO.DESCRICAO AS DESCCARDAPIO, case TAB_RESPLOCACAO.QTDRESPONSABILIDADE when 0 then 'Responsável por tudo' when -1 then 'Responsável por nada' else convert(varchar(50), TAB_RESPLOCACAO.QTDRESPONSABILIDADE) end AS QTDRESPONSABILIDADE, case TAB_RESPLOCACAO.TURNO when 'P' then 'Por Turno' when 'T' then 'Total' else convert(varchar(50), TAB_RESPLOCACAO.TURNO) end AS TURNO " & _
      "FROM ((((TAB_RESPLOCACAO LEFT JOIN RESPONSABILIDADE ON RESPONSABILIDADE.PKID = TAB_RESPLOCACAO.RESPONSABILIDADEID) " & _
      "LEFT JOIN TIPO ON TIPO.TIPO = TAB_RESPLOCACAO.TIPO) " & _
      "LEFT JOIN TIPO_DIRETORIA ON TIPO_DIRETORIA.TIPO_DIRETORIA = TAB_RESPLOCACAO.TIPO_DIRETORIA) " & _
      "LEFT JOIN GRUPOCARDAPIO ON GRUPOCARDAPIO.PKID = TAB_RESPLOCACAO.GRUPOCARDAPIOID) " & _
      "LEFT JOIN CARDAPIO ON CARDAPIO.PKID = TAB_RESPLOCACAO.CARDAPIOID " & _
      "WHERE TAB_RESPLOCACAO.PKID = " & _
        Formata_Dados(lngTabRespId, tpDados_Longo, tpNulo_NaoAceita)
  End If
  Set SelecionarTabResp = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTabResp.SelecionarTabResp]", _
            Err.Description
End Function



