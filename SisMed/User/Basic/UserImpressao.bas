Attribute VB_Name = "mdlUserImpressao"
Option Explicit


'local variable(s) to hold property value(s)
Private mvarScaleMode As Integer 'local copy
Private mvarCurrentX As Double 'local copy
Private mvarCurrentY As Double 'local copy
Private mvarFontName As String 'local copy
Private mvarDrawStyle As Integer 'local copy
Private mvarCopies As Integer 'local copy
Private mvarFontSize As Double 'local copy
'Constantes para impressao
Const lLARGURACAIXA = 11
Const lLARGURACAIXAA5 = 30
Const dLSTART = 1.3
'Espacamento
Const lDELTAX = 0.5
Const lDELTAY = 0.5
'Constantes das caixas
'Posicao Quadrantes 1 e 3
Const lPOSX = 0
Const lPOSY = 0 + lDELTAY

Public Sub COMPROV_SANGRIA(ByVal strEmpresa As String, _
                           ByVal intQtdVias As Integer, _
                           strUsuario As String)
  'Imprimir corpo
  Dim sSql As String
  Dim rs As ADODB.Recordset
  Dim objGeral As busSisMed.clsGeral
  '
  Dim vrRetDin    As Currency
  Dim vrRetCar    As Currency
  Dim vrRetCarDeb As Currency
  Dim vrRetChq    As Currency
  Dim vrRetPen    As Currency
  Dim vrRetFat    As Currency
  '
  Dim vrDepDin      As Currency
  Dim vrDepCar      As Currency
  Dim vrDepCarDeb   As Currency
  Dim vrDepChq      As Currency
  Dim vrDepPen      As Currency
  Dim vrDepFat      As Currency
  Dim VRTotal       As Currency
  
  Dim sMsg As String
  Dim dDataAtual  As Date
  Dim I As Integer
  '
  Set objGeral = New busSisMed.clsGeral
  '
  sSql = "SELECT TOP 1 SANGRIA.*, TURNO.DATA AS DATATURNO From " & _
    "SANGRIA INNER JOIN TURNO ON TURNO.PKID = SANGRIA.TURNOID" & _
    " ORDER BY SANGRIA.PKID DESC"
    
  Set rs = objGeral.ExecutarSQL(sSql)
  dDataAtual = Now
  Do While Not rs.EOF
    'Calculo dos campos
    vrRetDin = IIf(Not IsNumeric(rs!vrRetDin), 0, rs!vrRetDin)
    vrRetCar = IIf(Not IsNumeric(rs!vrRetCar), 0, rs!vrRetCar)
    vrRetCarDeb = IIf(Not IsNumeric(rs!vrRetCarDeb), 0, rs!vrRetCarDeb)
    vrRetChq = IIf(Not IsNumeric(rs!vrRetChq), 0, rs!vrRetChq)
    vrRetPen = IIf(Not IsNumeric(rs!vrRetPen), 0, rs!vrRetPen)
    vrRetFat = IIf(Not IsNumeric(rs!vrRetFat), 0, rs!vrRetFat)
    '
    vrDepDin = IIf(Not IsNumeric(rs!vrDepDin), 0, rs!vrDepDin)
    vrDepCar = IIf(Not IsNumeric(rs!vrDepCar), 0, rs!vrDepCar)
    vrDepCarDeb = IIf(Not IsNumeric(rs!vrDepCarDeb), 0, rs!vrDepCarDeb)
    vrDepChq = IIf(Not IsNumeric(rs!vrDepChq), 0, rs!vrDepChq)
    vrDepPen = IIf(Not IsNumeric(rs!vrDepPen), 0, rs!vrDepPen)
    vrDepFat = IIf(Not IsNumeric(rs!vrDepFat), 0, rs!vrDepFat)
    '
    VRTotal = vrRetDin + vrRetCar + vrRetCarDeb + vrRetChq + vrRetPen + vrRetFat - vrDepDin - vrDepCar - vrDepCarDeb - vrDepChq - vrDepPen - vrDepFat
    '
    For I = 1 To intQtdVias
      '--------
      QUEBRA_LINHA "MOVIMENTACAO CAIXA", lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA strEmpresa, lLARGURACAIXA - (dLSTART / 2)
      IMPRIMIR_LINHA 1
      PULAR_LINHA 1
      QUEBRA_LINHA Format(dDataAtual, "DD/MM/YYYY hh:mm"), lLARGURACAIXA - (dLSTART / 2)
      PULAR_LINHA 1
      QUEBRA_LINHA "Data Hora Turno " & Format(rs!DATATURNO, "DD/MM/YYYY hh:mm"), lLARGURACAIXA - (dLSTART / 2)
      PULAR_LINHA 1
      QUEBRA_LINHA "Usuario " & strUsuario, lLARGURACAIXA - (dLSTART / 2)
      PULAR_LINHA 1
      QUEBRA_LINHA "RETIRADO", lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA "Dinheiro R$ " & Format(vrRetDin, "###,##0.00"), lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA "Cartao Cred R$ " & Format(vrRetCar, "###,##0.00"), lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA "Cartao Deb R$ " & Format(vrRetCarDeb, "###,##0.00"), lLARGURACAIXA - (dLSTART / 2)
      'QUEBRA_LINHA "Cheque R$ " & Format(vrRetChq, "###,##0.00"), lLARGURACAIXA - (dLSTART / 2)
      'QUEBRA_LINHA "Penhor R$ " & Format(vrRetPen, "###,##0.00"), lLARGURACAIXA - (dLSTART / 2)
      'QUEBRA_LINHA "Fatura R$ " & Format(vrRetFat, "###,##0.00"), lLARGURACAIXA - (dLSTART / 2)
      If gbTrabComDepSangria Then
        PULAR_LINHA 1
        QUEBRA_LINHA "DEPOSITADO", lLARGURACAIXA - (dLSTART / 2)
        QUEBRA_LINHA "Dinheiro R$ " & Format(vrDepDin, "###,##0.00"), lLARGURACAIXA - (dLSTART / 2)
        QUEBRA_LINHA "Cartao Cred R$ " & Format(vrDepCar, "###,##0.00"), lLARGURACAIXA - (dLSTART / 2)
        QUEBRA_LINHA "Cartao Deb R$ " & Format(vrDepCarDeb, "###,##0.00"), lLARGURACAIXA - (dLSTART / 2)
        QUEBRA_LINHA "Cheque R$ " & Format(vrDepChq, "###,##0.00"), lLARGURACAIXA - (dLSTART / 2)
        QUEBRA_LINHA "Penhor R$ " & Format(vrDepPen, "###,##0.00"), lLARGURACAIXA - (dLSTART / 2)
        QUEBRA_LINHA "Fatura R$ " & Format(vrDepFat, "###,##0.00"), lLARGURACAIXA - (dLSTART / 2)
      End If
      PULAR_LINHA 1
      QUEBRA_LINHA "Total R$ " & Format(VRTotal, "###,##0.00"), lLARGURACAIXA - (dLSTART / 2)
      PULAR_LINHA 1
      QUEBRA_LINHA "Gerente " & rs!RESPONSAVEL, lLARGURACAIXA - (dLSTART / 2)
      
      PULAR_LINHA 1
      IMPRIMIR_LINHA 1
      '
      'PULAR_LINHA 30
    Next
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  Set objGeral = Nothing
  For I = 1 To intQtdPontos
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, ".", "Left", True
  Next
End Sub

Public Sub IMP_COMPROV_SANGRIA(strEmpresa As String, _
                              intQtdVias As Integer, _
                              strUsuario As String)
  'Impressao do comprovante de Penhor
On Error GoTo ErrHandler
  '
  INICIA_IMPRESSAO
  '
  COMPROV_SANGRIA strEmpresa, intQtdVias, _
                              strUsuario
  '
  TERMINA_IMPRESSAO
  '
  Exit Sub
ErrHandler:
  MsgBox "O seguinte erro ocorreu: " & Err.Description & ". Erro na impressao do comprovante de Entrada na portaria.", vbExclamation, TITULOSISTEMA
  TERMINA_IMPRESSAO
End Sub

Public Sub IMP_COMPROV_REC(lngGRID As Long, _
                           strEmpresa As String, _
                           intQtdVias As Integer)
  'Impressao do comprovante de Fatura
On Error GoTo ErrHandler
  '
  INICIA_IMPRESSAO
  '
  COMPROV_REC lngGRID, strEmpresa, intQtdVias
  '
  TERMINA_IMPRESSAO
  '
  Exit Sub
ErrHandler:
  TERMINA_IMPRESSAO
  Err.Raise Err.Number, _
            "[mdlUserImpressao.IMP_COMPROV_REC]", _
            Err.Description
End Sub

Public Sub IMP_COMPROV_FATURA(lngCCID As Long, _
                              strEmpresa As String, _
                              intQtdVias As Integer)
  'Impressao do comprovante de Fatura
On Error GoTo ErrHandler
  '
  'INICIA_IMPRESSAO
  '
  'COMPROV_FATURA lngCCID, strEmpresa, intQtdVias
  '
  'TERMINA_IMPRESSAO
  '
  Exit Sub
ErrHandler:
  TERMINA_IMPRESSAO
  Err.Raise Err.Number, _
            "[mdlUserImpressao.IMP_COMP_GR]", _
            Err.Description
End Sub

Public Sub COMPROV_FATURA(ByVal lngCCID As Long, _
                          ByVal strEmpresa As String, _
                          ByVal intQtdVias As Integer)
  'Imprimir corpo
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  Dim objGeral As busSisMed.clsGeral
  '
  Dim VRTotalFatura As Currency
  Dim datDataAtual  As Date
  Dim I As Integer
  '
  Set objGeral = New busSisMed.clsGeral
  strSql = "SELECT EMPRESA.NOME, APARTAMENTO.NUMERO, CONTACORRENTE.VALOR AS PGTOFATURA, LOCACAO.SEQUENCIAL, CONTACORRENTE.DTHORACC, PARCELA.VRPARCELA, PARCELA.PARCELA, PARCELA.DTVENCIMENTO" & _
    " FROM GR " & _
    "INNER JOIN CONTACORRENTE ON GR.PKID = CONTACORRENTE.GRID " & _
    "INNER JOIN PARCELA ON CONTACORRENTE.PKID = PARCELA.CONTACORRENTEID " & _
    "INNER JOIN VIAGEM ON LOCACAO.PKID = VIAGEM.LOCACAOID " & _
    "LEFT JOIN EMPRESA ON EMPRESA.PKID = VIAGEM.EMPRESAID " & _
    "WHERE CONTACORRENTE.PKID = " & Formata_Dados(lngCCID, tpDados_Longo) & _
    " ORDER BY PARCELA "
    
  Set objRs = objGeral.ExecutarSQL(strSql)
  datDataAtual = Now
  If Not objRs.EOF Then
    For I = 1 To intQtdVias
      VRTotalFatura = IIf(Not IsNumeric(objRs!PGTOFATURA), 0, objRs!PGTOFATURA)
      'Impressão do cabeçalho
      QUEBRA_LINHA "COMPROVANTE DE FATURA", lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA strEmpresa, lLARGURACAIXA - (dLSTART / 2)
      If objRs!NOME & "" <> "<PARTICULAR>" Then
        PULAR_LINHA 1
        QUEBRA_LINHA objRs!NOME & "", lLARGURACAIXA - (dLSTART / 2)
      End If
      PULAR_LINHA 1
      QUEBRA_LINHA objRs!NUMERO, lLARGURACAIXA - (dLSTART / 2)
      IMPRIMIR_LINHA 1
      QUEBRA_LINHA Format(objRs!DTHORACC, "DD/MM/YYYY hh:mm"), lLARGURACAIXA - (dLSTART / 2)
      IMPRIMIR_LINHA 0
      '------------------------------------
      'Rotina para imprimir na ordem correta
      IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, "Valor", "Right", False
      IMPRIMIR_POSICAO_CORRETA 0.5, lLARGURACAIXA - 2.2, Printer.CurrentY, "Parc", "Right", False
      'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 3.9, Printer.CurrentY, "Uni", "Right", False
      IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Data", "Left", True
      '
      'QUEBRA_LINHA "Descricao     Uni Qd   Valor", lLARGURACAIXA - (dLSTART / 2)
      IMPRIMIR_LINHA 0
      
      Do While Not objRs.EOF
        'QUEBRA_LINHA  , lLARGURACAIXA - (dLSTART / 2)
        IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(objRs!VRPARCELA, "#,##0.00"), "Right", False
        IMPRIMIR_POSICAO_CORRETA 0.5, lLARGURACAIXA - 2.2, Printer.CurrentY, Format(objRs!PARCELA, "#,##0"), "Right", False
        'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 3.9, Printer.CurrentY, Format(objRsInt!Valor, "#,##0.00"), "Right", False
        IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, Format(objRs!DTVENCIMENTO, "DD/MM/YYYY"), "Left", True
        '
        objRs.MoveNext
        If objRs.EOF Then
          IMPRIMIR_LINHA 0
          PULAR_LINHA 1
        End If
      Loop
      'TOTAL
      IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(VRTotalFatura, "#,##0.00"), "Right", False
      'IMPRIMIR_POSICAO_CORRETA 0.5, lLARGURACAIXA - 2.2, Printer.CurrentY, Format(DobjRs!PARCELA, "#,##0"), "Right", False
      'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 3.9, Printer.CurrentY, Format(objRsInt!Valor, "#,##0.00"), "Right", False
      IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "TOTAL", "Left", True
      objRs.MoveFirst
    Next
    For I = 1 To intQtdPontos
      IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, ".", "Left", True
    Next
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
End Sub


Public Sub IMP_COMPROV_FECHA_TURNO(ByVal lngTURNOID As Long, _
                                   strEmpresa As String, _
                                   intQtdVias As Integer)
  'Impressao de GR
On Error GoTo trata
  '
  INICIA_IMPRESSAO
  '
  COMPROV_FECHA_TURNO lngTURNOID, _
                      strEmpresa, _
                      intQtdVias
  '
  TERMINA_IMPRESSAO
  '
  Exit Sub
trata:
  TERMINA_IMPRESSAO
  Err.Raise Err.Number, _
            "[mdlUserImpressao.IMP_COMPROV_FECHA_TURNO]", _
            Err.Description
  
End Sub

Public Sub IMP_COMP_GR(ByVal lngGRID As Long, _
                      strEmpresa As String, _
                      intQtdVias As Integer, _
                      bln2Via As Boolean)
  'Impressao de GR
On Error GoTo trata
  '
  INICIA_IMPRESSAO
  '
  If gbTrabImpA5 = True Then
    COMPROV_GR_A5 lngGRID, _
                  strEmpresa, _
                  intQtdVias, _
                  bln2Via
  Else
    COMPROV_GR lngGRID, _
              strEmpresa, _
              intQtdVias, _
              bln2Via
  End If
  '
  TERMINA_IMPRESSAO
  '
  Exit Sub
trata:
  TERMINA_IMPRESSAO
  Err.Raise Err.Number, _
            "[mdlUserImpressao.IMP_COMP_GR]", _
            Err.Description
  
End Sub

Public Sub IMP_COMP_CANC_GR(ByVal lngGRID As Long, _
                            strEmpresa As String, _
                            intQtdVias As Integer)
  'Impressao de CANC DE GR
On Error GoTo trata
  '
  INICIA_IMPRESSAO
  '
  COMPROV_CANC_GR lngGRID, _
                  strEmpresa, _
                  intQtdVias
  '
  TERMINA_IMPRESSAO
  '
  Exit Sub
trata:
  TERMINA_IMPRESSAO
  Err.Raise Err.Number, _
            "[mdlUserImpressao.COMPROV_CANC_GR]", _
            Err.Description
  
End Sub


Public Sub PULAR_LINHA(ByVal pQtd As Integer)
  Printer.CurrentY = Printer.CurrentY + (Printer.TextHeight("X") * pQtd)
End Sub

Public Sub INICIA_IMPRESSAO()
  'Inicializa Printer
  ScaleMode = vbCentimeters
  CurrentX = lPOSX
  CurrentY = lPOSY
  FontName = "Arial"
  DrawStyle = vbSolid
  Copies = 1
  FontSize = 10
  '
  Printer.ScaleMode = ScaleMode
  Printer.CurrentX = CurrentX
  Printer.CurrentY = CurrentY
  Printer.FontName = FontName
  Printer.DrawStyle = DrawStyle
  On Error Resume Next
  Printer.Copies = Copies
  '
  Printer.FontSize = FontSize

End Sub

Public Sub TERMINA_IMPRESSAO()
  'Enviando para a impressora
  Printer.EndDoc
End Sub




Public Sub COMPROV_REC(ByVal lngGRID As Long, _
                       strEmpresa As String, _
                       intQtdVias As Integer)
  'Imprimir corpo
  Dim strSql          As String
  Dim objRs           As ADODB.Recordset
  Dim objGeral        As busSisMed.clsGeral
  Dim objCC           As busSisMed.clsContaCorrente
  Dim lngQtdTotal     As Long
  Dim curVrTotal      As Currency
  Dim curVrTroco      As Currency
  Dim intI As Integer
  '
  'Dim sVenda As String
  'Dim sCobranca As String
  'Dim sData As String
  'Dim sTipo As String
  '
  For intI = 1 To intQtdVias
    Set objGeral = New busSisMed.clsGeral
    '
    strSql = "SELECT GR.PKID, FUNCIONARIO.NOME AS FUNCIONARIO, GR.SEQUENCIAL, SALA.NUMERO, GR.DATA, PRONTUARIO.NOME AS PACIENTE, PRESTADOR.NOME AS PRESTADOR, GRPROCEDIMENTO.QTD, GRPROCEDIMENTO.VALOR, PROCEDIMENTO.PROCEDIMENTO " & _
      " FROM GR " & _
      " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
      "   INNER JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
      "   INNER JOIN PRONTUARIO AS PRESTADOR ON PRESTADOR.PKID = ATENDE.PRONTUARIOID " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = GR.PRONTUARIOID " & _
      " INNER JOIN PRONTUARIO AS FUNCIONARIO ON FUNCIONARIO.PKID = GR.FUNCIONARIOFECHAID " & _
      " LEFT JOIN GRPROCEDIMENTO ON GR.PKID = GRPROCEDIMENTO.GRID " & _
      " LEFT JOIN PROCEDIMENTO ON PROCEDIMENTO.PKID = GRPROCEDIMENTO.PROCEDIMENTOID " & _
      "WHERE GR.PKID = " & lngGRID
  
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      
      '----- IMPRIMIR HEADER
  
      IMPRIMIR_LINHA 1
      QUEBRA_LINHA "------ GUIA DE RECOLHIMENTO ------", lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA strEmpresa, lLARGURACAIXA - (dLSTART / 2)
      IMPRIMIR_LINHA 1
      PULAR_LINHA 2
      '
      QUEBRA_LINHA "CAIXA : " & objRs.Fields("FUNCIONARIO").Value, lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA "DATA : " & Format(Now, "DD/MM/YYYY - hh:mm"), lLARGURACAIXA - (dLSTART / 2)
      IMPRIMIR_LINHA 1
      PULAR_LINHA 2
      '
      QUEBRA_LINHA "GR : " & objRs.Fields("SEQUENCIAL").Value & " SALA : " & objRs.Fields("NUMERO").Value, lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA "DATA : " & Format(objRs.Fields("DATA").Value, "DD/MM/YYYY - hh:mm"), lLARGURACAIXA - (dLSTART / 2)
      IMPRIMIR_LINHA 1
      PULAR_LINHA 1
      
      '----- IMPRIMIR DETALHES
      
      lngQtdTotal = 0
      curVrTotal = 0
      curVrTroco = 0
      Do While Not objRs.EOF
        'Soma Totais
        lngQtdTotal = lngQtdTotal + IIf(Not IsNumeric(objRs.Fields("QTD").Value), 0, objRs.Fields("QTD").Value)
        curVrTotal = curVrTotal + IIf(Not IsNumeric(objRs.Fields("VALOR").Value), 0, objRs.Fields("VALOR").Value)
  
        objRs.MoveNext
      Loop
      '---
      '
      PULAR_LINHA 1
      IMPRIMIR_LINHA 0
      objRs.Close
      Set objRs = Nothing
      '---
      'PAGAMENTO
      Set objCC = New busSisMed.clsContaCorrente
      Set objRs = objCC.SelecionarPagamentos("RC", _
                                             lngGRID)
      
      If objRs.EOF Then
        IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Nao houve pagamento", "Left", True
      End If
      Do While Not objRs.EOF
        IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(objRs.Fields("VALOR").Value, "###,##0.00"), "Right", False
        
        If objRs.Fields("INDCONVENIO").Value & "" = "S" Then
          'CONVÊNIO
          'IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Convenio", "Left", True
          IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, objRs.Fields("NOME_CARTAO_DEBITO").Value & "", "Left", True
        Else
          IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, objRs.Fields("DESC_STATUSCC").Value, "Left", True
        End If
        
        curVrTroco = curVrTroco + IIf(Not IsNumeric(objRs.Fields("VRTROCO").Value), 0, objRs.Fields("VRTROCO").Value)
        objRs.MoveNext
      Loop
      If curVrTroco > 0 Then
        IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curVrTroco, "###,##0.00"), "Right", False
        IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Troco", "Left", True
      End If
      objRs.Close
      Set objRs = Nothing
      Set objCC = Nothing
      '---
      IMPRIMIR_LINHA 0
      '
      IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curVrTotal, "#,##0.00"), "Right", False
      'IMPRIMIR_POSICAO_CORRETA 0.5, lLARGURACAIXA - 2.2, Printer.CurrentY, Format(lngQtdTotal, "#,##0"), "Right", False
      'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 3.9, Printer.CurrentY, Format(objRs![CARDAPIO.Valor], "#,##0.00"), "Right", False
      IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "VALOR TOTAL", "Left", True
      '
      PULAR_LINHA 2
    End If
  Next
  Set objGeral = Nothing
  For intI = 1 To intQtdPontos
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, ".", "Left", True
  Next
End Sub


Public Sub COMPROV_FECHA_TURNO(ByVal lngTURNOID As Long, _
                               strEmpresa As String, _
                               intQtdVias As Integer)
  'Imprimir corpo
  Dim strSql          As String
  Dim objRs           As ADODB.Recordset
  Dim objGeral        As busSisMed.clsGeral
  Dim objCC           As busSisMed.clsContaCorrente
  Dim lngQtdTotal     As Long
  Dim curVrTotal      As Currency
  Dim curVrTroco      As Currency
  Dim intI As Integer
  '
  '-----------------------------------
  'VARIÁVEIS TOTALIZADORAS DO TURNO
  '-----------------------------------
  'Geral
  Dim strNome             As String
  Dim strInicio           As String
  Dim strTermino          As String
  'EM CAIXA
  Dim curCaixaInicial     As Currency
  'GR PAGAS
  Dim curGRDin            As Currency
  Dim curGRCC             As Currency
  Dim curGRCD             As Currency
  'GR DEVOLVIDAS
  Dim curGRDevDin         As Currency
  Dim curGRDevCC          As Currency
  Dim curGRDevCD          As Currency
  'GR DEVOLVIDAS (PRÓPRIO CAIXA)
  Dim curGRDevPCDin       As Currency
  Dim curGRDevPCCC        As Currency
  Dim curGRDevPCCD        As Currency
  'GR LAB PAGAS
  Dim curGRLabDin         As Currency
  Dim curGRLabCC          As Currency
  Dim curGRLabCD          As Currency
  'GR LAB DEVOLVIDAS
  Dim curGRLabDevDin      As Currency
  Dim curGRLabDevCC       As Currency
  Dim curGRLabDevCD       As Currency
  'SANGRIA
  Dim curSangriaDin       As Currency
  Dim curSangriaCC        As Currency
  Dim curSangriaCD        As Currency
  'SALDO
  Dim curSaldoDin         As Currency
  Dim curSaldoCC          As Currency
  Dim curSaldoCD          As Currency
  'TOTAL
  Dim curTotalDin         As Currency
  Dim curTotalCaixa       As Currency
  '
  For intI = 1 To intQtdVias
    strNome = ""
    strInicio = ""
    strTermino = ""
    'EM CAIXA
    curCaixaInicial = 0
    'GR PAGAS
    curGRDin = 0
    curGRCC = 0
    curGRCD = 0
    'GR DEVOLVIDAS
    curGRDevDin = 0
    curGRDevCC = 0
    curGRDevCD = 0
    'GR DEVOLVIDAS (PRÓPRIO CAIXA)
    curGRDevPCDin = 0
    curGRDevPCCC = 0
    curGRDevPCCD = 0
    
    'GR LAB PAGAS
    curGRLabDin = 0
    curGRLabCC = 0
    curGRLabCD = 0
    'GR LAB DEVOLVIDAS
    curGRLabDevDin = 0
    curGRLabDevCC = 0
    curGRLabDevCD = 0
    'SANGRIA
    curSangriaDin = 0
    curSangriaCC = 0
    curSangriaCD = 0
    'SALDO
    curSaldoDin = 0
    curSaldoCC = 0
    curSaldoCD = 0
    'TOTAL
    curTotalDin = 0
    curTotalCaixa = 0
    '
    'Captura totais
    
    Set objGeral = New busSisMed.clsGeral
    '
    strSql = "SELECT  TURNO.PKID, sum(vw_cons_t_sangria.VRRETDIN) as VRRETDIN, SUM(vw_cons_t_sangria.VRRETCAR) AS VRRETCAR, SUM(vw_cons_t_sangria.VRRETCARDEB) AS VRRETCARDEB " & _
      " FROM TURNO " & _
      " LEFT JOIN vw_cons_t_sangria ON TURNO.PKID = vw_cons_t_sangria.TURNOID " & _
      "WHERE TURNO.PKID = " & Formata_Dados(lngTURNOID, tpDados_Longo) & _
      " GROUP BY TURNO.PKID "
  
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      'Obtém dados gerais
      curSangriaDin = objRs.Fields("VRRETDIN").Value
      curSangriaCC = objRs.Fields("VRRETCAR").Value
      curSangriaCD = objRs.Fields("VRRETCARDEB").Value
    End If
    objRs.Close
    Set objRs = Nothing
    '
    strSql = "SELECT  TURNO.PKID, vw_cons_t_cred_turno.TURNOID, vw_cons_t_cred_turno.TURNOCANCID, vw_cons_t_cred_turno.TURNOCANCABREID, vw_cons_t_cred_turno.TURNOLABID, vw_cons_t_cred_turno.STATUS, MAX(TURNO.VRCAIXAINICIAL) AS VRCAIXAINICIAL, MAX(PRONTUARIO.NOME) AS NOME, MAX(TURNO.DATA) AS DATA, MAX(TURNO.DTFECHAMENTO) AS DTFECHAMENTO, " & _
      " sum(PgtoCartaoDeb) as PgtoCartaoDeb, sum(PgtoEspecie) as PgtoEspecie, sum(PgtoCartao) as PgtoCartao, sum(PgtoTroco) as PgtoTroco " & _
      " FROM TURNO " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = TURNO.PRONTUARIOID " & _
      " LEFT JOIN vw_cons_t_cred_turno ON TURNO.PKID = vw_cons_t_cred_turno.TURNOID " & _
      "WHERE TURNO.PKID = " & Formata_Dados(lngTURNOID, tpDados_Longo) & _
      " GROUP BY TURNO.PKID, vw_cons_t_cred_turno.TURNOID, vw_cons_t_cred_turno.TURNOCANCID, vw_cons_t_cred_turno.TURNOCANCABREID, vw_cons_t_cred_turno.TURNOLABID, vw_cons_t_cred_turno.STATUS "
  
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      'Obtém dados gerais
      If IsNull(objRs.Fields("VRCAIXAINICIAL").Value) = False Then
        curCaixaInicial = objRs.Fields("VRCAIXAINICIAL").Value
      End If
      strNome = objRs.Fields("NOME").Value & ""
      strInicio = Format(objRs.Fields("DATA").Value, "DD/MM/YYYY hh:mm")
      strTermino = Format(objRs.Fields("DTFECHAMENTO").Value, "DD/MM/YYYY hh:mm")
    End If
    Do While Not objRs.EOF
      If IsNull(objRs.Fields("TURNOLABID").Value) = True Then
        'CAIXA
        If objRs.Fields("STATUS").Value & "" = "C" Then
          'CANCELADA
          If objRs.Fields("TURNOCANCABREID").Value = lngTURNOID Then
            'GR DEVOLVIDAS (PRÓPRIO CAIXA)
            curGRDin = curGRDin + objRs.Fields("PgtoEspecie").Value - objRs.Fields("PgtoTroco").Value
            curGRCC = curGRCC + objRs.Fields("PgtoCartao").Value
            curGRCD = curGRCD + objRs.Fields("PgtoCartaoDeb").Value
          End If
          If objRs.Fields("TURNOCANCID").Value = lngTURNOID Then
            'GR DEVOLVIDAS (PRÓPRIO CAIXA)
            'GR DEVOLVIDAS
            curGRDevDin = curGRDevDin + objRs.Fields("PgtoEspecie").Value - objRs.Fields("PgtoTroco").Value
            curGRDevCC = curGRDevCC + objRs.Fields("PgtoCartao").Value
            curGRDevCD = curGRDevCD + objRs.Fields("PgtoCartaoDeb").Value
          End If
        ElseIf objRs.Fields("STATUS").Value & "" = "F" Then
          'FECHADA
          'GR PAGAS
          curGRDin = curGRDin + objRs.Fields("PgtoEspecie").Value - objRs.Fields("PgtoTroco").Value
          curGRCC = curGRCC + objRs.Fields("PgtoCartao").Value
          curGRCD = curGRCD + objRs.Fields("PgtoCartaoDeb").Value
        End If
      Else
        'LABORATÓRIO
        If objRs.Fields("STATUS").Value & "" = "C" Then
          'CANCELADA
          'GR LAB DEVOLVIDAS
          curGRLabDevDin = curGRLabDevDin + objRs.Fields("PgtoEspecie").Value - objRs.Fields("PgtoTroco").Value
          curGRLabDevCC = curGRLabDevCC + objRs.Fields("PgtoCartao").Value
          curGRLabDevCD = curGRLabDevCD + objRs.Fields("PgtoCartaoDeb").Value
        ElseIf objRs.Fields("STATUS").Value & "" = "F" Then
          'FECHADA
          'GR LAB PAGAS
          curGRLabDin = curGRLabDin + objRs.Fields("PgtoEspecie").Value - objRs.Fields("PgtoTroco").Value
          curGRLabCC = curGRLabCC + objRs.Fields("PgtoCartao").Value
          curGRLabCD = curGRLabCD + objRs.Fields("PgtoCartaoDeb").Value
        End If
      End If
      objRs.MoveNext
    Loop
    objRs.Close
    Set objRs = Nothing
    'Sumariza totais
    'Ajusta valores CAIXA
    'curGRDin = curGRDin + curGRDevPCDin
    'curGRCC = curGRCC + curGRDevPCCC
    'curGRCD = curGRCD + curGRDevPCCD
    'Ajusta valores LAB
    curGRLabDin = curGRLabDin + curGRLabDevDin
    curGRLabCC = curGRLabCC + curGRLabDevCC
    curGRLabCD = curGRLabCD + curGRLabDevCD
    'SALDO
    curSaldoDin = curGRDin + curGRLabDin - curGRDevDin - curGRLabDevDin - curSangriaDin
    curSaldoCC = curGRCC + curGRLabCC - curGRDevCC - curGRLabDevCC - curSangriaCC
    curSaldoCD = curGRCD + curGRLabCD - curGRDevCD - curGRLabDevCD - curSangriaCD
    'TOTAL
    curTotalDin = curSaldoDin + curCaixaInicial
    curTotalCaixa = curSaldoDin + curSaldoCC + curSaldoCD + curCaixaInicial
    '---------------------
    '----- IMPRIMIR HEADER
    '---------------------
    IMPRIMIR_LINHA 1
    QUEBRA_LINHA "---------- GUIA DE RECOLHIMENTO ----------", lLARGURACAIXA - (dLSTART / 2)
    QUEBRA_LINHA strEmpresa, lLARGURACAIXA - (dLSTART / 2)
    IMPRIMIR_LINHA 1
    PULAR_LINHA 2
    '
    IMPRIMIR_LINHA 1
    QUEBRA_LINHA "---------- FECHAMENTO DO CAIXA ----------", lLARGURACAIXA - (dLSTART / 2)
    IMPRIMIR_LINHA 1
    PULAR_LINHA 2
    '
    QUEBRA_LINHA "FUNCIONARIO : " & strNome, lLARGURACAIXA - (dLSTART / 2)
    QUEBRA_LINHA "ABERTURA : " & strInicio, lLARGURACAIXA - (dLSTART / 2)
    QUEBRA_LINHA "FECHAMENTO : " & strTermino, lLARGURACAIXA - (dLSTART / 2)
    IMPRIMIR_LINHA 1
    PULAR_LINHA 1
    '----- IMPRIMIR DETALHES
    'CAIXA INICIAL
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curCaixaInicial, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "CAIXA INICIAL", "Left", True
    IMPRIMIR_LINHA 0
    PULAR_LINHA 1
    'GR´S PAGAS
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "GR´S", "Center", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curGRDin, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "DINHEIRO", "Left", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curGRCC, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "CREDITO", "Left", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curGRCD, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "DEBITO", "Left", True
    '
    IMPRIMIR_LINHA 0
    PULAR_LINHA 1
    'DEVOLUÇÃO DE GR´S
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "DEVOLUCAO DE GR´S", "Center", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curGRDevDin, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "DINHEIRO", "Left", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curGRDevCC, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "CREDITO", "Left", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curGRDevCD, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "DEBITO", "Left", True
    '
    IMPRIMIR_LINHA 0
    PULAR_LINHA 1
    'LABORATÓRIO
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "LABORATORIO", "Center", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curGRLabDin, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "DINHEIRO", "Left", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curGRLabCC, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "CREDITO", "Left", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curGRLabCD, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "DEBITO", "Left", True
    '
    IMPRIMIR_LINHA 0
    PULAR_LINHA 1
    'DEVOLUÇÃOD E LAB
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "DEVOLUCAO DE LABORATORIO", "Center", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curGRLabDevDin, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "DINHEIRO", "Left", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curGRLabDevCC, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "CREDITO", "Left", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curGRLabDevCD, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "DEBITO", "Left", True
    '
    IMPRIMIR_LINHA 0
    PULAR_LINHA 1
    'SANGRIA
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "SANGRIA", "Center", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curSangriaDin, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "DINHEIRO", "Left", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curSangriaCC, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "CREDITO", "Left", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curSangriaCD, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "DEBITO", "Left", True
    '
    IMPRIMIR_LINHA 0
    PULAR_LINHA 1
    'SALDO
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "SALDO", "Center", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curSaldoDin, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "DINHEIRO", "Left", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curSaldoCC, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "CREDITO", "Left", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curSaldoCD, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "DEBITO", "Left", True
    '
    IMPRIMIR_LINHA 0
    PULAR_LINHA 1
    'TOTAIS
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curTotalDin, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "TOTAL DE DINHEIRO", "Left", True
    '
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curTotalCaixa, "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "TOTAL DO CAIXA", "Left", True
    '
    IMPRIMIR_LINHA 0
    PULAR_LINHA 1
    '---
    PULAR_LINHA 2
  Next
  Set objGeral = Nothing
  For intI = 1 To intQtdPontos
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, ".", "Left", True
  Next
End Sub


Public Sub COMPROV_GR(ByVal lngGRID As Long, _
                      strEmpresa As String, _
                      intQtdVias As Integer, _
                      bln2Via As Boolean)
  'Imprimir corpo
  Dim strSql          As String
  Dim objRs           As ADODB.Recordset
  Dim objGeral        As busSisMed.clsGeral
  Dim objCC           As busSisMed.clsContaCorrente
  Dim lngQtdTotal     As Long
  Dim curVrTotal      As Currency
  Dim curVrTroco      As Currency
  Dim intI            As Integer
  Dim lngAuxFontSize  As Long
  '
  Dim curVrPrest      As Currency
  Dim curVrCasa       As Currency
  Dim curVrPrestCasa  As Currency
  '
  lngAuxFontSize = Printer.FontSize
  '
  'Dim sVenda As String
  'Dim sCobranca As String
  'Dim sData As String
  'Dim sTipo As String
  '
  For intI = 1 To intQtdVias
    Set objGeral = New busSisMed.clsGeral
    '
    strSql = "SELECT GR.DESCRICAO, GR.SENHA, FUNCIONARIO.NOME AS FUNCIONARIO, PRONTUARIO.DTNASCIMENTO, PRONTUARIO.TELEFONE, FUNCIONARIODET.NIVEL, GR.PKID, GR.SEQUENCIAL, SALA.NUMERO, GR.DATA, PRONTUARIO.NOME AS PACIENTE, PRONTUARIO.PKID AS PACIENTEID, PRESTADOR.NOME AS PRESTADOR, GRPROCEDIMENTO.QTD, GRPROCEDIMENTO.VALOR, PROCEDIMENTO.PROCEDIMENTO " & _
      " FROM GR " & _
      " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
      "   INNER JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
      "   INNER JOIN PRONTUARIO AS PRESTADOR ON PRESTADOR.PKID = ATENDE.PRONTUARIOID " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = GR.PRONTUARIOID " & _
      " LEFT JOIN GRPROCEDIMENTO ON GR.PKID = GRPROCEDIMENTO.GRID " & _
      " LEFT JOIN PROCEDIMENTO ON PROCEDIMENTO.PKID = GRPROCEDIMENTO.PROCEDIMENTOID " & _
      " LEFT JOIN PRONTUARIO AS FUNCIONARIO ON FUNCIONARIO.PKID = GR.FUNCIONARIOID " & _
      " LEFT JOIN FUNCIONARIO AS FUNCIONARIODET ON FUNCIONARIODET.PRONTUARIOID = GR.FUNCIONARIOID " & _
      "WHERE GR.PKID = " & lngGRID
  
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      
      '----- IMPRIMIR HEADER
  
'''      IMPRIMIR_LINHA 1
'''      QUEBRA_LINHA "           GUIA DE RECOLHIMENTO           ", lLARGURACAIXA - (dLSTART / 2)
'''      'QUEBRA_LINHA strEmpresa, lLARGURACAIXA - (dLSTART / 2)
'''      IMPRIMIR_LINHA 1
'''      PULAR_LINHA 2
'''      '
'''      Printer.FontSize = 14
'''      QUEBRA_LINHA "SENHA " & Format(objRs.Fields("SENHA").Value, "###,000"), lLARGURACAIXA - (dLSTART / 2)
'''      Printer.FontSize = lngAuxFontSize
'''      PULAR_LINHA 2
      '
      IMPRIMIR_LINHA 1
      QUEBRA_LINHA "GUIA DE RECOLHIMENTO - " & IIf(bln2Via = True, "2", "1") & " VIA DE GR", lLARGURACAIXA - (dLSTART / 2)
      Printer.FontSize = 14
      QUEBRA_LINHA "SENHA " & Format(objRs.Fields("SENHA").Value, "###,000"), lLARGURACAIXA - (dLSTART / 2)
      Printer.FontSize = lngAuxFontSize
      IMPRIMIR_LINHA 1
      PULAR_LINHA 1
      '
      'QUEBRA_LINHA "GR : " & objRs.Fields("PKID").Value & " ORDEM : " & objRs.Fields("SEQUENCIAL").Value & " SALA : " & objRs.Fields("NUMERO").Value, lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA "GR : " & objRs.Fields("SEQUENCIAL").Value & " SALA : " & objRs.Fields("NUMERO").Value, lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA "FUNCIONARIO : " & objRs.Fields("FUNCIONARIO").Value, lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA "DATA : " & Format(objRs.Fields("DATA").Value, "DD/MM/YYYY - hh:mm"), lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA "DESCRICAO : " & Format(objRs.Fields("DESCRICAO").Value, "DD/MM/YYYY - hh:mm"), lLARGURACAIXA - (dLSTART / 2)
      
      IMPRIMIR_LINHA 1
      Printer.FontSize = 14
      QUEBRA_LINHA "Prontuario : " & objRs.Fields("PACIENTEID").Value & " - " & objRs.Fields("PACIENTE").Value, lLARGURACAIXA - (dLSTART / 2)
      'If objRs.Fields("NIVEL").Value & "" = gsLaboratorio Then
        QUEBRA_LINHA "Telefone : " & objRs.Fields("TELEFONE").Value & "     Nascimento : " & Format(objRs.Fields("DTNASCIMENTO").Value, "DD/MM/YYYY"), lLARGURACAIXA - (dLSTART / 2)
      Printer.FontSize = lngAuxFontSize
      'End If
      IMPRIMIR_LINHA 1
      Printer.FontSize = 14
      QUEBRA_LINHA "Prestador : " & objRs.Fields("PRESTADOR").Value, lLARGURACAIXA - (dLSTART / 2)
      Printer.FontSize = lngAuxFontSize
      IMPRIMIR_LINHA 1
      
      '----- IMPRIMIR DETALHES
      
      PULAR_LINHA 1
      IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, "Valor", "Right", False
      'IMPRIMIR_POSICAO_CORRETA 0.5, lLARGURACAIXA - 2.2, Printer.CurrentY, "Qd", "Right", False
      'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 3.9, Printer.CurrentY, "Uni", "Right", False
      IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Procedimento", "Left", True
      '
      IMPRIMIR_LINHA 0
      '---
      lngQtdTotal = 0
      curVrTotal = 0
      curVrTroco = 0
      '
      Do While Not objRs.EOF
        'Soma Totais
        lngQtdTotal = lngQtdTotal + IIf(Not IsNumeric(objRs.Fields("QTD").Value), 0, objRs.Fields("QTD").Value)
        curVrTotal = curVrTotal + IIf(Not IsNumeric(objRs.Fields("VALOR").Value), 0, objRs.Fields("VALOR").Value)
  
        'QUEBRA_LINHA , lLARGURACAIXA - (dLSTART / 2)
        IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(objRs.Fields("VALOR").Value, "#,##0.00"), "Right", False
        'IMPRIMIR_POSICAO_CORRETA 0.5, lLARGURACAIXA - 2.2, Printer.CurrentY, Format(objRs.Fields("QTD").Value, "#,##0"), "Right", False
        'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 3.9, Printer.CurrentY, Format(objRs!VALORCARDAPIO, "#,##0.00"), "Right", False
        'IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, objRs!Codigo & "/" & objRs!DESCRICAO, "Left", True
        IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, objRs.Fields("PROCEDIMENTO").Value & "", "Left", True
        objRs.MoveNext
      Loop
      '---
      IMPRIMIR_LINHA 0
      '
  
      IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curVrTotal, "#,##0.00"), "Right", False
      'IMPRIMIR_POSICAO_CORRETA 0.5, lLARGURACAIXA - 2.2, Printer.CurrentY, Format(lngQtdTotal, "#,##0"), "Right", False
      'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 3.9, Printer.CurrentY, Format(objRs![CARDAPIO.Valor], "#,##0.00"), "Right", False
      IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "VALOR TOTAL", "Left", True
      '
      PULAR_LINHA 1
      IMPRIMIR_LINHA 0
      objRs.Close
      Set objRs = Nothing
      '---
      'PAGAMENTO
      Set objCC = New busSisMed.clsContaCorrente
      Set objRs = objCC.SelecionarPagamentos("RC", _
                                             lngGRID)
      
      If objRs.EOF Then
        IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Nao houve pagamento", "Left", True
      End If
      Do While Not objRs.EOF
        IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(objRs.Fields("VALOR").Value, "###,##0.00"), "Right", False
        
        If objRs.Fields("INDCONVENIO").Value & "" = "S" Then
          'CONVÊNIO
          'IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Convenio", "Left", True
          IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, objRs.Fields("NOME_CARTAO_DEBITO").Value & "", "Left", True
        Else
          IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, objRs.Fields("DESC_STATUSCC").Value, "Left", True
        End If
        curVrTroco = curVrTroco + IIf(Not IsNumeric(objRs.Fields("VRTROCO").Value), 0, objRs.Fields("VRTROCO").Value)
        objRs.MoveNext
      Loop
      If curVrTroco > 0 Then
        IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curVrTroco, "###,##0.00"), "Right", False
        IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Troco", "Left", True
      End If
      objRs.Close
      Set objRs = Nothing
      Set objCC = Nothing
      'Tratar totais casa / prestador
      Retorna_totais_GR lngGRID, _
                        curVrPrest, _
                        curVrCasa, _
                        curVrPrestCasa
      PULAR_LINHA 1
      IMPRIMIR_LINHA 0
      'CASA
      IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curVrCasa, "#,##0.00"), "Right", False
      IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "CASA", "Left", True
      'PRESTADOR
      IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curVrPrest, "#,##0.00"), "Right", False
      IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "PRESTADOR", "Left", True
      '
      '---
      PULAR_LINHA 2
    
    End If
  Next
  Set objGeral = Nothing
  For intI = 1 To intQtdPontos
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, ".", "Left", True
  Next
End Sub

Public Sub COMPROV_GR_A5(ByVal lngGRID As Long, _
                         strEmpresa As String, _
                         intQtdVias As Integer, _
                         bln2Via As Boolean)
  'Imprimir corpo
  Dim strSql          As String
  Dim objRs           As ADODB.Recordset
  Dim objGeral        As busSisMed.clsGeral
  Dim objCC           As busSisMed.clsContaCorrente
  Dim lngQtdTotal     As Long
  Dim curVrTotal      As Currency
  Dim curVrTroco      As Currency
  Dim intI            As Integer
  Dim lngAuxFontSize  As Long
  '
  Dim curVrPrest      As Currency
  Dim curVrCasa       As Currency
  Dim curVrPrestCasa  As Currency
  '
  lngAuxFontSize = Printer.FontSize
  '
  'Dim sVenda As String
  'Dim sCobranca As String
  'Dim sData As String
  'Dim sTipo As String
  '
  For intI = 1 To intQtdVias
    Set objGeral = New busSisMed.clsGeral
    '
    strSql = "SELECT ESPECIALIDADE.ESPECIALIDADE, GR.DESCRICAO, GR.SENHA, FUNCIONARIO.NOME AS FUNCIONARIO, PRONTUARIO.DTNASCIMENTO, PRONTUARIO.TELEFONE, FUNCIONARIODET.NIVEL, GR.PKID, GR.SEQUENCIAL, SALA.NUMERO, GR.DATA, PRONTUARIO.NOME AS PACIENTE, PRONTUARIO.PKID AS PACIENTEID, PRESTADOR.NOME AS PRESTADOR, GRPROCEDIMENTO.QTD, GRPROCEDIMENTO.VALOR, PROCEDIMENTO.PROCEDIMENTO " & _
      " FROM GR " & _
      " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
      "   INNER JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
      "   INNER JOIN PRONTUARIO AS PRESTADOR ON PRESTADOR.PKID = ATENDE.PRONTUARIOID " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = GR.PRONTUARIOID " & _
      " INNER JOIN ESPECIALIDADE ON ESPECIALIDADE.PKID = GR.ESPECIALIDADEID " & _
      " LEFT JOIN GRPROCEDIMENTO ON GR.PKID = GRPROCEDIMENTO.GRID " & _
      " LEFT JOIN PROCEDIMENTO ON PROCEDIMENTO.PKID = GRPROCEDIMENTO.PROCEDIMENTOID " & _
      " LEFT JOIN PRONTUARIO AS FUNCIONARIO ON FUNCIONARIO.PKID = GR.FUNCIONARIOID " & _
      " LEFT JOIN FUNCIONARIO AS FUNCIONARIODET ON FUNCIONARIODET.PRONTUARIOID = GR.FUNCIONARIOID " & _
      "WHERE GR.PKID = " & lngGRID
  
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      
      '----- IMPRIMIR HEADER
      Printer.FontBold = True
      QUEBRA_LINHA strEmpresa, lLARGURACAIXAA5 - (dLSTART / 2)
      PULAR_LINHA 2
      QUEBRA_LINHA "BOLETIM DE ATENDIMENTO - " & IIf(bln2Via = True, "2", "1") & " VIA", lLARGURACAIXAA5 - (dLSTART / 2)
      Printer.FontBold = False
      PULAR_LINHA 1
      QUEBRA_LINHA "DATA: " & Format(objRs.Fields("DATA").Value, "DD/MM/YYYY") & "                        HORA : " & Format(objRs.Fields("DATA").Value, "hh:mm"), lLARGURACAIXAA5 - (dLSTART / 2)
      Printer.FontSize = 18
      QUEBRA_LINHA "SENHA " & Format(objRs.Fields("SENHA").Value, "###,000") & "          SALA : " & objRs.Fields("NUMERO").Value & "          GR : " & objRs.Fields("SEQUENCIAL").Value, lLARGURACAIXAA5 - (dLSTART / 2)
      Printer.FontSize = lngAuxFontSize
      'IMPRIMIR_LINHA 1
      'PULAR_LINHA 1
      QUEBRA_LINHA "PACIENTE: " & objRs.Fields("PACIENTE").Value & "          Nro Prontuario : " & objRs.Fields("PACIENTEID").Value, lLARGURACAIXAA5 - (dLSTART / 2)
      QUEBRA_LINHA "NASCIMENTO: " & Format(objRs.Fields("DTNASCIMENTO").Value, "DD/MM/YYYY") & "          TELEFONE: " & objRs.Fields("TELEFONE").Value, lLARGURACAIXAA5 - (dLSTART / 2)
      QUEBRA_LINHA "ESPECIALIDADE: " & objRs.Fields("ESPECIALIDADE").Value, lLARGURACAIXAA5 - (dLSTART / 2)
      Printer.FontSize = 18
      QUEBRA_LINHA "PRESTADOR: " & objRs.Fields("PRESTADOR").Value, lLARGURACAIXAA5 - (dLSTART / 2)
      Printer.FontSize = 8
      QUEBRA_LINHA "RECEPCIONISTA: " & objRs.Fields("FUNCIONARIO").Value, lLARGURACAIXAA5 - (dLSTART / 2)
      Printer.FontSize = lngAuxFontSize
      
      
      
      
      
      
'''      '
'''      IMPRIMIR_LINHA 1
'''      QUEBRA_LINHA "GUIA DE RECOLHIMENTO - " & IIf(bln2Via = True, "2", "1") & " VIA DE GR", lLARGURACAIXAA5 - (dLSTART / 2)
'''      Printer.FontSize = 14
'''      QUEBRA_LINHA "SENHA " & Format(objRs.Fields("SENHA").Value, "###,000"), lLARGURACAIXAA5 - (dLSTART / 2)
'''      Printer.FontSize = lngAuxFontSize
'''      IMPRIMIR_LINHA 1
'''      PULAR_LINHA 1
'''      '
'''      'QUEBRA_LINHA "GR : " & objRs.Fields("PKID").Value & " ORDEM : " & objRs.Fields("SEQUENCIAL").Value & " SALA : " & objRs.Fields("NUMERO").Value, lLARGURACAIXAA5 - (dLSTART / 2)
'''      QUEBRA_LINHA "GR : " & objRs.Fields("SEQUENCIAL").Value & " SALA : " & objRs.Fields("NUMERO").Value, lLARGURACAIXAA5 - (dLSTART / 2)
'''      QUEBRA_LINHA "FUNCIONARIO : " & objRs.Fields("FUNCIONARIO").Value, lLARGURACAIXAA5 - (dLSTART / 2)
'''      QUEBRA_LINHA "DATA : " & Format(objRs.Fields("DATA").Value, "DD/MM/YYYY - hh:mm"), lLARGURACAIXAA5 - (dLSTART / 2)
'''      QUEBRA_LINHA "DESCRICAO : " & Format(objRs.Fields("DESCRICAO").Value, "DD/MM/YYYY - hh:mm"), lLARGURACAIXAA5 - (dLSTART / 2)
'''
'''      IMPRIMIR_LINHA 1
'''      Printer.FontSize = 14
'''      QUEBRA_LINHA "Prontuario : " & objRs.Fields("PACIENTEID").Value & " - " & objRs.Fields("PACIENTE").Value, lLARGURACAIXAA5 - (dLSTART / 2)
'''      'If objRs.Fields("NIVEL").Value & "" = gsLaboratorio Then
'''        QUEBRA_LINHA "Telefone : " & objRs.Fields("TELEFONE").Value & "     Nascimento : " & Format(objRs.Fields("DTNASCIMENTO").Value, "DD/MM/YYYY"), lLARGURACAIXAA5 - (dLSTART / 2)
'''      Printer.FontSize = lngAuxFontSize
'''      'End If
'''      IMPRIMIR_LINHA 1
'''      Printer.FontSize = 14
'''      QUEBRA_LINHA "Prestador : " & objRs.Fields("PRESTADOR").Value, lLARGURACAIXAA5 - (dLSTART / 2)
'''      Printer.FontSize = lngAuxFontSize
'''      IMPRIMIR_LINHA 1
      
      '----- IMPRIMIR DETALHES
      
      PULAR_LINHA 1
      IMPRIMIR_POSICAO_CORRETA1 2, dLSTART + 11, Printer.CurrentY, "Valor", "Right", False
      IMPRIMIR_POSICAO_CORRETA1 10, dLSTART, Printer.CurrentY, "Procedimento", "Left", True
      'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXAA5 - 1.5, Printer.CurrentY, "Valor", "Right", False
      'IMPRIMIR_POSICAO_CORRETA 0.5, lLARGURACAIXAA5 - 2.2, Printer.CurrentY, "Qd", "Right", False
      'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXAA5 - 3.9, Printer.CurrentY, "Uni", "Right", False
      'IMPRIMIR_POSICAO_CORRETA 6, dLSTART, Printer.CurrentY, "Procedimento", "Left", True
      '
      'IMPRIMIR_LINHA 0
      '---
      lngQtdTotal = 0
      curVrTotal = 0
      curVrTroco = 0
      '
      Do While Not objRs.EOF
        'Soma Totais
        lngQtdTotal = lngQtdTotal + IIf(Not IsNumeric(objRs.Fields("QTD").Value), 0, objRs.Fields("QTD").Value)
        curVrTotal = curVrTotal + IIf(Not IsNumeric(objRs.Fields("VALOR").Value), 0, objRs.Fields("VALOR").Value)
  
        'QUEBRA_LINHA , lLARGURACAIXAA5 - (dLSTART / 2)
        IMPRIMIR_POSICAO_CORRETA1 2, dLSTART + 11, Printer.CurrentY, Format(objRs.Fields("VALOR").Value, "#,##0.00"), "Right", False
        IMPRIMIR_POSICAO_CORRETA1 10, dLSTART, Printer.CurrentY, objRs.Fields("PROCEDIMENTO").Value & "", "Left", True
        
        'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXAA5 - 1.5, Printer.CurrentY, Format(objRs.Fields("VALOR").Value, "#,##0.00"), "Right", False
        'IMPRIMIR_POSICAO_CORRETA 0.5, lLARGURACAIXAA5 - 2.2, Printer.CurrentY, Format(objRs.Fields("QTD").Value, "#,##0"), "Right", False
        'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXAA5 - 3.9, Printer.CurrentY, Format(objRs!VALORCARDAPIO, "#,##0.00"), "Right", False
        'IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, objRs!Codigo & "/" & objRs!DESCRICAO, "Left", True
        'IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, objRs.Fields("PROCEDIMENTO").Value & "", "Left", True
        objRs.MoveNext
      Loop
      '---
      'IMPRIMIR_LINHA 0
      '
      IMPRIMIR_POSICAO_CORRETA1 2, dLSTART + 11, Printer.CurrentY, Format(curVrTotal, "#,##0.00"), "Right", False
      IMPRIMIR_POSICAO_CORRETA1 10, dLSTART, Printer.CurrentY, "VALOR TOTAL", "Left", True
  
      'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXAA5 - 1.5, Printer.CurrentY, Format(curVrTotal, "#,##0.00"), "Right", False
      'IMPRIMIR_POSICAO_CORRETA 0.5, lLARGURACAIXAA5 - 2.2, Printer.CurrentY, Format(lngQtdTotal, "#,##0"), "Right", False
      'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXAA5 - 3.9, Printer.CurrentY, Format(objRs![CARDAPIO.Valor], "#,##0.00"), "Right", False
      'IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "VALOR TOTAL", "Left", True
      '
'''      PULAR_LINHA 1
'''      'IMPRIMIR_LINHA 0
'''      objRs.Close
'''      Set objRs = Nothing
'''      '---
'''      'PAGAMENTO
'''      Set objCC = New busSisMed.clsContaCorrente
'''      Set objRs = objCC.SelecionarPagamentos("RC", _
'''                                             lngGRID)
'''
'''      If objRs.EOF Then
'''        IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Nao houve pagamento", "Left", True
'''      End If
'''      Do While Not objRs.EOF
'''        IMPRIMIR_POSICAO_CORRETA1 2, dLSTART + 11, Printer.CurrentY, Format(objRs.Fields("VALOR").Value, "###,##0.00"), "Right", False
'''
'''        'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXAA5 - 1.5, Printer.CurrentY, Format(objRs.Fields("VALOR").Value, "###,##0.00"), "Right", False
'''
'''        If objRs.Fields("INDCONVENIO").Value & "" = "S" Then
'''          'CONVÊNIO
'''          'IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, objRs.Fields("NOME_CARTAO_DEBITO").Value & "", "Left", True
'''          IMPRIMIR_POSICAO_CORRETA1 10, dLSTART, Printer.CurrentY, objRs.Fields("NOME_CARTAO_DEBITO").Value & "", "Left", True
'''
'''        Else
'''          'IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, objRs.Fields("DESC_STATUSCC").Value, "Left", True
'''          IMPRIMIR_POSICAO_CORRETA1 10, dLSTART, Printer.CurrentY, objRs.Fields("DESC_STATUSCC").Value & "", "Left", True
'''
'''        End If
'''        curVrTroco = curVrTroco + IIf(Not IsNumeric(objRs.Fields("VRTROCO").Value), 0, objRs.Fields("VRTROCO").Value)
'''        objRs.MoveNext
'''      Loop
'''      If curVrTroco > 0 Then
'''        'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXAA5 - 1.5, Printer.CurrentY, Format(curVrTroco, "###,##0.00"), "Right", False
'''        'IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Troco", "Left", True
'''        IMPRIMIR_POSICAO_CORRETA1 2, dLSTART + 11, Printer.CurrentY, Format(curVrTroco, "###,##0.00"), "Right", False
'''        IMPRIMIR_POSICAO_CORRETA1 10, dLSTART, Printer.CurrentY, "Troco", "Left", True
'''      End If
'''      objRs.Close
'''      Set objRs = Nothing
'''      Set objCC = Nothing
'''      'Tratar totais casa / prestador
'''      Retorna_totais_GR lngGRID, _
'''                        curVrPrest, _
'''                        curVrCasa, _
'''                        curVrPrestCasa
'''      PULAR_LINHA 1
'''      'IMPRIMIR_LINHA 0
'''      'CASA
'''      'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXAA5 - 1.5, Printer.CurrentY, Format(curVrCasa, "#,##0.00"), "Right", False
'''      'IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "CASA", "Left", True
'''      IMPRIMIR_POSICAO_CORRETA1 2, dLSTART + 11, Printer.CurrentY, Format(curVrCasa, "#,##0.00"), "Right", False
'''      IMPRIMIR_POSICAO_CORRETA1 10, dLSTART, Printer.CurrentY, "CASA", "Left", True
'''
'''      'PRESTADOR
'''      'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXAA5 - 1.5, Printer.CurrentY, Format(curVrPrest, "#,##0.00"), "Right", False
'''      'IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "PRESTADOR", "Left", True
'''      IMPRIMIR_POSICAO_CORRETA1 2, dLSTART + 11, Printer.CurrentY, Format(curVrPrest, "#,##0.00"), "Right", False
'''      IMPRIMIR_POSICAO_CORRETA1 10, dLSTART, Printer.CurrentY, "PRESTADOR", "Left", True
'''      '
      '---
      PULAR_LINHA 2
    
    End If
  Next
  Set objGeral = Nothing
'''  For intI = 1 To intQtdPontos
'''    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, ".", "Left", True
'''  Next
End Sub

Public Sub COMPROV_CANC_GR(ByVal lngGRID As Long, _
                           strEmpresa As String, _
                           intQtdVias As Integer)
  'Imprimir corpo
  Dim strSql          As String
  Dim objRs           As ADODB.Recordset
  Dim objGeral        As busSisMed.clsGeral
  Dim objCC           As busSisMed.clsContaCorrente
  Dim lngQtdTotal     As Long
  Dim curVrTotal      As Currency
  Dim curVrTroco      As Currency
  Dim intI            As Integer
  Dim lngAuxFontSize  As Long
  
  lngAuxFontSize = Printer.FontSize
  '
  'Dim sVenda As String
  'Dim sCobranca As String
  'Dim sData As String
  'Dim sTipo As String
  '
  For intI = 1 To intQtdVias
    Set objGeral = New busSisMed.clsGeral
    '
    strSql = "SELECT GR.SENHA, FUNCIONARIO.NOME AS FUNCIONARIO, PRONTUARIO.DTNASCIMENTO, PRONTUARIO.TELEFONE, FUNCIONARIODET.NIVEL, GR.PKID, GR.SEQUENCIAL, SALA.NUMERO, GR.DATA, PRONTUARIO.NOME AS PACIENTE, PRONTUARIO.PKID AS PACIENTEID, PRESTADOR.NOME AS PRESTADOR, GRPROCEDIMENTO.QTD, GRPROCEDIMENTO.VALOR, PROCEDIMENTO.PROCEDIMENTO " & _
      " FROM GR " & _
      " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
      "   INNER JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
      "   INNER JOIN PRONTUARIO AS PRESTADOR ON PRESTADOR.PKID = ATENDE.PRONTUARIOID " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = GR.PRONTUARIOID " & _
      " LEFT JOIN GRPROCEDIMENTO ON GR.PKID = GRPROCEDIMENTO.GRID " & _
      " LEFT JOIN PROCEDIMENTO ON PROCEDIMENTO.PKID = GRPROCEDIMENTO.PROCEDIMENTOID " & _
      " LEFT JOIN PRONTUARIO AS FUNCIONARIO ON FUNCIONARIO.PKID = GR.FUNCIONARIOID " & _
      " LEFT JOIN FUNCIONARIO AS FUNCIONARIODET ON FUNCIONARIODET.PRONTUARIOID = GR.FUNCIONARIOID " & _
      "WHERE GR.PKID = " & lngGRID
  
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      
      '----- IMPRIMIR HEADER
  
      IMPRIMIR_LINHA 1
      QUEBRA_LINHA "---------- COMPROVANTE DE CANCELAMENTO ----------", lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA strEmpresa, lLARGURACAIXA - (dLSTART / 2)
      IMPRIMIR_LINHA 1
      PULAR_LINHA 2
      '
      Printer.FontSize = 14
      QUEBRA_LINHA "SENHA " & Format(objRs.Fields("SENHA").Value, "###,000"), lLARGURACAIXA - (dLSTART / 2)
      Printer.FontSize = lngAuxFontSize
      PULAR_LINHA 2
      '
      IMPRIMIR_LINHA 1
      QUEBRA_LINHA "CANCELADO POR : " & gsNomeUsu, lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA "LIBERADO POR : " & IIf(gsNomeUsuLib = "", gsNomeUsu, gsNomeUsuLib), lLARGURACAIXA - (dLSTART / 2)
      IMPRIMIR_LINHA 1
      PULAR_LINHA 2
      '
      'QUEBRA_LINHA "GR : " & objRs.Fields("PKID").Value & " ORDEM : " & objRs.Fields("SEQUENCIAL").Value & " SALA : " & objRs.Fields("NUMERO").Value, lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA "GR : " & objRs.Fields("SEQUENCIAL").Value & " SALA : " & objRs.Fields("NUMERO").Value, lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA "FUNCIONARIO : " & objRs.Fields("FUNCIONARIO").Value, lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA "DATA : " & Format(Now, "DD/MM/YYYY - hh:mm"), lLARGURACAIXA - (dLSTART / 2)
      IMPRIMIR_LINHA 1
      objRs.Close
      Set objRs = Nothing
      Set objCC = Nothing
      '---
      PULAR_LINHA 2
    End If
  Next
  Set objGeral = Nothing
  For intI = 1 To intQtdPontos
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, ".", "Left", True
  Next
End Sub


Public Sub IMPRIMIR_LINHA(ByVal pTipo As Integer)
  Dim sLinha As String
  'pTipo Assume 0 Simples, 1 - Dupla
  sLinha = "--------------------------------------"
  
  If pTipo = 1 Then
    sLinha = Replace(sLinha, "-", "=")
  End If
  Printer.CurrentX = lPOSX + dLSTART ' + dLEND
  Printer.Print sLinha
End Sub
Sub QUEBRA_LINHA(ByVal sTexto As String, dLargura As Double)
'sTexto retorna a parte da esquerda que cabe numa linha
'sTexto retorna o resto
   Dim iPosCRLf As Integer
   Dim sLinhaTexto As String
   Dim sAux As String
   Dim sLinha As String
   Dim sResto As String
   Dim sTexto2 As String
   Dim sLinhaAux As String
   Dim sRestoAux As String
   Dim iPosEspaco As Long
   Dim blnSemEspaco As Boolean
   
   'Printer.CurrentX = dPoxX
   
   'Pega primeira linha
   iPosCRLf = InStr(sTexto, vbCrLf)
   If iPosCRLf = 0 Then
      iPosCRLf = Len(sTexto)
   End If
   'If len(sTexto)
   If Mid(sTexto, IIf(iPosCRLf = 0, 1, iPosCRLf), 2) = vbCrLf Then
    sLinhaTexto = Left$(sTexto, IIf(iPosCRLf - 1 = -1, 0, iPosCRLf - 1))
   Else
    sLinhaTexto = Left$(sTexto, iPosCRLf)
   End If
   If Len(sTexto) = iPosCRLf Then
    sTexto = ""
   Else
    sTexto = Right$(sTexto, Len(sTexto) - iPosCRLf - 1)
   End If
   
   'Enquanto tiver quebra de linha
   While iPosCRLf <> 0
      'Se a linha nao passar do limite, imprime
      If Printer.TextWidth(sLinhaTexto) <= (dLargura - 0.7) Then
        '
        'VerificaQuebra
        '
         Printer.CurrentX = lPOSX + dLSTART ' + dLEND
         Printer.Print sLinhaTexto
      Else
         
         While Printer.TextWidth(sLinhaTexto) > (dLargura - 0.7)
            
            'Obtem pedaco da linha que cabe na largura
            sLinhaAux = ""
            sRestoAux = sLinhaTexto
            blnSemEspaco = False
            While Printer.TextWidth(sLinhaAux) < (dLargura - 0.7)
               iPosEspaco = InStr(sRestoAux, " ")
               If iPosEspaco = 0 Then
                  iPosEspaco = Len(sRestoAux)
                  blnSemEspaco = True
               End If
               'Guarda linha e resto atuais
               sLinha = sLinhaAux
               sResto = sRestoAux
               'coloca mais uma palavra
               sLinhaAux = sLinhaAux & Left$(sRestoAux, iPosEspaco)
               sRestoAux = Right$(sRestoAux, Len(sRestoAux) - iPosEspaco)
            Wend
            'NOVO Verifica se texto nao possui espaco
            If blnSemEspaco = True And Len(sResto) <> 0 And Len(sLinha) = 0 Then
              sLinha = sLinhaAux
              sResto = sRestoAux
            End If
            'VerificaQuebra
            '
            'Imprime o pedaco que cabe e guarda o resto
            Printer.CurrentX = lPOSX + dLSTART ' + dLEND
            Printer.Print sLinha
            sLinhaTexto = sResto
         
         Wend
        '
        'VerificaQuebra
        '
        'Imprime o pedaco restante
        Printer.CurrentX = lPOSX + dLSTART '+ dLEND
        Printer.Print sLinhaTexto
      End If
      
      'Pega proxima linha
      iPosCRLf = InStr(sTexto, vbCrLf)
      If iPosCRLf = 0 Then
         iPosCRLf = Len(sTexto)
      End If
      If iPosCRLf <> 0 Then
         If Mid(sTexto, IIf(iPosCRLf = 0, 1, iPosCRLf), 2) = vbCrLf Then
          sLinhaTexto = Left$(sTexto, iPosCRLf - 1)
         Else
          sLinhaTexto = Left$(sTexto, iPosCRLf)
         End If
         If Len(sTexto) = iPosCRLf Then
          sTexto = ""
         Else
          sTexto = Right$(sTexto, Len(sTexto) - iPosCRLf - 1)
         End If
      Else
         sLinhaTexto = ""
         sTexto = ""
      End If
   
   Wend

End Sub

Public Sub IMPRIMIR_POSICAO_CORRETA(ByVal pComprimento As Double, ByVal pPosX As Double, ByVal pPosY As Double, ByVal pTexto As String, ByVal pAlign As String, ByVal pPularLinha As Boolean)
  'Rotina para imprimir uma frase na posicao correta da pagina
  Dim CurX As Double
  Dim CurY As Double
  '
  'Guarda pos inicial
  CurX = Printer.CurrentX
  CurY = Printer.CurrentY
  '
  Printer.CurrentY = pPosY
  If pAlign = "Left" Then
    Printer.CurrentX = pPosX
  Else
    Printer.CurrentX = pPosX + pComprimento - Printer.TextWidth(pTexto)
  End If
  '
  If pAlign = "Left" Then
    QUEBRA_LINHA pTexto, pComprimento - (dLSTART / 2)
  Else
    Printer.Print pTexto
  End If
  If pPularLinha Then
    Printer.CurrentX = CurX
  Else
    'Restaura posicao inicial
    Printer.CurrentX = CurX
    Printer.CurrentY = CurY
  End If
  
End Sub

Public Sub IMPRIMIR_POSICAO_CORRETA1(ByVal pComprimento As Double, ByVal pPosX As Double, ByVal pPosY As Double, ByVal pTexto As String, ByVal pAlign As String, ByVal pPularLinha As Boolean)
  'Rotina para imprimir uma frase na posicao correta da pagina
  Dim CurX As Double
  Dim CurY As Double
  '
  'Guarda pos inicial
  CurX = Printer.CurrentX
  CurY = Printer.CurrentY
  '
  Printer.CurrentY = pPosY
  If pAlign = "Left" Then
    Printer.CurrentX = pPosX
  Else
    Printer.CurrentX = pPosX - Printer.TextWidth(pTexto)
  End If
  '
  If pAlign = "Left" Then
    QUEBRA_LINHA pTexto, pComprimento
  Else
    Printer.Print pTexto
  End If
  If pPularLinha Then
    Printer.CurrentX = CurX
  Else
    'Restaura posicao inicial
    Printer.CurrentX = CurX
    Printer.CurrentY = CurY
  End If
  
End Sub

Public Sub IMPRIMIR_NA_POSICAO(ByVal pComprimento As Double, ByVal pPosX As Double, ByVal pPosY As Double, ByVal pTexto As String, ByVal pAlign As String, ByVal pPularLinha As Boolean)
  'Rotina para imprimir uma frase na posicao correta da pagina
  Dim CurX As Double
  Dim CurY As Double
  '
  'Guarda pos inicial
  CurX = Printer.CurrentX
  CurY = Printer.CurrentY
  '
  Printer.CurrentY = pPosY
  If pAlign = "Left" Then
    Printer.CurrentX = pPosX
  Else
    Printer.CurrentX = pPosX + pComprimento - Printer.TextWidth(pTexto)
  End If
  '
  If pAlign = "Left" Then
    'QUEBRA_LINHA pTexto, pComprimento - (dLSTART / 2)
    Printer.Print pTexto
  Else
    Printer.Print pTexto
  End If
  If pPularLinha Then
    Printer.CurrentX = CurX
  Else
    'Restaura posicao inicial
    Printer.CurrentX = CurX
    Printer.CurrentY = CurY
  End If
  
End Sub

Public Property Let FontSize(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontSize = 5
    mvarFontSize = vData
End Property


Public Property Get FontSize() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontSize
    FontSize = mvarFontSize
End Property



Public Property Let Copies(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Copies = 5
    mvarCopies = vData
End Property


Public Property Get Copies() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Copies
    Copies = mvarCopies
End Property



Public Property Let DrawStyle(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DrawStyle = 5
    mvarDrawStyle = vData
End Property


Public Property Get DrawStyle() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DrawStyle
    DrawStyle = mvarDrawStyle
End Property



Public Property Let FontName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontName = 5
    mvarFontName = vData
End Property


Public Property Get FontName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontName
    FontName = mvarFontName
End Property



Public Property Let CurrentY(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CuurentY = 5
    mvarCurrentY = vData
End Property


Public Property Get CurrentY() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CuurentY
    CurrentY = mvarCurrentY
End Property



Public Property Let CurrentX(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CurrentX = 5
    mvarCurrentX = vData
End Property


Public Property Get CurrentX() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrentX
    CurrentX = mvarCurrentX
End Property



Public Property Let ScaleMode(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ScaleMode = 5
    mvarScaleMode = vData
End Property


Public Property Get ScaleMode() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ScaleMode
    ScaleMode = mvarScaleMode
End Property


