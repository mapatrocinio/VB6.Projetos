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
Const dLSTART = 1.3
'Espacamento
Const lDELTAX = 0.5
Const lDELTAY = 0.5
'Constantes das caixas
'Posicao Quadrantes 1 e 3
Const lPOSX = 0
Const lPOSY = 0 + lDELTAY

Const intQtdPontos = 10

'MS Windows API Function Prototypes
Public Declare Function GetProfileString Lib "kernel32" Alias _
     "GetProfileStringA" (ByVal lpAppName As String, _
     ByVal LpKeyName As String, ByVal lpDefault As String, _
     ByVal lpReturnedString As String, _
     ByVal nSize As Long) As Long

'---------------------------------------------------------------
' Retreive the vb object "printer" corresponding to the window's
' default printer.
'---------------------------------------------------------------
Public Function GetDefaultPrinter(Optional strImpressora As String) As Printer
  On Error GoTo trata
  Dim strBuffer As String * 254
  Dim iRetValue As Long
  Dim strDefaultPrinterInfo As String
  Dim tblDefaultPrinterInfo() As String
  Dim objPrinter As Printer
  
  ' Retreive current default printer information
  iRetValue = GetProfileString("windows", "device", ",,,", _
              strBuffer, 254)
  strDefaultPrinterInfo = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
  tblDefaultPrinterInfo = Split(strDefaultPrinterInfo, ",")
  For Each objPrinter In Printers
    If objPrinter.DeviceName = IIf(strImpressora = "", tblDefaultPrinterInfo(0), strImpressora) Then
      ' Default printer found !
      Exit For
    End If
  Next
  
  ' If not found, return nothing
  'If objPrinter.DeviceName <> tblDefaultPrinterInfo(0) Then
  '    Set objPrinter = Nothing
  'End If
  
  Set GetDefaultPrinter = objPrinter
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserImpressao.GetPrinters]", _
            Err.Description
End Function
Public Sub GetPrinters(objList As ListBox, _
                       strImpressoraPadrao As String)
  On Error GoTo trata
  Dim objPrinter As Printer
  
  For Each objPrinter In Printers
    objList.AddItem objPrinter.DeviceName
  Next
  If strImpressoraPadrao <> "" Then
    On Error Resume Next
    objList.Text = strImpressoraPadrao
  End If
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[mdlUserImpressao.GetPrinters]", _
            Err.Description
End Sub


Public Sub TERMINA_IMPRESSAO()
  'Enviando para a impressora
  Printer.EndDoc
End Sub



Public Function FormataSaida(pDesc As String)
  FormataSaida = ""
  If pDesc <> "" Then
    FormataSaida = " | " & pDesc
  End If
End Function

Public Sub ESPACO(ByVal pQtd As Integer)
  Printer.CurrentX = Printer.CurrentX + (Printer.TextWidth("X") * pQtd)
End Sub

Public Sub PULAR_LINHA(ByVal pQtd As Integer)
  Printer.CurrentY = Printer.CurrentY + (Printer.TextHeight("X") * pQtd)
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


Private Sub TratarPos(lngPos As Long, _
                      lngPosH As Long, _
                      lngPosD As Long, _
                      strHeader As String, _
                      curDetail As Currency)
  On Error GoTo trata
  Dim intPos1H        As Integer
  Dim intPos1D        As Integer
  Dim intPos2H        As Integer
  Dim intPos2D        As Integer
  Dim intPos3H        As Integer
  Dim intPos3D        As Integer
  Dim lngLARGURACAIXA As Long
  '
'  lngLARGURACAIXA = 11
'  intPos3H = lngLARGURACAIXA - 2
'  intPos3D = lngLARGURACAIXA - 0
'  intPos2H = lngLARGURACAIXA - 6
'  intPos2D = lngLARGURACAIXA - 4
'  intPos1H = lngLARGURACAIXA - 10
'  intPos1D = lngLARGURACAIXA - 8
  lngLARGURACAIXA = 9
  intPos1H = lngLARGURACAIXA - 9
  intPos1D = lngLARGURACAIXA - 7
  intPos2H = lngLARGURACAIXA - 5
  intPos2D = lngLARGURACAIXA - 3
  intPos3H = lngLARGURACAIXA - 1
  intPos3D = lngLARGURACAIXA - 0
  '
  If lngPos = 1 Then
    lngPosH = intPos1H
    lngPosD = intPos1D
  ElseIf lngPos = 2 Then
    lngPosH = intPos2H
    lngPosD = intPos2D
  Else
    lngPosH = intPos3H
    lngPosD = intPos3D
  End If
  'Imprimir
  IMPRIMIR_POSICAO_CORRETA 2, lngPosH, Printer.CurrentY, strHeader, "Right", False
  IMPRIMIR_POSICAO_CORRETA 2, lngPosD, Printer.CurrentY, Format(curDetail, "###,##0.00"), "Right", IIf(lngPos = 3, True, False)
  'Imcrementa variáveis
  lngPos = lngPos + 1
  If lngPos > 3 Then lngPos = 1
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[clsImpressao.TratarPos]", _
            Err.Description
End Sub


Sub IMPRIMIR_TEXTO(ByVal strTexto As String, _
                   ByVal sngPoxX As Single, _
                   ByVal sngPoxY As Single)
  On Error GoTo trata
  'dblPoxX em milímetros
  'dblPoxY em milímetros
  'Convertar para milímtros
  '
  Printer.CurrentX = sngPoxX
  Printer.CurrentY = sngPoxY
  Printer.Print strTexto
  'Convertar para centimetros
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Description, _
            "[clsImpressao.IMPRIMIR_TEXTO]"
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

Public Sub COMPROV_FECHA_TURNO(ByVal lngPESSOAID As Long, _
                               ByVal strData As String, _
                               ByVal strFabrica As String)
  'Imprimir corpo
  On Error GoTo trata
  Dim strSql                    As String
  Dim objRs                     As ADODB.Recordset
'''  Dim objRsCred                 As ADODB.Recordset
'''  Dim objRsRelPass              As ADODB.Recordset
'''  Dim objRsCampVendas           As ADODB.Recordset
'''  Dim objRsCampVendasInterno    As ADODB.Recordset
'''  Dim objRsCampVendasInterno1   As ADODB.Recordset
'''  Dim objRsCartao               As ADODB.Recordset
'''  Dim objRsCta                  As ADODB.Recordset
'''  Dim objRsDespesa              As ADODB.Recordset
'''  Dim objRsGorjetas             As ADODB.Recordset

  Dim curPgtoDinheiro           As Currency
  Dim curPgtoCheque             As Currency
  Dim curPgtoCC                 As Currency
  Dim curPgtoCD                 As Currency
  Dim curPgtoFatura             As Currency
'''  Dim curPgtoTroco              As Currency

'''  Dim curPgtoPenhor             As Currency
'''  Dim curVrCalcTaxa             As Currency
  Dim datDataTurno              As Date
'''  '
'''  Dim curPgtoFatura             As Currency
'''  Dim curPgtoReserva            As Currency
'''  Dim curPgtoGorjeta            As Currency
  Dim curTotalRecebimentos      As Currency
'''  Dim curFaturamento            As Currency
'''  Dim curFatDesp                As Currency
'''
'''  Dim curDespTotal              As Currency
'''  Dim curDespTotalChq           As Currency
'''  Dim curDespTotalDin           As Currency
'''  Dim curConsumo                As Currency
'''  Dim curExtras                 As Currency
'''  Dim curRetiradas              As Currency
'''  Dim curPgtoGorj               As Currency
'''  '
'''  Dim curVrAuxExtra             As Currency
'''  Dim curEmCaixaDin             As Currency
'''  '
'''  Dim curVrRecDin               As Currency
'''  Dim curVrRecChq               As Currency
'''  Dim curVrRecCC                As Currency
'''  Dim curVrRecCD                As Currency
'''  Dim curVrRecPen               As Currency
'''  Dim curVrRecTot               As Currency
'''  'Perc
'''  Dim curPercRecDin             As Currency
'''  Dim curPercRecChq             As Currency
'''  Dim curPercRecCC              As Currency
'''  Dim curPercRecCD              As Currency
'''  Dim curPercRecPen             As Currency
'''  Dim curPercRecTot             As Currency
'''  '
'''  Dim curEmCaixaDinGorjeta      As Currency
'''  Dim curEmCaixaChq             As Currency
'''  Dim curEmCaixaCar             As Currency
'''  Dim curEmCaixaCarDeb          As Currency
'''  Dim curEmCaixaPen             As Currency
'''  Dim curEmCaixaRes             As Currency
'''  Dim curEmCaixaTot             As Currency
'''  Dim curDinLiquido             As Currency
'''  Dim strRelaPass               As String
'''  Dim lngInterditadas           As Long
'''  Dim lngQtdContasRec           As Long
'''  Dim strTipoConta              As String
  Dim objGer                    As busSisMetal.clsGeral
'''  Dim objTurno                  As busSisMetal.clsTurno
  Dim I As Integer
  '
  Set objGer = New busSisMetal.clsGeral
  '
  strSql = "SELECT " & _
    " PEDIDOVENDA.PKID, CONVERT(DATETIME, convert(VARCHAR(10), DATA,103), 103) AS DATA, " & _
    " ISNULL(SUM(vw_cons_t_cred_ped.PgtoEspecie), 0) AS PgtoEspecie, " & _
    " ISNULL(SUM(vw_cons_t_cred_ped.PgtoCartao), 0) AS PgtoCartao, " & _
    " ISNULL(SUM(vw_cons_t_cred_ped.PgtoCartaoDeb), 0) AS PgtoCartaoDeb, " & _
    " ISNULL(SUM(vw_cons_t_cred_ped.PgtoCheque), 0) AS PgtoCheque, " & _
    " ISNULL(SUM(vw_cons_t_cred_ped.PgtoFatura), 0) AS PgtoFatura  " & _
    " FROM PEDIDOVENDA " & _
    " LEFT JOIN vw_cons_t_cred_ped ON vw_cons_t_cred_ped.PKID = PEDIDOVENDA.PKID " & _
    "WHERE PEDIDOVENDA.CAIXAID = " & Formata_Dados(lngPESSOAID, tpDados_Longo) & _
    " AND CONVERT(DATETIME, convert(VARCHAR(10), DATA,103), 103) = " & Formata_Dados(strData, tpDados_DataHora) & _
    " GROUP BY PEDIDOVENDA.PKID, CONVERT(DATETIME, convert(VARCHAR(10), DATA,103), 103)"
  '
  
  Set objRs = objGer.ExecutarSQL(strSql)
  Do While Not objRs.EOF
    'Calculo dos campos
    curPgtoDinheiro = IIf(Not IsNumeric(objRs.Fields("PgtoEspecie").Value), 0, objRs.Fields("PgtoEspecie").Value)
    'Calcula troco
'''    curPgtoTroco = IIf(Not IsNumeric(objRsCred.Fields("PgtoTroco").Value), 0, objRsCred.Fields("PgtoTroco").Value)
    '
    curTotalRecebimentos = curPgtoDinheiro
    'Outros Recebimentos e data
    curPgtoFatura = IIf(Not IsNumeric(objRs.Fields("PgtoFatura").Value), 0, objRs.Fields("PgtoFatura").Value)
    curPgtoCheque = IIf(Not IsNumeric(objRs.Fields("PgtoCheque").Value), 0, objRs.Fields("PgtoCheque").Value)
    curPgtoCC = IIf(Not IsNumeric(objRs.Fields("PgtoCartao").Value), 0, objRs.Fields("PgtoCartao").Value)
    curPgtoCD = IIf(Not IsNumeric(objRs.Fields("PgtoCartaoDeb").Value), 0, objRs.Fields("PgtoCartaoDeb").Value)
    datDataTurno = objRs.Fields("Data").Value
    '''curVrCalcTaxa = IIf(Not IsNumeric(objRsCred.Fields("VrCalcTaxa").Value), 0, objRsCred.Fields("VrCalcTaxa").Value)
    '
    curTotalRecebimentos = curTotalRecebimentos + curPgtoCC
    curTotalRecebimentos = curTotalRecebimentos + curPgtoCD
    curTotalRecebimentos = curTotalRecebimentos + curPgtoCheque
    curTotalRecebimentos = curTotalRecebimentos + curPgtoFatura
    '--------
    '--------
    IMPRIMIR_LINHA 1
    QUEBRA_LINHA strFabrica, lLARGURACAIXA - (dLSTART / 2)
    IMPRIMIR_LINHA 1
    QUEBRA_LINHA "FECHAMENTO DO CAIXA", lLARGURACAIXA - (dLSTART / 2)
    PULAR_LINHA 1
    QUEBRA_LINHA "FUNCIONARIO " & gsNomeUsuLib, lLARGURACAIXA - (dLSTART / 2)
    PULAR_LINHA 1
    '------------------------------------
    'Rotina para imprimir na ordem correta
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(objRs.Fields("DATA").Value, "DD/MM/YYYY hh:mm"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Data", "Left", True
    IMPRIMIR_LINHA 0
    PULAR_LINHA 1
    
    'DIFERENCIADO ======================================================
    
    QUEBRA_LINHA "RECEBIMENTOS", lLARGURACAIXA - (dLSTART / 2)
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curPgtoDinheiro, "###,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Dinheiro", "Left", True
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curPgtoCheque, "###,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Cheque", "Left", True
    
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curPgtoCC, "###,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Cartao Cred.", "Left", True
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curPgtoCD, "###,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Cartao Deb.", "Left", True
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curPgtoFatura, "###,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Fatura", "Left", True
    
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curTotalRecebimentos, "###,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "TOTAL", "Left", True
'''    IMPRIMIR_LINHA 0
'''    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curPgtoTroco, "###,##0.00"), "Right", False
'''    IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "(-) Troco", "Left", True
    
    'FIM DIFERENCIADO ======================================================
   
    PULAR_LINHA 1
    objRs.MoveNext
  Loop
  objRs.Close
  Set objRs = Nothing
  IMPRIMIR_LINHA 0
  PULAR_LINHA 2
  Set objGer = Nothing
  For I = 1 To intQtdPontos
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, ".", "Left", True
  Next
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub


Public Sub IMP_COMPROV_FECHA_TURNO(ByVal lngPESSOAID As Long, ByVal strData As String, ByVal strFabrica As String)
  'Impressao do comprovante de entrada na portaria
  'pStatus Assume:
  'R - REIMPRESSAO
  'I - IMPRESSAO
On Error GoTo ErrHandler
  '
  INICIA_IMPRESSAO
  '
  'COMPROV_FECHA_TURNO pTURNOID, strFabrica, pStatus
  COMPROV_FECHA_TURNO lngPESSOAID, strData, strFabrica
  '
  TERMINA_IMPRESSAO
  '
  Exit Sub
ErrHandler:
  MsgBox "O seguinte erro ocorreu: " & Err.Description & ". Erro na impressao do comprovante de fechamento.", vbExclamation, TITULOSISTEMA
  TERMINA_IMPRESSAO
End Sub

Public Sub IMP_COMP_PEDIDO(ByVal pPedidoId As Long, ByVal pFabrica As String)
  'Impressao do comprovante de entrada na portaria
On Error GoTo ErrHandler
  '
  INICIA_IMPRESSAO
  '
  COMPROV_PEDIDO pPedidoId, pFabrica
  '
  TERMINA_IMPRESSAO
  '
  Exit Sub
ErrHandler:
  MsgBox "O seguinte erro ocorreu: " & Err.Description & ". Erro na impressao do comprovante de pedido.", vbExclamation, TITULOSISTEMA
  TERMINA_IMPRESSAO
End Sub

Public Sub COMPROV_PEDIDO(ByVal pPedidoId As Long, ByVal pFabrica As String)
  'Imprimir corpo
  Dim sSql As String
  Dim rs As ADODB.Recordset
'''  Dim rsInterno As ADODB.Recordset
  Dim objGeral As busSisMetal.clsGeral
'''  Dim cQtdTotal  As Currency
'''  Dim cVrTotal  As Currency
'''  Dim blnImprimirSegVia As Boolean
  Dim I As Integer
  '
  Set objGeral = New busSisMetal.clsGeral
  sSql = "SELECT PEDIDOVENDA.*, ITEM_PEDIDOVENDA.QUANTIDADE, ITEM_PEDIDOVENDA.VALOR, PRODUTO.NOME, INSUMO.CODIGO, PRODUTO.PRECO, ISNULL(VALOR_CALC_TOTAL, 0) - ISNULL(VALOR_CALC_DESCONTO, 0) AS VALOR_CALC_TOTAL_CDESC " & _
    " FROM PEDIDOVENDA INNER JOIN ITEM_PEDIDOVENDA ON PEDIDOVENDA.PKID = ITEM_PEDIDOVENDA.PEDIDOVENDAID " & _
    " INNER JOIN PRODUTO ON PRODUTO.INSUMOID = ITEM_PEDIDOVENDA.PRODUTOID " & _
    " INNER JOIN INSUMO ON INSUMO.PKID = PRODUTO.INSUMOID " & _
    "WHERE PEDIDOVENDA.PKID = " & Formata_Dados(pPedidoId, tpDados_Longo)
  Set rs = objGeral.ExecutarSQL(sSql)
  If Not rs.EOF Then
    IMPRIMIR_LINHA 1
    QUEBRA_LINHA "Loja: " & pFabrica, lLARGURACAIXA - (dLSTART / 2)
    IMPRIMIR_LINHA 1
    QUEBRA_LINHA "Pedido: " & rs!PED_NUMERO, lLARGURACAIXA - (dLSTART / 2)
    'PULAR_LINHA 1
'''    QUEBRA_LINHA "TOTAL DO PEDIDO", lLARGURACAIXA - (dLSTART / 2)
    'PULAR_LINHA 1
    QUEBRA_LINHA "DATA HORA " & Format(rs!Data, "DD/MM/YYYY hh:mm"), lLARGURACAIXA - (dLSTART / 2)
    '-----
    'QUEBRA_LINHA "Cod   Ddescricao Unit Qt   Valor", lLARGURACAIXA - (dLSTART / 2)
    'Rotina para imprimir na ordem correta
    PULAR_LINHA 1
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, "Valor", "Right", False
    IMPRIMIR_POSICAO_CORRETA 0.5, lLARGURACAIXA - 2.2, Printer.CurrentY, "Qd", "Right", False
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 3.9, Printer.CurrentY, "Uni", "Right", False
    IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Descricao", "Left", True
    '
    IMPRIMIR_LINHA 0
    '---
    Do While Not rs.EOF
      'Soma Totais
      'QUEBRA_LINHA , lLARGURACAIXA - (dLSTART / 2)
      IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(rs!VALOR, "#,##0.00"), "Right", False
      IMPRIMIR_POSICAO_CORRETA 0.5, lLARGURACAIXA - 2.2, Printer.CurrentY, Format(rs!QUANTIDADE, "#,##0"), "Right", False
      IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 3.9, Printer.CurrentY, Format(rs!PRECO, "#,##0.00"), "Right", False
      'IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, rs!Codigo & "/" & rs!DESCRICAO, "Left", True
      IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, rs!CODIGO & " - " & rs!NOME, "Left", True
      rs.MoveNext
    Loop
    '---
    IMPRIMIR_LINHA 0
    '
    rs.MoveFirst
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(rs!VALOR_CALC_TOTAL, "#,##0.00"), "Right", False
    'IMPRIMIR_POSICAO_CORRETA 0.5, lLARGURACAIXA - 2.2, Printer.CurrentY, Format(cQtdTotal, "#,##0"), "Right", False
    'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 3.9, Printer.CurrentY, Format(rs![CARDAPIO.Valor], "#,##0.00"), "Right", False
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "TOTAL DO PEDIDO", "Left", True
    '
    If IIf(IsNull(rs!VALOR_CALC_DESCONTO), 0, rs!VALOR_CALC_DESCONTO) > 0 Then
      'PULAR_LINHA 1
      IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(rs!VALOR_CALC_DESCONTO, "#,##0.00"), "Right", False
      'IMPRIMIR_POSICAO_CORRETA 0.5, lLARGURACAIXA - 2.2, Printer.CurrentY, Format(cQtdTotal, "#,##0"), "Right", False
      'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 3.9, Printer.CurrentY, Format(rs![CARDAPIO.Valor], "#,##0.00"), "Right", False
      IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "DESCONTO", "Left", True
      '
      'PULAR_LINHA 1
      IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(rs!VALOR_CALC_TOTAL_CDESC, "#,##0.00"), "Right", False
      'IMPRIMIR_POSICAO_CORRETA 0.5, lLARGURACAIXA - 2.2, Printer.CurrentY, Format(cQtdTotal, "#,##0"), "Right", False
      'IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 3.9, Printer.CurrentY, Format(rs![CARDAPIO.Valor], "#,##0.00"), "Right", False
      IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "A PAGAR", "Left", True
      '
    End If
    PULAR_LINHA 1
    '
    rs.Close
    Set rs = Nothing
  End If
  Set objGeral = Nothing
  For I = 1 To intQtdPontos
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, ".", "Left", True
  Next
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





