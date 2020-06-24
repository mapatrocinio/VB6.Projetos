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
Global intQtdPontos As Integer
'MS Windows API Function Prototypes
Public Declare Function GetProfileString Lib "kernel32" Alias _
     "GetProfileStringA" (ByVal lpAppName As String, _
     ByVal LpKeyName As String, ByVal lpDefault As String, _
     ByVal lpReturnedString As String, _
     ByVal nSize As Long) As Long

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
Public Sub GetPrinters(objList As ListBox)
  On Error GoTo trata
  Dim objPrinter As Printer
  
  For Each objPrinter In Printers
    objList.AddItem objPrinter.DeviceName
  Next
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[mdlUserImpressao.GetPrinters]", _
            Err.Description
End Sub
Public Function Selecionar_impressora(strTipo As String)
  On Error GoTo trata
  Dim objFormPrinter    As Apler.frmUserPrinter
  Dim strImpressora     As String
  Dim strKey            As String
  'strTipo assume :
  'NF  REC
  'Capturar do registrer os dados referentes a última
  'impressão
  If strTipo = "NF" Then
    'Nota Fiscal
    strKey = "ImpressoraNF"
  Else
    'Recibo
    strKey = "ImpressoraREC"
  End If
    
  strImpressora = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:=strKey)
  If Len(Trim(strImpressora)) = 0 Then
    'Registro está em branco
    'Chamar ou não form impressora
    Set objFormPrinter = New Apler.frmUserPrinter
    objFormPrinter.strKey = strKey
    objFormPrinter.Show vbModal
    Set objFormPrinter = Nothing
  Else
    'Encontrou no register
    'Tenta setar impressora para impressão da NF ou Recibo
    Set Printer = GetDefaultPrinter(strImpressora)
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserImpressao.GetPrinters]", _
            Err.Description
End Function

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

Public Sub COMPROV_DESPESA(ByVal pDespesaId As Long, ByVal pMotel As String, ByVal pQtdVias As Integer, pDescrTurno As String, ByVal blnTrabalhaComFuncAssoc As Boolean)
  'Imprimir corpo
  Dim sSql      As String
  Dim rs        As ADODB.Recordset
  Dim objRs     As ADODB.Recordset
  Dim objGeral  As busApler.clsGeral
  Dim objCC     As busApler.clsContaCorrente
  '
  Dim VRTotalDespesa As Currency
  Dim sMsg As String
  Dim dDataAtual  As Date
  Dim I As Integer
  Dim sVale As String
  '
  Set objGeral = New busApler.clsGeral
  Set objCC = New busApler.clsContaCorrente
  '
  sSql = "SELECT DESPESA.*, FUNCIONARIO.NOME From " & _
    "DESPESA LEFT JOIN FUNCIONARIO ON FUNCIONARIO.PKID = DESPESA.FUNCIONARIOID " & _
    "WHERE DESPESA.PKID = " & pDespesaId
  Set rs = objGeral.ExecutarSQL(sSql)
  'Set objRs = objCC.SelecionarPagamentos("DE", _
                                         pDespesaId)
  dDataAtual = Now
  Do While Not rs.EOF
    'Calculo dos campos
    VRTotalDespesa = IIf(Not IsNumeric(rs!VR_PAGO), 0, rs!VR_PAGO)
    
    For I = 1 To pQtdVias
      '--------
      IMPRIMIR_LINHA 0
      QUEBRA_LINHA pMotel, lLARGURACAIXA - (dLSTART / 2)
      IMPRIMIR_LINHA 0
      QUEBRA_LINHA rs!SEQUENCIAL, lLARGURACAIXA - (dLSTART / 2)
      PULAR_LINHA 1
      
      QUEBRA_LINHA "DESPESA " & IIf(rs!Vale & "" = "S", " - VALE", ""), lLARGURACAIXA - (dLSTART / 2)
      'QUEBRA_LINHA Format(rs!DT_PAGAMENTO, "DD/MM/YYYY hh:mm"), lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA "" & pDescrTurno, lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA "OPERADOR - " & rs!usuario, lLARGURACAIXA - (dLSTART / 2)
      If rs!usuario & "" <> rs!USUARIOAUTORIZACAO & "" Then
        'So imprime se usuarios diferentes
        QUEBRA_LINHA "AUTORIZACAO - " & rs!USUARIOAUTORIZACAO, lLARGURACAIXA - (dLSTART / 2)
      End If
      
      PULAR_LINHA 2
      IMPRIMIR_LINHA 0
      '------------
      'CABECALHO
      IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, "Valor", "Right", False
      IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Descricao", "Left", True
      '
      IMPRIMIR_LINHA 0
      '---
      'CORPO
      IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(VRTotalDespesa, "#,##0.00"), "Right", False
      IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, rs!Descricao, "Left", True
      '---
      IMPRIMIR_LINHA 0
      '---
      'CORPO
      Set objCC = New busApler.clsContaCorrente
      Set objRs = objCC.SelecionarPagamentos("DE", _
                                             pDespesaId)
      Do While Not objRs.EOF
        IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(objRs.Fields("VALOR").Value, "###,##0.00"), "Right", False
        IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, objRs.Fields("DESC_STATUSCC").Value, "Left", True
        objRs.MoveNext
      Loop
      objRs.Close
      Set objRs = Nothing
      Set objCC = Nothing
      '---
      'FIM DO CORPO
      '------------
      PULAR_LINHA 3
      IMPRIMIR_LINHA 0
      If (rs.Fields("VALE").Value & "" = "S") And (blnTrabalhaComFuncAssoc = True) Then
        QUEBRA_LINHA rs.Fields("NOME").Value & "", lLARGURACAIXA - (dLSTART / 2)
      Else
        QUEBRA_LINHA "Responsavel", lLARGURACAIXA - (dLSTART / 2)
      End If
      
      PULAR_LINHA 3
      IMPRIMIR_LINHA 0
      QUEBRA_LINHA pMotel, lLARGURACAIXA - (dLSTART / 2)
      '
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
  Set objCC = Nothing
  For I = 1 To intQtdPontos
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, ".", "Left", True
  Next
End Sub
Public Sub TERMINA_IMPRESSAO()
  'Enviando para a impressora
  Printer.EndDoc
End Sub

Public Sub COMPROV_PENHOR(ByVal pLocacaoID As Long, ByVal pMotel As String, ByVal pQtdVias As Integer)
  'Imprimir corpo
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  Dim objGeral As busApler.clsGeral
  '
  Dim VRTotalPenhor As Currency
  Dim sMsg As String
  Dim dDataAtual  As Date
  Dim I As Integer
  '
  Set objGeral = New busApler.clsGeral
  strSql = "SELECT APARTAMENTO.NUMERO, LOCACAO.PKID, CONTACORRENTE.CLIENTE, CONTACORRENTE.DESCOBJETO, CONTACORRENTE.DOCUMENTOPENHOR, CONTACORRENTE.VALOR AS PGTOPENHOR, LOCACAO.SEQUENCIAL, CONTACORRENTE.DESCOBJETO From " & _
    "APARTAMENTO INNER JOIN LOCACAO ON APARTAMENTO.PKID = LOCACAO.APARTAMENTOID " & _
    "INNER JOIN CONTACORRENTE ON LOCACAO.PKID = CONTACORRENTE.LOCACAOID " & _
    "WHERE LOCACAO.PKID = " & Formata_Dados(pLocacaoID, tpDados_Longo) & _
    " AND CONTACORRENTE.STATUSCC = " & Formata_Dados("PH", tpDados_Texto)
    
  Set objRs = objGeral.ExecutarSQL(strSql)
  dDataAtual = Now
  If Not objRs.EOF Then
    Do While Not objRs.EOF
      'Calculo dos campos
      VRTotalPenhor = IIf(Not IsNumeric(objRs!PGTOPENHOR), 0, objRs!PGTOPENHOR)
      sMsg = "Eu " & objRs!CLIENTE & " Doc. Nº " & objRs!DOCUMENTOPENHOR
      sMsg = sMsg & " deixei por minha livre e expontânea  vontade " & objRs!DESCOBJETO
      sMsg = sMsg & " como garantia da divida contraida, no valor de R$ " & Format(VRTotalPenhor, "###,##0.00")
      sMsg = sMsg & ", pois me faltava este numerario para quitar minha despesa realizada neste estabelecimento na data "
      sMsg = sMsg & Format(dDataAtual, "DD/MM/YYYY") & " no apartamento "
      sMsg = sMsg & objRs!NUMERO & " Assim sendo comprometo-me a resgatar o bem no prazo de 15 dias, apos os quais autorizo a liquidacao deste objeto para saldar minha divida."
      For I = 1 To pQtdVias
        '--------
        QUEBRA_LINHA "TERMO DE CONFISSAO DE DIVIDA", lLARGURACAIXA - (dLSTART / 2)
        QUEBRA_LINHA pMotel, lLARGURACAIXA - (dLSTART / 2)
        PULAR_LINHA 1
        QUEBRA_LINHA objRs!NUMERO, lLARGURACAIXA - (dLSTART / 2)
        IMPRIMIR_LINHA 1
        QUEBRA_LINHA Format(dDataAtual, "DD/MM/YYYY hh:mm"), lLARGURACAIXA - (dLSTART / 2)
        PULAR_LINHA 2
        QUEBRA_LINHA sMsg, lLARGURACAIXA - (dLSTART / 2)
            
        PULAR_LINHA 3
        IMPRIMIR_LINHA 0
        QUEBRA_LINHA "Hospede", lLARGURACAIXA - (dLSTART / 2)
        PULAR_LINHA 3
        IMPRIMIR_LINHA 0
        QUEBRA_LINHA pMotel, lLARGURACAIXA - (dLSTART / 2)
        '
        PULAR_LINHA 1
        IMPRIMIR_LINHA 1
        QUEBRA_LINHA "Nº " & objRs!SEQUENCIAL, lLARGURACAIXA - (dLSTART / 2)
        '
        PULAR_LINHA 3
      Next
      objRs.MoveNext
    Loop
    For I = 1 To intQtdPontos
      IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, ".", "Left", True
    Next
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
End Sub

Public Sub IMP_COMPROV_PENHOR(pLocacaoID As Long, pMotel As String, pQtdVias As Integer)
  'Impressao do comprovante de Penhor
On Error GoTo ErrHandler
  '
  INICIA_IMPRESSAO
  '
  COMPROV_PENHOR pLocacaoID, pMotel, pQtdVias
  '
  TERMINA_IMPRESSAO
  '
  Exit Sub
ErrHandler:
  MsgBox "O seguinte erro ocorreu: " & Err.Description & ". Erro na impressao do comprovante de Entrada na portaria.", vbExclamation, TITULOSISTEMA
  TERMINA_IMPRESSAO
End Sub

Public Sub COMPROV_FATURA(ByVal pCCID As Long, ByVal pMotel As String, ByVal pQtdVias As Integer)
  'Imprimir corpo
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  Dim objGeral As busApler.clsGeral
  '
  Dim VRTotalFatura As Currency
  Dim sMsg As String
  Dim dDataAtual  As Date
  Dim I As Integer
  '
  Set objGeral = New busApler.clsGeral
  strSql = "SELECT EMPRESA.NOME, APARTAMENTO.NUMERO, CONTACORRENTE.VALOR AS PGTOFATURA, LOCACAO.SEQUENCIAL, CONTACORRENTE.DTHORACC, PARCELA.VRPARCELA, PARCELA.PARCELA, PARCELA.DTVENCIMENTO" & _
    " FROM APARTAMENTO INNER JOIN LOCACAO ON APARTAMENTO.PKID = LOCACAO.APARTAMENTOID " & _
    "INNER JOIN CONTACORRENTE ON LOCACAO.PKID = CONTACORRENTE.LOCACAOID " & _
    "INNER JOIN PARCELA ON CONTACORRENTE.PKID = PARCELA.CONTACORRENTEID " & _
    "LEFT JOIN VIAGEM ON LOCACAO.PKID = VIAGEM.LOCACAOID " & _
    "LEFT JOIN EMPRESA ON EMPRESA.PKID = VIAGEM.EMPRESAID " & _
    "WHERE CONTACORRENTE.PKID = " & Formata_Dados(pCCID, tpDados_Longo) & _
    " ORDER BY PARCELA "
    
  Set objRs = objGeral.ExecutarSQL(strSql)
  dDataAtual = Now
  If Not objRs.EOF Then
    For I = 1 To pQtdVias
      VRTotalFatura = IIf(Not IsNumeric(objRs!PGTOFATURA), 0, objRs!PGTOFATURA)
      'Impressão do cabeçalho
      QUEBRA_LINHA "COMPROVANTE DE FATURA", lLARGURACAIXA - (dLSTART / 2)
      QUEBRA_LINHA pMotel, lLARGURACAIXA - (dLSTART / 2)
      If objRs!NOME & "" <> "<PARTICULAR>" Then
        PULAR_LINHA 1
        QUEBRA_LINHA objRs!NOME, lLARGURACAIXA - (dLSTART / 2)
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

Public Sub IMP_COMPROV_FATURA(pCCID As Long, pMotel As String, pQtdVias As Integer)
  'Impressao do comprovante de Fatura
On Error GoTo ErrHandler
  '
  INICIA_IMPRESSAO
  '
  COMPROV_FATURA pCCID, pMotel, pQtdVias
  '
  TERMINA_IMPRESSAO
  '
  Exit Sub
ErrHandler:
  MsgBox "O seguinte erro ocorreu: " & Err.Description & ". Erro na impressao do comprovante de Fatura.", vbExclamation, TITULOSISTEMA
  TERMINA_IMPRESSAO
End Sub

Public Sub COMPROV_DEPOSITO(ByVal pLocacaoID As Long, ByVal pMotel As String, ByVal pQtdVias As Integer, ByVal strUnidade As String)
  'Imprimir corpo
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim objCC     As busApler.clsContaCorrente
  Dim dtaAtual  As Date
  Dim I         As Integer
  '
  Dim curTotalDeposito    As Currency
  '
  Set objCC = New busApler.clsContaCorrente
  '
  Set objRs = objCC.SelecionarPagamentos("DP", _
                                         pLocacaoID)
  
  dtaAtual = Now
    
  For I = 1 To pQtdVias
    '--------
    IMPRIMIR_LINHA 0
    QUEBRA_LINHA pMotel, lLARGURACAIXA - (dLSTART / 2)
    IMPRIMIR_LINHA 0
    QUEBRA_LINHA Format(dtaAtual, "DD/MM/YYYY hh:mm"), lLARGURACAIXA - (dLSTART / 2)
    PULAR_LINHA 1
    
    QUEBRA_LINHA "COMPROVANTE DE DEPOSITO", lLARGURACAIXA - (dLSTART / 2)
    PULAR_LINHA 1
    QUEBRA_LINHA "UNIDADE " & strUnidade, lLARGURACAIXA - (dLSTART / 2)
    
    PULAR_LINHA 2
    IMPRIMIR_LINHA 0
    '------------
    'CABECALHO
    IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, "Valor", "Right", False
    IMPRIMIR_POSICAO_CORRETA 6, lPOSX + dLSTART, Printer.CurrentY, "Tipo pgto.", "Left", True
    '
    IMPRIMIR_LINHA 0
    '---
    'CORPO
    curTotalDeposito = 0
    If Not objRs.EOF Then
      objRs.MoveFirst
      Do While Not objRs.EOF
        IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(objRs.Fields("VALOR").Value, "###,##0.00"), "Right", False
        IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, objRs.Fields("DESC_STATUSCC").Value, "Left", True
        curTotalDeposito = curTotalDeposito + objRs.Fields("VALOR").Value
        '---
        objRs.MoveNext
      Loop
      IMPRIMIR_LINHA 0
      IMPRIMIR_POSICAO_CORRETA 1.5, lLARGURACAIXA - 1.5, Printer.CurrentY, Format(curTotalDeposito, "###,##0.00"), "Right", False
      IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, "TOTAL", "Left", True
      
    End If
    '---
    'FIM DO CORPO
    '------------
    PULAR_LINHA 2
    'IMPRIMIR_LINHA 1
    '
  Next
  objRs.Close
  Set objRs = Nothing
  Set objCC = Nothing
  For I = 1 To intQtdPontos
    IMPRIMIR_POSICAO_CORRETA 7.5, lPOSX + dLSTART, Printer.CurrentY, ".", "Left", True
  Next
End Sub

Public Sub IMP_COMPROV_DEPOSITO(pLocacaoID As Long, pMotel As String, pQtdVias As Integer, pUnidade As String)
  'Impressao do comprovante de Penhor
On Error GoTo ErrHandler
  '
  INICIA_IMPRESSAO
  '
  COMPROV_DEPOSITO pLocacaoID, pMotel, pQtdVias, pUnidade
  '
  TERMINA_IMPRESSAO
  '
  Exit Sub
ErrHandler:
  MsgBox "O seguinte erro ocorreu: " & Err.Description & ". Erro na impressao do comprovante de Entrada na portaria.", vbExclamation, TITULOSISTEMA
  TERMINA_IMPRESSAO
End Sub

Public Sub IMP_COMPROV_DESPESA(pDespesaId As Long, pMotel As String, pQtdVias As Integer, pDescrTurno As String, ByVal blnTrabalhaComFuncAssoc As Boolean)
  'Impressao do comprovante de Vendas
On Error GoTo ErrHandler
  '
  INICIA_IMPRESSAO
  '
  COMPROV_DESPESA pDespesaId, pMotel, pQtdVias, pDescrTurno, blnTrabalhaComFuncAssoc
  '
  TERMINA_IMPRESSAO
  '
  Exit Sub
ErrHandler:
  MsgBox "O seguinte erro ocorreu: " & Err.Description & ". Erro na impressao do comprovante de Entrada na portaria.", vbExclamation, TITULOSISTEMA
  TERMINA_IMPRESSAO
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


