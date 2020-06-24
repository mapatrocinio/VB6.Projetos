VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmItemOSFinalInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de anodização"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   6975
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Controle de anodização"
      TabPicture(0)   =   "userItemOSFinalInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(17)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "grdOS"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin TrueDBGrid60.TDBGrid grdOS 
         Height          =   4890
         Left            =   90
         OleObjectBlob   =   "userItemOSFinalInc.frx":001C
         TabIndex        =   0
         Top             =   780
         Width           =   9210
      End
      Begin VB.Label Label1 
         Caption         =   $"userItemOSFinalInc.frx":6C33
         ForeColor       =   &H000000FF&
         Height          =   435
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   6360
         Width           =   8715
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "* Tecle ENTER para mudar de coluna, ao final o sistema validará os dados na base salvando as informações"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   6090
         Width           =   8715
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "* Clique no botão SALVAR para registrar todas as definições de anodização para os perfis ou tecle ESC para sair"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   5820
         Width           =   8715
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Anodização"
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   6
         Top             =   420
         Width           =   2655
      End
   End
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7245
      Left            =   9660
      ScaleHeight     =   7245
      ScaleWidth      =   1860
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2115
         Left            =   120
         ScaleHeight     =   2055
         ScaleWidth      =   1545
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   4980
         Width           =   1605
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Height          =   885
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   150
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   885
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1020
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmItemOSFinalInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public lngOSFINALID             As Long
Public lngANODIZACAOITEMID      As Long
Public lngITEMOSFINALID         As Long
Public lngOSID                  As Long
Public lngCORID                 As Long

Public strOSNumero              As String
Public strCor                   As String

Dim blnFechar                   As Boolean
Public blnRetorno               As Boolean
Public blnPrimeiraVez           As Boolean
'
Dim ANOD_COLUNASMATRIZ        As Long
Dim ANOD_LINHASMATRIZ         As Long
Private ANOD_Matriz()         As String


Public Sub ANOD_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGeral  As busSisMetal.clsGeral
  '
  On Error GoTo trata

  Set objGeral = New busSisMetal.clsGeral
  '
  'strSql = "SELECT VW_CONS_ITEM_OS.ITEM_OSID, VW_CONS_ITEM_OS.LINHAID,'N',VW_CONS_ITEM_OS.NOME, VW_CONS_ITEM_OS.CODIGO, VW_CONS_ITEM_OS.QUANTIDADE, VW_CONS_ITEM_OS.ANOD_BRA_QUANTIDADE, VW_CONS_ITEM_OS.ANOD_BRI_QUANTIDADE, VW_CONS_ITEM_OS.ANOD_BRO_QUANTIDADE, VW_CONS_ITEM_OS.ANOD_NAT_QUANTIDADE " & _
           "FROM VW_CONS_ITEM_OS " & _
           " WHERE VW_CONS_ITEM_OS.OSID = " & Formata_Dados(lngANODIZACAOITEMID, tpDados_Longo) & _
           " ORDER BY VW_CONS_ITEM_OS.NOME, VW_CONS_ITEM_OS.CODIGO "
  '
  '
  strSql = "SELECT " & _
            " ANODIZACAO_ITEM.PKID, ITEM_OS_FINAL.PKID, LINHA.PKID, 'N', TIPO_LINHA.NOME, LINHA.CODIGO, " & _
            " ANODIZACAO_ITEM.QUANTIDADE, " & _
            " ISNULL(BAIXA_ITEM_OS.QUANTIDADE_BAIXA,0) - ISNULL(ITEM_OS_FINAL.QUANTIDADE,0), " & _
            " ITEM_OS_FINAL.QUANTIDADE " & _
            " From ITEM_OS " & _
            " INNER JOIN ANODIZACAO_ITEM ON ITEM_OS.PKID = ANODIZACAO_ITEM.ITEM_OSID " & _
            "   AND ANODIZACAO_ITEM.CORID = " & Formata_Dados(lngCORID, tpDados_Longo) & _
            " LEFT JOIN ITEM_OS_FINAL ON ANODIZACAO_ITEM.PKID = ITEM_OS_FINAL.ANODIZACAO_ITEMID " & _
            "     AND ITEM_OS_FINAL.OS_FINALID = " & Formata_Dados(lngOSFINALID, tpDados_Longo) & _
            " LEFT JOIN " & _
            "     (SELECT " & _
            "     OS.ANODIZACAO_ITEMID, " & _
            "     SUM(OS.QUANTIDADE) AS QUANTIDADE_BAIXA " & _
            "     FROM ITEM_OS_FINAL OS " & _
            "     GROUP BY OS.ANODIZACAO_ITEMID) " & _
            "     AS BAIXA_ITEM_OS ON BAIXA_ITEM_OS.ANODIZACAO_ITEMID = ANODIZACAO_ITEM.PKID " & _
            " LEFT JOIN LINHA ON LINHA.PKID = ITEM_OS.LINHAID " & _
            " LEFT JOIN TIPO_LINHA ON TIPO_LINHA.PKID = LINHA.TIPO_LINHAID " & _
            " WHERE ITEM_OS.OSID = " & Formata_Dados(lngOSID, tpDados_Longo) & _
            " ORDER BY ITEM_OS.PKID"

  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    ANOD_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim ANOD_Matriz(0 To ANOD_COLUNASMATRIZ - 1, 0 To ANOD_LINHASMATRIZ - 1)
  Else
    ReDim ANOD_Matriz(0 To ANOD_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To ANOD_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To ANOD_COLUNASMATRIZ - 1  'varre as colunas
          If intJ = ANOD_COLUNASMATRIZ - 1 Then
            ANOD_Matriz(intJ, intI) = intI & ""
          Else
            ANOD_Matriz(intJ, intI) = objRs(intJ) & ""
          End If
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set objGeral = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdOK_Click()
  Dim objItemOSFinal As busSisMetal.clsItemOSFinal
  On Error GoTo trata
  Dim intI      As Integer
  '
  Select Case tabDetalhes.Tab
  Case 0 'Gravar Anodização
    If ValidaCamposAnodOrigemAll Then
      SetarFoco grdOS
      grdOS.Col = 8
      grdOS.Row = 0
      Exit Sub
    End If
    'OK procede com o cadastro
    '
    Set objItemOSFinal = New busSisMetal.clsItemOSFinal
    For intI = 0 To ANOD_LINHASMATRIZ - 1
      grdOS.Bookmark = CLng(intI)
      'If grdOS.Columns("Branco").Text & "" <> "" Or _
        grdOS.Columns("Brilho").Text & "" <> "" Or _
        grdOS.Columns("Bronze").Text & "" <> "" Or _
        grdOS.Columns("Natural").Text & "" <> "" Then
      If grdOS.Columns("*").Text & "" = "<Bitmap>.S" Then
        'Propósito: Cadastrar anodização
        '
        objItemOSFinal.InserirItemOSFinalItem lngOSFINALID, _
                                              lngOSID, _
                                              lngCORID, _
                                              grdOS.Columns("ANODIZACAOITEMID").Text & "", _
                                              IIf(grdOS.Columns("ITEMOSFINALID").Text & "" = "", 0, grdOS.Columns("ITEMOSFINALID").Text & ""), _
                                              grdOS.Columns("LINHAID").Text & "", _
                                              grdOS.Columns("Quantidade").Text & ""
        blnRetorno = True
      End If
    Next
    Set objItemOSFinal = Nothing
    '
    blnFechar = True
    Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Montar RecordSet
    ANOD_COLUNASMATRIZ = grdOS.Columns.Count
    ANOD_LINHASMATRIZ = 0
    ANOD_MontaMatriz
    grdOS.Bookmark = Null
    grdOS.ReBind
    grdOS.ApproxCount = ANOD_LINHASMATRIZ
    '
    SetarFoco grdOS
    grdOS.Col = 8
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmItemOSFinalInc.Form_Activate]"
End Sub


Private Sub grdOS_BeforeUpdate(Cancel As Integer)
  On Error GoTo trata
  'Atualiza Matriz
  ANOD_Matriz(3, grdOS.Row) = "S"
  ANOD_Matriz(8, grdOS.Row) = grdOS.Columns(8).Text
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmItemOSFinalInc.grdOS_BeforeRowColChange]"
End Sub

Private Sub grdOS_UnboundReadDataEx( _
     ByVal RowBuf As TrueDBGrid60.RowBuffer, _
    StartLocation As Variant, ByVal Offset As Long, _
    ApproximatePosition As Long)
  ' UnboundReadData is fired by an unbound grid whenever
  ' it requires data for display. This event will fire
  ' when the grid is first shown, when Refresh or ReBind
  ' is used, when the grid is scrolled, and after a
  ' record in the grid is modified and the user commits
  ' the change by moving off of the current row. The
  ' grid fetches data in "chunks", and the number of rows
  ' the grid is asking for is given by RowBuf.RowCount.
  ' RowBuf is the row buffer where you place the data
  ' the bookmarks for the rows that the grid is
  ' requesting to display. It will also hold the number
  ' of rows that were successfully supplied to the grid.
  ' StartLocation is a vrtBookmark which, together with
  ' Offset, specifies the row for the programmer to start
  ' transferring data. A StartLocation of Null indicates
  ' a request for data from BOF or EOF.
  ' Offset specifies the relative position (from
  ' StartLocation) of the row for the programmer to start
  ' transferring data. A positive number indicates a
  ' forward relative position while a negative number
  ' indicates a backward relative position. Regardless
  ' of whether the rows to be read are before or after
  ' StartLocation, rows are always fetched going forward
  ' (this is why there is no ReadPriorRows parameter to
  ' the procedure).
  ' If you page down on the grid, for instance, the new
  ' top row of the grid will have an index greater than
  ' the StartLocation (Offset > 0). If you page up on
  ' the grid, the new index is less than that of
  ' StartLocation, so Offset < 0. If StartLocation is
  ' a vrtBookmark to row N, the grid always asks for row
  ' data in the following order:
  '   (N + Offset), (N + Offset + 1), (N + Offset + 2)...
  ' ApproximatePosition is a value you can set to indicate
  ' the ordinal position of (StartLocation + Offset).
  ' Setting this variable will enhance the ability of the
  ' grid to display its vertical scroll bar accurately.
  ' If the exact ordinal position of the new location is
  ' not known, you can set it to a reasonable,
  ' approximate value, or just ignore this parameter.

  On Error GoTo trata
  '
  Dim intColIndex      As Integer
  Dim intJ             As Integer
  Dim intRowsFetched   As Integer
  Dim intI             As Long
  Dim lngNewPosition   As Long
  Dim vrtBookmark      As Variant
  '
  intRowsFetched = 0
  For intI = 0 To RowBuf.RowCount - 1
    ' Get the vrtBookmark of the next available row
    vrtBookmark = GetRelativeBookmarkGeral(StartLocation, _
               Offset + intI, ANOD_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, ANOD_COLUNASMATRIZ, ANOD_LINHASMATRIZ, ANOD_Matriz)
    Next intJ

    ' Set the vrtBookmark for the row
    RowBuf.Bookmark(intI) = vrtBookmark

    ' Increment the count of fetched rows
    intRowsFetched = intRowsFetched + 1
  Next intI

' Tell the grid how many rows were fetched
  RowBuf.RowCount = intRowsFetched

' Set the approximate scroll bar position. Only
' nonnegative values of IndexFromBookmark() are valid.
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ANOD_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmAnodizadoraInc.grdOS_UnboundReadDataEx]"
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs As ADODB.Recordset
  Dim strSql As String
  '
  blnFechar = False 'Não Pode Fechar pelo X
  blnRetorno = False
  blnPrimeiraVez = True
  '
  AmpS
  Me.Height = 7620
  Me.Width = 11610
  CenterForm Me
  Me.Caption = Me.Caption & " - " & strOSNumero
  grdOS.Caption = "Anodização [" & UCase(strCor) & "]"
  
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdFechar
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Function ValidaCamposAnodOrigemLinha(intLinha As Integer) As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  Dim strSql                As String
  Dim objRs                 As ADODB.Recordset
  Dim objGeral              As busSisMetal.clsGeral
  '
  Dim lngTotal              As Long
  Dim lngTotalANOD          As Long
  '
  blnSetarFocoControle = True
  '
  strMsg = ""
  'Validção da anodização
  If Not Valida_Moeda(grdOS.Columns("Quantidade"), TpnaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
    strMsg = strMsg & "Quantidade de anodização inválida na linha " & intLinha + 1 & vbCrLf
  End If
  '
  If Len(strMsg) = 0 Then
    'Validações dos totais
    lngTotal = 0
    lngTotalANOD = 0
    '
    lngTotal = CLng(grdOS.Columns("Qtd. Total")) - CLng(grdOS.Columns("Qtd. Baixa"))
    lngTotalANOD = CLng(IIf(Not IsNumeric(grdOS.Columns("Quantidade")), 0, grdOS.Columns("Quantidade")))
    '
    If lngTotalANOD > (lngTotal) Then
      strMsg = strMsg & "O total lançado dos perfis deve ser igual ou menor a quantidade na linha " & intLinha + 1 & vbCrLf
    End If
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmItemOSFinalInc.ValidaCamposAnodOrigemLinha]"
    ValidaCamposAnodOrigemLinha = False
  Else
    ValidaCamposAnodOrigemLinha = True
  End If
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserLeituraFechaInc.ValidaCamposAnodOrigemLinha]"
  ValidaCamposAnodOrigemLinha = False
End Function

Private Function ValidaCamposAnodOrigemAll() As Boolean
  On Error GoTo trata
  Dim blnRetorno            As Boolean
  Dim blnCadastrou1Linha    As Boolean
  Dim blnEncontrouErro      As Boolean
  Dim blnEncontrouErroLinha As Boolean
  Dim intRows               As Integer
  'Validar todas as linhas da matriz
  blnEncontrouErro = False
  blnCadastrou1Linha = False
  blnEncontrouErroLinha = False
  blnRetorno = True
  
  
  For intRows = 0 To ANOD_LINHASMATRIZ - 1
    grdOS.Bookmark = CLng(intRows)
    '
    If grdOS.Columns("*").Text & "" = "<Bitmap>.S" Then
      'Somente válida se preencheu algo, sneão considera ok
      If grdOS.Columns("Quantidade").Text & "" <> "" Then
        If Not ValidaCamposAnodOrigemLinha(grdOS.Row) Then
          blnEncontrouErro = True
          blnEncontrouErroLinha = True
        Else
          blnCadastrou1Linha = True
        End If
      Else
        'tudo brnao, considera OK
        blnCadastrou1Linha = True
      End If
    Else
      'blnEncontrouErro = True
    End If
    If blnEncontrouErro = True Then Exit For
  Next
  '
  If blnEncontrouErro = False And blnCadastrou1Linha = True Then
    blnRetorno = False
  End If
  If blnEncontrouErroLinha = False And blnEncontrouErro = False And blnCadastrou1Linha = False Then
    TratarErroPrevisto "Entre com ao menos 1 item para cadastro", "[frmItemOSFinalInc.ValidaCamposAnodOrigemAll]"
  End If
  grdOS.ReBind
  grdOS.SetFocus
  ValidaCamposAnodOrigemAll = blnRetorno
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserLeituraFechaInc.ValidaCamposAnodOrigemAll]"
  ValidaCamposAnodOrigemAll = False
End Function
Private Function ValidaCamposAnodOrigem(intLinha As Integer, intColuna As Integer) As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  '
  Select Case intColuna
  Case 5
    'Validção da quantidade branco
    If Not Valida_Moeda(grdOS.Columns(5), TpnaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
      strMsg = strMsg & "Quantidade de perfil branco inválida" & vbCrLf
    End If
  Case 6
    'Validção da quantidade brilho
    If Not Valida_Moeda(grdOS.Columns(6), TpnaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
      strMsg = strMsg & "Quantidade de perfil brilho inválida" & vbCrLf
    End If
  Case 7
    'Validção da quantidade bronze
    If Not Valida_Moeda(grdOS.Columns(7), TpnaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
      strMsg = strMsg & "Quantidade de perfil bronze inválida" & vbCrLf
    End If
  Case 8
    'Validção da quantidade natural
    If Not Valida_Moeda(grdOS.Columns(8), TpnaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
      strMsg = strMsg & "Quantidade de perfil natural inválida" & vbCrLf
    End If
  End Select
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmItemOSFinalInc.ValidaCamposAnodOrigem]"
    ValidaCamposAnodOrigem = False
  Else
    ValidaCamposAnodOrigem = True
  End If
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserGrupoCln.ValidaCamposAnodOrigem]"
  ValidaCamposAnodOrigem = False
End Function
Private Sub cmdFechar_Click()
  blnFechar = True
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then Cancel = True
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  '
  If Me.ActiveControl.Name <> "grdOS" Then
    If KeyAscii = 13 Then
      SendKeys "{tab}"
    End If
  Else
    
      
    If KeyAscii = 13 And grdOS.Row <> -1 Then
      If grdOS.Col = 8 Then
        If grdOS.Columns("ROWNUM").Value + 1 = ANOD_LINHASMATRIZ Then
          cmdOK_Click
        Else
          grdOS.Col = 8
          grdOS.MoveNext
        End If
      Else
        grdOS.Col = grdOS.Col + 1
      End If
    ElseIf (KeyAscii = 8) Then
    ElseIf (KeyAscii < 48 Or KeyAscii > 57) Then
      KeyAscii = 0
    End If
  End If
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmItemOSFinalInc.Form_KeyPress]"
End Sub


