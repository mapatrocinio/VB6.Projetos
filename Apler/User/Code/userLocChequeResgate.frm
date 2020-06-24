VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmUserLocChequeResgate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo de resgate de cheque"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5505
      Left            =   9405
      ScaleHeight     =   5505
      ScaleWidth      =   1860
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2115
         Left            =   90
         ScaleHeight     =   2055
         ScaleWidth      =   1605
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   3300
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1020
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   135
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5295
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do cheque"
      TabPicture(0)   =   "userLocChequeResgate.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame12"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame12 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4785
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   9015
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   4545
            Index           =   4
            Left            =   120
            ScaleHeight     =   4545
            ScaleWidth      =   8775
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   120
            Width           =   8775
            Begin VB.TextBox txtLancado 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4500
               Locked          =   -1  'True
               TabIndex        =   2
               TabStop         =   0   'False
               Text            =   "txtLancado"
               Top             =   4200
               Width           =   1455
            End
            Begin VB.TextBox txtRestante 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6930
               Locked          =   -1  'True
               TabIndex        =   3
               TabStop         =   0   'False
               Text            =   "txtRestante"
               Top             =   4200
               Width           =   1455
            End
            Begin VB.TextBox txtTotalaPagar 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2010
               Locked          =   -1  'True
               TabIndex        =   1
               TabStop         =   0   'False
               Text            =   "txtTotalaPagar"
               Top             =   4200
               Width           =   1455
            End
            Begin TrueDBGrid60.TDBGrid grdGeral 
               Height          =   4050
               Left            =   60
               OleObjectBlob   =   "userLocChequeResgate.frx":001C
               TabIndex        =   10
               Top             =   120
               Width           =   8310
            End
            Begin VB.Label Label3 
               Caption         =   "Lançado"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3690
               TabIndex        =   13
               Top             =   4200
               Width           =   795
            End
            Begin VB.Label Label6 
               Caption         =   "Restante"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6090
               TabIndex        =   12
               Top             =   4200
               Width           =   795
            End
            Begin VB.Label Label38 
               Caption         =   "Total a pagar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   630
               TabIndex        =   11
               Top             =   4200
               Width           =   1245
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserLocChequeResgate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                 As tpStatus
Public blnRetorno             As Boolean
Public blnFechar              As Boolean
Public strNumeroAptoPrinc     As String
'Informa Qual Grupo irá chamar os Tabs 3 - Fechamento...
Public intGrupo               As Integer
Public strGrupo               As String
Private blnFatura             As Boolean
'
Public strStatusLanc          As String
'CC - Conta Corrente
'RC - Recebimento
'RE - Recebimento Empresa
'DP - Depósito
'
Public lngLOCDESPVDAEXTID     As Long
Public lngCCId                As Long
Public lngTurnoRecebeId       As Long

Private blnPrimeiraVez        As Boolean

Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String


Private Sub cmdCancelar_Click()
  On Error GoTo trata
'''  Dim vrValorJaPago       As Currency
'''  Dim vrTotLoc            As Currency
'''  Dim strSql              As String
'''
'''  Dim objRs               As ADODB.Recordset
'''  Dim objGeral            As busApler.clsGeral
'''
'''  Select Case strStatusLanc
'''  Case "DE", "VD", "EX"
'''    'Capturar valor já pago
'''    vrValorJaPago = 0
'''    Set objGeral = New busApler.clsGeral
'''    strSql = "SELECT SUM(VALOR) AS VALORJAPAGO, SUM(VRGORJETA) AS VRGORJETAJAPAGO, SUM(VRTROCO) AS VRTROCOJAPAGO " & _
'''      "FROM CONTACORRENTE " & _
'''      " WHERE STATUSLANCAMENTO = " & Formata_Dados(strStatusLanc, tpDados_Texto, tpNulo_Aceita)
'''    If strStatusLanc = "DE" Then
'''      strSql = strSql & " AND DESPESAID = " & Formata_Dados(lngLOCDESPVDAEXTID, tpDados_Longo, tpNulo_Aceita)
'''    ElseIf strStatusLanc = "VD" Then
'''      strSql = strSql & " AND VENDAID = " & Formata_Dados(lngLOCDESPVDAEXTID, tpDados_Longo, tpNulo_Aceita)
'''    ElseIf strStatusLanc = "EX" Then
'''      strSql = strSql & " AND EXTRAID = " & Formata_Dados(lngLOCDESPVDAEXTID, tpDados_Longo, tpNulo_Aceita)
'''    End If
'''
'''    Set objRs = objGeral.ExecutarSQL(strSql)
'''    If Not objRs.EOF Then
'''      If Not IsNull(objRs.Fields("VALORJAPAGO").Value) Then
'''        vrValorJaPago = objRs.Fields("VALORJAPAGO").Value
'''      End If
'''      If Not IsNull(objRs.Fields("VRGORJETAJAPAGO").Value) Then
'''        vrValorJaPago = vrValorJaPago - objRs.Fields("VRGORJETAJAPAGO").Value
'''      End If
'''      If Not IsNull(objRs.Fields("VRTROCOJAPAGO").Value) Then
'''        vrValorJaPago = vrValorJaPago - objRs.Fields("VRTROCOJAPAGO").Value
'''      End If
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
'''    Set objGeral = Nothing
'''    'Depende do Tipo
'''    vrTotLoc = CCur(IIf(Not IsNumeric(txtTotalaPagar.Text), 0, txtTotalaPagar.Text))
'''    If vrValorJaPago <> vrTotLoc Then
'''      'Valor do pagamento < que valor a pagar
'''      TratarErroPrevisto "Pagamento não pode ser diferente do restante. Favor complementá-la."
'''      SetarFoco cboTipoPagamento
'''    Else
'''      'Cancelar Cartão
'''      blnFechar = True
'''      Unload Me
'''    End If
'''  Case Else
    'Cancelar Cartão
    blnFechar = True
    Unload Me
'''  End Select
'''
'''
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub
Private Sub cmdOk_Click()
  Dim objCC               As busApler.clsContaCorrente
  '
  On Error GoTo trata
  If Not ValidaContaCorrente Then
    Exit Sub
  End If

  Set objCC = New busApler.clsContaCorrente

  If Status = tpStatus_Incluir Then
    'Inclusão de cheque pra despesa
    objCC.AssociaCCDespesa lngLOCDESPVDAEXTID & "", _
                           grdGeral.Columns("ID").Value & ""
  ElseIf Status = tpStatus_Alterar Then
    'Alteração
  End If
  '
  Set objCC = Nothing
  blnRetorno = True
  blnFechar = True
  'Está ok, se for recebimento,
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub
'Propósito: Validar o ContaCorrente
Public Function ValidaContaCorrente() As Boolean
  Dim strMsg        As String
  Dim strMsgAlerta  As String
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim vrPago        As Currency

  Dim vrTotLoc      As Currency
  Dim vrTotDescLoc  As Currency

  Dim vrValor       As Currency
  Dim vrValorJaPago As Currency
   
  Dim vrGorjeta     As Currency

  Dim DtAtualMenosNDias  As Date

  Dim strMsgAux     As String
  Dim objGeral      As busApler.clsGeral
  Dim blnSetarFocoControle As Boolean
  Dim strCredito    As String
  '
  On Error GoTo trata
  Set objGeral = New busApler.clsGeral
  blnSetarFocoControle = True
  'CHEQUE
  If grdGeral.Columns("ID").Value & "" = "" Then
    'Selecionar o pagamento em cheque
    SetarFoco grdGeral
    strMsg = strMsg & "Selecionar o cheque" & vbCrLf
    blnSetarFocoControle = False
  End If
  If Len(strMsg) = 0 Then
    'Capturar valor já pago
    vrValorJaPago = 0
    strSql = "SELECT SUM(VALOR) AS VALORJAPAGO, SUM(VRGORJETA) AS VRGORJETAJAPAGO, SUM(VRTROCO) AS VRTROCOJAPAGO " & _
      "FROM CONTACORRENTE "
    Select Case strStatusLanc
    Case "DE"
      strSql = strSql & "WHERE DESPESAID = " & Formata_Dados(lngLOCDESPVDAEXTID, tpDados_Longo, tpNulo_Aceita)
    End Select
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      If Not IsNull(objRs.Fields("VALORJAPAGO").Value) Then
        vrValorJaPago = objRs.Fields("VALORJAPAGO").Value
      End If
      If Not IsNull(objRs.Fields("VRGORJETAJAPAGO").Value) Then
        vrValorJaPago = vrValorJaPago - objRs.Fields("VRGORJETAJAPAGO").Value
      End If
      If Not IsNull(objRs.Fields("VRTROCOJAPAGO").Value) Then
        vrValorJaPago = vrValorJaPago - objRs.Fields("VRTROCOJAPAGO").Value
      End If
    End If
    objRs.Close
    Set objRs = Nothing
    'Validar Valor
    vrValor = CCur(IIf(Not IsNumeric(grdGeral.Columns("Valor").Value & ""), 0, grdGeral.Columns("Valor").Value & ""))
    vrGorjeta = 0
    'Calcula Valor Pago
    vrPago = vrValor + vrValorJaPago - vrGorjeta
    'Depende do Tipo
    Select Case strStatusLanc
    Case "DE"
      vrTotLoc = CCur(IIf(Not IsNumeric(txtTotalaPagar.Text), 0, txtTotalaPagar.Text))
      'vrTotDescLoc = CCur(IIf(Not IsNumeric(txtDesconto.Text), 0, txtDesconto.Text))
      'Validar Valor
      If vrPago < vrTotLoc Then
        'Valor do pagamento < que valor a pagar
        strMsgAux = "" & vbCrLf
        strMsgAux = "Valor pago menor que valor a pagar" & vbCrLf & vbCrLf & _
          "Caso confirme, terá que fazer um novo lançamento para complementar o recebimento. Deseja continuar ?"
        If MsgBox(strMsgAux, vbYesNo, TITULOSISTEMA) = vbNo Then
          strMsg = "Valor pago não pode ser menor que valor a pagar" & vbCrLf
        End If
      ElseIf vrPago > vrTotLoc Then
        strMsg = "Valor pago não pode ser maior que valor a pagar" & vbCrLf
      End If
    End Select
  End If
  Set objGeral = Nothing
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "ValidaContaCorrente"
    ValidaContaCorrente = False
  Else
    ValidaContaCorrente = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    blnPrimeiraVez = False
    SetarFoco grdGeral
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserLocChequeResgate.Form_Activate]"
End Sub


Public Sub MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busApler.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busApler.clsGeral
  '
  strSql = "SELECT CONTACORRENTE.STATUSLANCAMENTO, CONTACORRENTE.PKID, " & _
            " case STATUSLANCAMENTO when 'CC' then 'Conta Corrente' when 'RC' then 'Recebimento' when 'RE' then 'Recebimento Empresa' when 'DP' then 'Depósito' when 'DE' then 'Despesa' when 'VD' then 'Venda' when 'EX' then 'Extra' else '' end, CONTACORRENTE.DTHORACC, BANCO.NOME, CONTACORRENTE.AGENCIA, CONTACORRENTE.CONTA, CONTACORRENTE.VALOR  " & _
            "FROM CONTACORRENTE LEFT JOIN BANCO ON BANCO.PKID = CONTACORRENTE.BANCOID " & _
            " WHERE CONTACORRENTE.STATUSCC IN ('CH') " & _
            " AND TURNOCCID = " & RetornaCodTurnoCorrente & _
            " AND CONTACORRENTE.STATUSLANCAMENTO IN ('RE', 'CC', 'RC', 'DP', 'VD', 'EX') "
  strSql = strSql & " AND DESPESAID is null "
  strSql = strSql & " ORDER BY CONTACORRENTE.PKID DESC;"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim Matriz(0 To COLUNASMATRIZ - 1, 0 To LINHASMATRIZ - 1)
  Else
    ReDim Matriz(0 To COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To COLUNASMATRIZ - 1  'varre as colunas
          Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Load()
On Error GoTo trata
  '
  blnFechar = False
  blnRetorno = False
  blnFatura = False
  '
  AmpS
  Me.Height = 5985
  Me.Width = 11355
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  If Status = tpStatus_Incluir Or Status = tpStatus_Consultar Then
  ElseIf Status = tpStatus_Alterar Then
  End If
  'Carregar Grid
  COLUNASMATRIZ = grdGeral.Columns.Count
  LINHASMATRIZ = 0
  MontaMatriz
  grdGeral.Bookmark = Null
  grdGeral.ReBind
  grdGeral.ApproxCount = LINHASMATRIZ
  
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo trata
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub


Private Sub grdGeral_UnboundReadDataEx( _
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
               Offset + intI, LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, COLUNASMATRIZ, LINHASMATRIZ, Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserLocChequeResgate.grdGeral_UnboundReadDataEx]"
End Sub

