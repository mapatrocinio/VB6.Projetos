VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmOSLis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de OS"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11610
   Begin VB.Frame fraImpressao 
      Caption         =   "Impressão"
      Height          =   525
      Left            =   5340
      TabIndex        =   13
      Top             =   5790
      Width           =   2355
      Begin VB.Label Label72 
         Caption         =   "CTRL + A - Anodização"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   210
         Width           =   2205
      End
   End
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   6360
      Left            =   9750
      ScaleHeight     =   6360
      ScaleWidth      =   1860
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   5775
         Left            =   60
         ScaleHeight     =   5715
         ScaleWidth      =   1635
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   180
         Width           =   1695
         Begin VB.CommandButton cmdOSFinal 
            Caption         =   "&V"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   180
            Width           =   1335
         End
         Begin VB.CommandButton cmdAnodizacao 
            Caption         =   "&X"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton cmdItemOS 
            Caption         =   "&Y"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1980
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   3750
            Width           =   1335
         End
         Begin VB.CommandButton cmdInserir 
            Caption         =   "&Z"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   4650
            Width           =   1335
         End
      End
   End
   Begin TrueDBGrid60.TDBGrid grdGeral 
      Height          =   5730
      Left            =   0
      OleObjectBlob   =   "userOSLis.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   9645
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   7800
      Top             =   5790
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Anodização parcial"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   2190
      TabIndex        =   12
      Top             =   5820
      Width           =   1545
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Inicial"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   1620
      TabIndex        =   11
      Top             =   5820
      Width           =   525
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      Caption         =   "Status Anodização e OS Final :"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   30
      TabIndex        =   10
      Top             =   5760
      Width           =   1515
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Anodização total"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   3780
      TabIndex        =   9
      Top             =   5820
      Width           =   1485
   End
End
Attribute VB_Name = "frmOSLis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim COLUNASMATRIZ               As Long
Dim LINHASMATRIZ                As Long
Private Matriz()                As String
Public Status                   As tpStatus




Private Sub cmdAlterar_Click()
  On Error GoTo trata
  If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
    MsgBox "Selecione uma OS!", vbExclamation, TITULOSISTEMA
    Exit Sub
  End If
  frmOSInc.Status = tpStatus_Alterar
  frmOSInc.lngOSID = grdGeral.Columns("ID").Value
  frmOSInc.Show vbModal
  
  If frmOSInc.blnRetorno Then
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  End If
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub cmdAnodizacao_Click()
  On Error GoTo trata
  If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
    MsgBox "Selecione uma OS!", vbExclamation, TITULOSISTEMA
    Exit Sub
  End If
  'frmAnodizacaoInc.Status = tpStatus_Incluir
  frmAnodizacaoInc.strOSNumero = grdGeral.Columns("Número").Value
  frmAnodizacaoInc.lngOSID = grdGeral.Columns("ID").Value
  frmAnodizacaoInc.Show vbModal
  
  'If frmAnodizacaoInc.blnRetorno Then
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  'End If
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

'''Private Sub cmdExcluir_Click()
'''  Dim objOS     As busSisMetal.clsOS
'''  Dim objGer        As busSisMetal.clsGeral
'''  Dim objRs         As ADODB.Recordset
'''  Dim strSql        As String
'''  '
'''  On Error GoTo trata
'''  If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
'''    MsgBox "Selecione uma OS para exclusão.", vbExclamation, TITULOSISTEMA
'''    Exit Sub
'''  End If
'''  '
'''  Set objGer = New busSisMetal.clsGeral
'''  'ITEM_PEDIDO
'''  strSql = "Select * from ITEM_PEDIDO WHERE PEDIDOID = " & grdGeral.Columns("ID").Value
'''  Set objRs = objGer.ExecutarSQL(strSql)
'''  If Not objRs.EOF Then
'''    objRs.Close
'''    Set objRs = Nothing
'''    Set objGer = Nothing
'''    TratarErroPrevisto "OS não pode ser excluido pois já possui itens lançados.", "frmOSLis.cmdExcluir_Click"
'''    SetarFoco grdGeral
'''    Exit Sub
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
'''  '
'''  Set objGer = Nothing
'''  '
'''  '
'''  If MsgBox("Confirma exclusão do OS " & grdGeral.Columns("Ano-OS").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then Exit Sub
'''  'OK
'''  Set objOS = New busSisMetal.clsOS
'''
'''  objOS.ExcluirOS CLng(grdGeral.Columns("ID").Value)
'''  '
'''  MontaMatriz
'''  grdGeral.Bookmark = Null
'''  grdGeral.ReBind
'''
'''  Set objOS = Nothing
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdInserir_Click()
  On Error GoTo trata
  frmOSInc.Status = tpStatus_Incluir
  frmOSInc.lngOSID = 0
  frmOSInc.Show vbModal
  
  If frmOSInc.blnRetorno Then
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  End If
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub


Private Sub cmdItemOS_Click()
  On Error GoTo trata
  If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
    MsgBox "Selecione uma OS!", vbExclamation, TITULOSISTEMA
    Exit Sub
  End If
  frmOSInc.Status = tpStatus_Consultar
  frmOSInc.lngOSID = grdGeral.Columns("ID").Value
  frmOSInc.Show vbModal
  
  If frmOSInc.blnRetorno Then
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  End If
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub cmdOSFinal_Click()
  On Error GoTo trata
  If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
    MsgBox "Selecione uma OS!", vbExclamation, TITULOSISTEMA
    Exit Sub
  End If
  frmOSFinalLis.strOSNumero = grdGeral.Columns("Número").Value
  frmOSFinalLis.lngOSID = grdGeral.Columns("ID").Value
  frmOSFinalLis.Show vbModal
  
  'If frmOSFinalLis.blnRetorno Then
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  'End If
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  'Tratamento de tecla para verificação de chamada de Outras telas
  'verifica se tem permissão
  'Tudo ok, faz chamada
  Select Case KeyAscii
  Case 1
    'NOVO - IMPRIME PEDIDO EM TELA
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione uma OS para imprimí-la.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    Report1.Connect = ConnectRpt
    Report1.ReportFileName = gsReportPath & "OS.rpt"
    '
    'If optSai1.Value Then
      Report1.Destination = 0 'Video
    'ElseIf optSai2.Value Then
    '  Report2.Destination = 1   'Impressora
    'End If
    Report1.CopiesToPrinter = 1
    Report1.WindowState = crptMaximized
    '
    Report1.Formulas(0) = "OSID = " & Formata_Dados(grdGeral.Columns("ID").Value, tpDados_Longo)
    '
    Report1.Action = 1
    '
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmOSLis.Form_KeyPress]"
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  AmpS
  Me.Height = 6840
  Me.Width = 11700

  CenterForm Me

  LerFiguras Me, tpBmp_Vazio, , , cmdFechar, , , cmdInserir, cmdAlterar
  LerFigurasAvulsas cmdItemOS, "Cortesia.ico", "CortesiaDown.ico", "Itens do pedido"
  LerFigurasAvulsas cmdAnodizacao, "Anodizacao.ico", "AnodizacaoDown.ico", "Definir Anodizacao"
  LerFigurasAvulsas cmdOSFinal, "OSFinal.ico", "OSFinalDown.ico", "Controlar OS Final"
  'Captura o Dados da Unidade
  '
  COLUNASMATRIZ = grdGeral.Columns.Count
  LINHASMATRIZ = 0

  MontaMatriz
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Public Sub MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGeral    As busSisMetal.clsGeral
  '
  On Error GoTo trata

  Set objGeral = New busSisMetal.clsGeral
  '
  strSql = "SELECT OS.PKID, OS.NUMERO , OS.NF, OS.DATA, FORNECEDOR.NOME, FABRICA.NOME, OS.STATUS, OS.STATUS_FINAL " & _
        "FROM OS LEFT JOIN LOJA FORNECEDOR ON OS.FORNECEDORID = FORNECEDOR.PKID " & _
        " LEFT JOIN LOJA FABRICA ON OS.FABRICAID = FABRICA.PKID " & _
        " ORDER BY OS.PKID DESC;"
  '
  Set objRs = objGeral.ExecutarSQL(strSql)
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
  Set objGeral = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
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
  TratarErro Err.Number, Err.Description, "[frmOSLis.grdGeral_UnboundReadDataEx]"
End Sub

