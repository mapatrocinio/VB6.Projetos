VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmUserPrestEspecCons 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta prestador / especialidade"
   ClientHeight    =   6360
   ClientLeft      =   2580
   ClientTop       =   3105
   ClientWidth     =   11385
   Icon            =   "userPrestEspecCons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6360
   ScaleWidth      =   11385
   Begin VB.CommandButton cmdNormal 
      Caption         =   "&Consultar"
      Height          =   255
      Index           =   1
      Left            =   10020
      TabIndex        =   1
      Top             =   270
      Width           =   1215
   End
   Begin VB.TextBox txtItem 
      Height          =   525
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   7905
   End
   Begin VB.PictureBox picBotoes 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   11385
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5460
      Width           =   11385
      Begin VB.PictureBox picAlinDir 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   912
         Left            =   0
         ScaleHeight     =   915
         ScaleWidth      =   11325
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   11325
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   10020
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdConfirmar 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   8820
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   $"userPrestEspecCons.frx":000C
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   1680
            TabIndex        =   8
            Top             =   60
            Width           =   5565
         End
      End
   End
   Begin TrueDBGrid60.TDBGrid grdGeral 
      Height          =   4335
      Left            =   0
      OleObjectBlob   =   "userPrestEspecCons.frx":0095
      TabIndex        =   2
      Top             =   960
      Width           =   11355
   End
   Begin VB.Label Label1 
      Caption         =   "Informe o prestador ou especialidade"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmUserPrestEspecCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strPrestadorEspecialidade As String
Dim blnFechar As Boolean

Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String
Dim blnPrimeiraVez        As Boolean

Private Sub cmdCancelar_Click()
  blnFechar = True
  strPrestadorEspecialidade = ""
  Unload Me
End Sub

Private Sub cmdConfirmar_Click()
  On Error GoTo trata
  If grdGeral.Columns(0).Value & "" = "" Then
    TratarErroPrevisto "Selecionar um atendimento", "cmdOK_Click"
    Pintar_Controle txtItem, tpCorContr_Erro
    SetarFoco txtItem
    Exit Sub
  End If
  
  frmUserGRCons.objUserGRInc.txtPrestador.Text = grdGeral.Columns(0).Value
  frmUserGRCons.objUserGRInc.txtEspecialidade.Text = grdGeral.Columns(1).Value
  frmUserGRCons.objUserGRInc.txtSala.Text = grdGeral.Columns(2).Value
  frmUserGRCons.objUserGRInc.txtPeriodo.Text = grdGeral.Columns(3).Value
  '
  blnFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub



Private Sub cmdNormal_Click(Index As Integer)
  Dim strWhere As String
  '
  If Not ValidaCampos Then
    Exit Sub
  End If
  COLUNASMATRIZ = grdGeral.Columns.Count
  LINHASMATRIZ = 0
  strWhere = txtItem.Text
  MontaMatriz strWhere
  grdGeral.Bookmark = Null
  grdGeral.ReBind
  grdGeral.ApproxCount = LINHASMATRIZ

  grdGeral.SetFocus

End Sub

Public Sub MontaMatriz(Optional strWhere As String)
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGR     As busSisMed.clsGR
  '
  AmpS
  On Error GoTo trata
  '
  Set objGR = New busSisMed.clsGR
  '
  If Len(Trim(strWhere)) = 0 Then
    Set objRs = objGR.CapturaPrestEspec(strPrestadorEspecialidade, _
                                        Retorna_DIADASEMANA_Descr(Date))
  Else
    Set objRs = objGR.CapturaPrestEspec(strWhere, _
                                        Retorna_DIADASEMANA_Descr(Date))
  End If
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
  objRs.Close
  Set objRs = Nothing
  Set objGR = Nothing
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Activate()
  If blnPrimeiraVez Then
    DoEvents
    '
    COLUNASMATRIZ = grdGeral.Columns.Count
    LINHASMATRIZ = 0

    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = LINHASMATRIZ
    blnPrimeiraVez = False
    SetarFoco grdGeral
  End If
End Sub

Private Sub grdGeral_Click()
  'cmdConfirmar_Click
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
  TratarErro Err.Number, Err.Description, "[frmUserPrestEspecCons.grdGeral_UnboundReadDataEx]"
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  AmpS
  Me.Height = 6840
  Me.Width = 11475
  blnPrimeiraVez = True
  blnFechar = False
  CenterForm Me
  'Me.Caption = Me.Caption
  LerFiguras Me, tpBmp_Vazio, cmdConfirmar, cmdCancelar
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

  

Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then
    Cancel = True
    Exit Sub
  End If
End Sub

Private Sub txtItem_GotFocus()
  Seleciona_Conteudo_Controle txtItem
End Sub

Private Sub txtItem_LostFocus()
  Pintar_Controle txtItem, tpCorContr_Normal
End Sub



Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCampos = False
  If Not Valida_String(txtItem, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o Prestador / especialidade" & vbCrLf
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserPrestEspecCons.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserPrestEspecCons.ValidaCampos]", _
            Err.Description
End Function


