VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserProntuarioCons 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta prontuario"
   ClientHeight    =   6360
   ClientLeft      =   2580
   ClientTop       =   3105
   ClientWidth     =   9465
   Icon            =   "userProntuarioCons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6360
   ScaleWidth      =   9465
   Begin VB.CommandButton cmdNormal 
      Caption         =   "&Inserir"
      Height          =   255
      Index           =   1
      Left            =   8130
      TabIndex        =   13
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton cmdNormal 
      Caption         =   "&Consultar"
      Height          =   255
      Index           =   0
      Left            =   8130
      TabIndex        =   3
      Top             =   420
      Width           =   1215
   End
   Begin VB.TextBox txtNome 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   90
      Width           =   7545
   End
   Begin VB.PictureBox picBotoes 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   9465
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5460
      Width           =   9465
      Begin VB.PictureBox picAlinDir 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   912
         Left            =   0
         ScaleHeight     =   915
         ScaleWidth      =   9345
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   9345
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   8070
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdConfirmar 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   6870
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   $"userProntuarioCons.frx":000C
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   630
            TabIndex        =   10
            Top             =   60
            Width           =   5565
         End
      End
   End
   Begin TrueDBGrid60.TDBGrid grdGeral 
      Height          =   4335
      Left            =   0
      OleObjectBlob   =   "userProntuarioCons.frx":0095
      TabIndex        =   4
      Top             =   960
      Width           =   9345
   End
   Begin MSMask.MaskEdBox mskDtNascimento 
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   390
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskCpf 
      Height          =   255
      Left            =   4710
      TabIndex        =   2
      Top             =   390
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   450
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   14
      Mask            =   "###.###.###-##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label7 
      Caption         =   "Dt. Nascimento"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   390
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "CPF"
      Height          =   195
      Index           =   4
      Left            =   3480
      TabIndex        =   11
      Top             =   420
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nome"
      Height          =   225
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmUserProntuarioCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strNome As String
Public strCPF As String
Public strDtNascimento As String

Dim blnFechar As Boolean

Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String
Dim blnPrimeiraVez        As Boolean

Private Sub cmdCancelar_Click()
  blnFechar = True
  strNome = ""
  strCPF = ""
  strDtNascimento = ""
  Unload Me
End Sub

Private Sub cmdConfirmar_Click()
  On Error GoTo trata
  If grdGeral.Columns(0).Value & "" = "" Then
    TratarErroPrevisto "Selecionar um prestador", "cmdOK_Click"
    Pintar_Controle txtNome, tpCorContr_Erro
    SetarFoco txtNome
    Exit Sub
  End If
  frmUserGRCons.objUserGRInc.txtProntuarioFim.Text = grdGeral.Columns(0).Value
  INCLUIR_VALOR_NO_MASK frmUserGRCons.objUserGRInc.mskDataNascFim, grdGeral.Columns(2).Value, TpMaskData
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
  Dim objUserProntuarioInc As SisMed.frmUserProntuarioInc

  '
  If Index = 0 Then
    If Not ValidaCampos Then
      Exit Sub
    End If
    COLUNASMATRIZ = grdGeral.Columns.Count
    LINHASMATRIZ = 0
    MontaMatriz txtNome.Text, _
                IIf(mskDtNascimento.ClipText = "", "", mskDtNascimento.Text), _
                mskCpf.ClipText
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = LINHASMATRIZ
  
    grdGeral.SetFocus
  ElseIf Index = 1 Then
    Set objUserProntuarioInc = New SisMed.frmUserProntuarioInc
    objUserProntuarioInc.Status = tpStatus_Incluir
    objUserProntuarioInc.IcProntuario = tpIcProntuario.tpIcProntuario_Pac
    objUserProntuarioInc.lngPKID = 0
    objUserProntuarioInc.intQuemChamou = 1
    objUserProntuarioInc.strNomeInicial = txtNome.Text
    objUserProntuarioInc.Show vbModal
    
    If objUserProntuarioInc.blnRetorno Then
      '
      Set objUserProntuarioInc = Nothing
      blnFechar = True
      Unload Me
      Exit Sub
    End If
    Set objUserProntuarioInc = Nothing
  End If
End Sub

Public Sub MontaMatriz(Optional strNomePar As String, _
                       Optional strDtNascimentoPar As String, _
                       Optional strCPFPar As String)
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
  If Len(Trim(strNomePar)) = 0 And Len(Trim(strDtNascimentoPar)) = 0 And Len(Trim(strCPFPar)) = 0 Then
    Set objRs = objGR.CapturaProntuario(strNome, _
                                        strCPF, _
                                        strDtNascimento)
  Else
    Set objRs = objGR.CapturaProntuario(strNomePar, _
                                        strCPFPar, _
                                        strDtNascimentoPar)
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
    If LINHASMATRIZ = 0 Then
      SetarFoco txtNome
    Else
      SetarFoco grdGeral
    End If
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
  Me.Width = 9555
  blnPrimeiraVez = True
  blnFechar = False
  CenterForm Me
  txtNome.Text = strNome
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


Private Sub mskCpf_GotFocus()
  Seleciona_Conteudo_Controle mskCpf
End Sub

Private Sub mskCpf_LostFocus()
  Pintar_Controle mskCpf, tpCorContr_Normal
End Sub

Private Sub mskDtNascimento_GotFocus()
  Seleciona_Conteudo_Controle mskDtNascimento
End Sub

Private Sub mskDtNascimento_LostFocus()
  Pintar_Controle mskDtNascimento, tpCorContr_Normal
End Sub

Private Sub txtNome_GotFocus()
  Seleciona_Conteudo_Controle txtNome
End Sub

Private Sub txtNome_LostFocus()
  Pintar_Controle txtNome, tpCorContr_Normal
End Sub


Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCampos = False
'''  If Not Valida_String(txtNome, TpObrigatorio, blnSetarFocoControle) Then
'''    strMsg = strMsg & "Preencher o nome" & vbCrLf
'''  End If
  
  If Len(Trim(mskCpf.ClipText)) > 0 Then
    If Not TestaCPF(mskCpf.ClipText) Then
      strMsg = strMsg & "Informar o CPF válido" & vbCrLf
      Pintar_Controle mskCpf, tpCorContr_Erro
      SetarFoco mskCpf
      blnSetarFocoControle = False
    End If
  End If
  If Not Valida_Data(mskDtNascimento, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de nascimento válida" & vbCrLf
  End If
  If txtNome.Text & "" = "" And Len(Trim(mskCpf.ClipText)) = 0 And Len(Trim(mskDtNascimento.ClipText)) = 0 Then
    strMsg = strMsg & "Informar o Nome ou CPF ou Data de nascimento para realizar a pesquisa" & vbCrLf
    Pintar_Controle txtNome, tpCorContr_Erro
    SetarFoco txtNome
    blnSetarFocoControle = False
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserProntuarioCons.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserProntuarioCons.ValidaCampos]", _
            Err.Description
End Function

