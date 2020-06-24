VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmUserGRTotalPrestCons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelamento de GR´s"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   8850
      ScaleHeight     =   1155
      ScaleWidth      =   2985
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6000
      Width           =   3045
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   880
         Left            =   1500
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   150
         Width           =   1335
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "ENTER"
         Default         =   -1  'True
         Height          =   880
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   150
         Width           =   1335
      End
   End
   Begin VB.Frame fraUnidade 
      Caption         =   "GR´s"
      Height          =   5565
      Left            =   60
      TabIndex        =   3
      Top             =   330
      Width           =   11835
      Begin TrueDBGrid60.TDBGrid grdGeral 
         Height          =   5310
         Left            =   90
         OleObjectBlob   =   "userGRTotalPrestCons.frx":0000
         TabIndex        =   0
         Top             =   180
         Width           =   11580
      End
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Confirmação de Expiração"
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
      Index           =   9
      Left            =   4560
      TabIndex        =   17
      Top             =   6270
      Width           =   2265
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Liberada para atendimento"
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
      Index           =   11
      Left            =   2220
      TabIndex        =   16
      Top             =   6270
      Width           =   2295
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Não Atendida"
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
      Index           =   12
      Left            =   990
      TabIndex        =   15
      Top             =   6270
      Width           =   1185
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Atendida a posteriore"
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
      Index           =   13
      Left            =   6870
      TabIndex        =   14
      Top             =   6270
      Width           =   1785
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Atendida"
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
      Left            =   6030
      TabIndex        =   13
      Top             =   5970
      Width           =   675
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "Movimento após o fechamento"
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
      Index           =   7
      Left            =   3360
      TabIndex        =   11
      Top             =   5970
      Width           =   2625
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Não"
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
      Index           =   6
      Left            =   1800
      TabIndex        =   10
      Top             =   6690
      Width           =   525
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Sim"
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
      Index           =   5
      Left            =   1230
      TabIndex        =   9
      Top             =   6690
      Width           =   525
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      Caption         =   "Impressão:"
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
      Index           =   4
      Left            =   60
      TabIndex        =   8
      Top             =   6690
      Width           =   915
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      Caption         =   "Status :"
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
      Index           =   0
      Left            =   60
      TabIndex        =   7
      Top             =   5970
      Width           =   765
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Cancelada"
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
      Index           =   3
      Left            =   2280
      TabIndex        =   6
      Top             =   5970
      Width           =   1035
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
      Left            =   990
      TabIndex        =   5
      Top             =   5970
      Width           =   525
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Fechada"
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
      Left            =   1560
      TabIndex        =   4
      Top             =   5970
      Width           =   675
   End
End
Attribute VB_Name = "frmUserGRTotalPrestCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public blnRetorno               As Boolean
Public blnFechar                As Boolean
Public blnPrimeiraVez           As Boolean
Private COLUNASMATRIZ           As Long
Private LINHASMATRIZ            As Long
Private Matriz()                As String
Private Sub cmdFechar_Click()
  On Error GoTo trata
  blnFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

Private Sub cmdOk_Click()
  Dim strMsg As String
  Dim objGR As busSisMed.clsGR
  On Error GoTo trata
  'Cancelamento da GR
  If RetornaCodTurnoCorrente = 0 Then
    MsgBox "Não há turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If
  If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
    MsgBox "Selecione uma GR para excluí-la.", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If
  If Trim(grdGeral.Columns("Status").Value & "") <> "F" Then
    MsgBox "Apenas pode ser excluída uma GR fechada.", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If
  'If Trim(grdGeral.Columns("Status").Value & "") = "F" Then
    'Pedir senha superior para alterar uma GR já fechada
    '----------------------------
    '----------------------------
    'Pede Senha Superior (Diretor, Gerente ou Administrador
    gsNomeUsuLib = ""
    gsNivelUsuLib = ""
    If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
      'Só pede senha superior se quem estiver logado não for superior
      frmUserLoginSup.Show vbModal
      
      If Len(Trim(gsNomeUsuLib)) = 0 Then
        strMsg = "É necessário a confirmação com senha superior para cancelar uma GR."
        TratarErroPrevisto strMsg, "cmdConfirmar_Click"
        SetarFoco grdGeral
        Exit Sub
      Else
        'Capturou Nome do Usuário, continua com processo
      End If
    End If
    '--------------------------------
    '--------------------------------
  'End If
  'Confirmação
  If MsgBox("Confirma cancelamento da GR " & grdGeral.Columns("Seq.").Value & " de " & grdGeral.Columns("Prontuário").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
    SetarFoco grdGeral
    Exit Sub
  End If
  
  Set objGR = New busSisMed.clsGR
  objGR.AlterarStatusGR grdGeral.Columns("ID").Value, _
                        "C", _
                        "", _
                        RetornaCodTurnoCorrente
  Set objGR = Nothing
  IMP_COMP_CANC_GR grdGeral.Columns("ID").Value, gsNomeEmpresa, 1
  blnPrimeiraVez = True
  Form_Activate
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim strSql            As String
  Dim datDataTurno      As Date
  Dim datDataIniAtual   As Date
  Dim datDataFimAtual   As Date
  '
  blnFechar = False
  blnRetorno = False
  blnPrimeiraVez = True
  AmpS
  Me.Height = 7725
  Me.Width = 12090
  CenterForm Me
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdFechar
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
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
  TratarErro Err.Number, Err.Description, "[frmUserGRTotalPrestCons.grdGeral_UnboundReadDataEx]"
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

Public Sub MontaMatriz()
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
  Set objRs = objGR.CapturaGRTurnoCorrenteTODOS(RetornaCodTurnosPelaDataTODOS(Now))
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

