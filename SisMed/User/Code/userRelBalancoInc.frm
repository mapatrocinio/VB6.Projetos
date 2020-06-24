VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmUserRelBalancoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupo de Despesa"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5070
      Left            =   8250
      ScaleHeight     =   5070
      ScaleWidth      =   1860
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4725
         Left            =   120
         ScaleHeight     =   4665
         ScaleWidth      =   1605
         TabIndex        =   11
         Top             =   150
         Width           =   1665
         Begin VB.CommandButton cmdInserir 
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   3600
            Width           =   1335
         End
         Begin VB.CommandButton cmdRelatorio 
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   2730
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4785
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   8440
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Impressão"
      TabPicture(0)   =   "userRelBalancoInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Saldos"
      TabPicture(1)   =   "userRelBalancoInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdGeral"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   90
         TabIndex        =   14
         Top             =   1080
         Width           =   7665
         Begin MSMask.MaskEdBox mskData 
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   2
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskData 
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   3
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin Crystal.CrystalReport Report1 
            Left            =   300
            Top             =   1350
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            PrintFileLinesPerPage=   60
         End
         Begin VB.Label lblCliente 
            Caption         =   "Período : "
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Tag             =   "lblIdCliente"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Até"
            Height          =   255
            Left            =   2640
            TabIndex        =   15
            Tag             =   "lblIdCliente"
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   90
         TabIndex        =   13
         Top             =   360
         Width           =   7665
         Begin VB.OptionButton optSai2 
            Caption         =   "Impressora"
            Height          =   255
            Left            =   1440
            TabIndex        =   1
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optSai1 
            Caption         =   "Vídeo"
            Height          =   255
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin TrueDBGrid60.TDBGrid grdGeral 
         Height          =   3300
         Left            =   -74880
         OleObjectBlob   =   "userRelBalancoInc.frx":0038
         TabIndex        =   4
         Top             =   480
         Width           =   7545
      End
   End
End
Attribute VB_Name = "frmUserRelBalancoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                         As tpStatus
Public lngRELBALANCOID                As Long
Public bRetorno                       As Boolean
Public blnPrimeiraVez                 As Boolean
Public bFechar                        As Boolean
Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String



Private Sub cmdAlterar_Click()
  
  If Not IsNumeric(grdGeral.Columns("ID").Value) Then
    MsgBox "Selecione um sub grupo de despesa !", vbExclamation, TITULOSISTEMA
    Exit Sub
  End If
  frmUserRelBalancoCad.Status = tpStatus_Alterar
  frmUserRelBalancoCad.lngRELBALANCOID = lngRELBALANCOID
  frmUserRelBalancoCad.lngRELBALANCOID = CLng(grdGeral.Columns("ID").Value)
  frmUserRelBalancoCad.Show vbModal
  
  If frmUserRelBalancoCad.blnRetorno Then
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  End If

End Sub

Private Sub cmdCancelar_Click()
  bFechar = True
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdExcluir_Click()
  Dim objRelBalanco         As busSisMed.clsRelBalanco
  Dim objGeral              As busSisMed.clsGeral
  Dim objRs                 As ADODB.Recordset
  Dim strSql                As String
  '
  On Error GoTo trata
  If Len(Trim(grdGeral.Columns("ID").Value)) = 0 Then
    MsgBox "Selecione um saldo para excluir.", vbExclamation, TITULOSISTEMA
    Exit Sub
  End If
  If MsgBox("Confirma exclusão do saldo de " & grdGeral.Columns("Data").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then Exit Sub
  'OK
  Set objRelBalanco = New busSisMed.clsRelBalanco
  
  objRelBalanco.ExcluirRelBalanco CLng(grdGeral.Columns("ID").Value)
  '
  MontaMatriz
  grdGeral.Bookmark = Null
  grdGeral.ReBind
  
  Set objRelBalanco = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdInserir_Click()
  frmUserRelBalancoCad.Status = tpStatus_Incluir
  frmUserRelBalancoCad.lngRELBALANCOID = lngRELBALANCOID
  frmUserRelBalancoCad.Show vbModal
  
  If frmUserRelBalancoCad.blnRetorno Then
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  End If

End Sub


Private Sub cmdRelatorio_Click()
  On Error GoTo TratErro
  Dim objRelBalanco         As busSisMed.clsRelBalanco
  Dim curSaldoAnterior      As Currency
  Dim datDataSaldoAnterior  As Date
  Dim curReceita            As Currency
  Dim curPrestador          As Currency
  Dim curDespesa            As Currency
  AmpS
  
  If Not IsDate(mskData(0).Text) Then
    AmpN
    MsgBox "Data Inicial Inválida !", vbOKOnly, TITULOSISTEMA
    SetarFoco mskData(0)
    Pintar_Controle mskData(0), tpCorContr_Erro
    Exit Sub
  ElseIf Not IsDate(mskData(1).Text) Then
    AmpN
    MsgBox "Data Final Inválida !", vbOKOnly, TITULOSISTEMA
    SetarFoco mskData(1)
    Pintar_Controle mskData(1), tpCorContr_Erro
    Exit Sub
  End If
  '
  If optSai1.Value Then
    Report1.Destination = 0 'Video
  ElseIf optSai2.Value Then
    Report1.Destination = 1   'Impressora
  End If
  Report1.CopiesToPrinter = 1
  Report1.WindowState = crptMaximized
  '
  Report1.Formulas(0) = "DataIni = Date(" & Right(mskData(0).Text, 4) & ", " & Mid(mskData(0).Text, 4, 2) & ", " & Left(mskData(0).Text, 2) & ")"
  Report1.Formulas(1) = "DataFim = Date(" & Right(mskData(1).Text, 4) & ", " & Mid(mskData(1).Text, 4, 2) & ", " & Left(mskData(1).Text, 2) & ")"
  '
  'Obter valores
  Set objRelBalanco = New busSisMed.clsRelBalanco
  objRelBalanco.SelecionarSaldoBalanco curSaldoAnterior, _
                                       datDataSaldoAnterior, _
                                       curReceita, _
                                       curPrestador, _
                                       curDespesa, _
                                       mskData(0).Text, _
                                       mskData(1).Text
  Set objRelBalanco = Nothing
  'Inclui os campos de retorno
  Report1.Formulas(2) = "VrSaldoAnterior = " & IIf(curSaldoAnterior & "" = "", 0, Replace(Replace(curSaldoAnterior, ".", ""), ",", "."))
  Report1.Formulas(3) = "DescrSaldo = '" & IIf(datDataSaldoAnterior = CDate("00:00"), "SALDO ANTERIOR NÃO CADASTRADO", "SALDO ANTERIOR ATÉ " & Format(datDataSaldoAnterior, "DD/MM/YYYY")) & "'"
  Report1.Formulas(4) = "VrReceita = " & IIf(curReceita & "" = "", 0, Replace(Replace(curReceita, ".", ""), ",", "."))
  Report1.Formulas(5) = "VrPrestador = " & IIf(curPrestador & "" = "", 0, Replace(Replace(curPrestador, ".", ""), ",", "."))
  Report1.Action = 1
  '
  AmpN
  Exit Sub
  
TratErro:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub


Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco mskData(0)
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserRelBalancoInc.Form_Activate]"
End Sub


Public Sub MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMed.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMed.clsGeral
  '
  strSql = "SELECT RELBALANCO.PKID, RELBALANCO.DATA, RELBALANCO.SALDO FROM RELBALANCO " & _
      " ORDER BY RELBALANCO.DATA DESC;"
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
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim objGrupoDespesa As busSisMed.clsGrupoDespesa
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 5550
  Me.Width = 10200
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, , , cmdCancelar, cmdExcluir, , cmdInserir, cmdAlterar, pbtnImprimir:=cmdRelatorio
  '
  Report1.Connect = ConnectRpt
  Report1.ReportFileName = gsReportPath & "Balanco.rpt"
  '
  mskData(0).Text = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
  mskData(1).Text = Format(Now, "DD/MM/YYYY")
  '
  If Status = tpStatus_Incluir Then
    '
    cmdExcluir.Enabled = False
    cmdInserir.Enabled = False
    cmdAlterar.Enabled = False
    tabDetalhes.TabEnabled(1) = False
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
'''    'Pega Dados do Banco de dados
'''    Set objGrupoDespesa = New busSisMed.clsGrupoDespesa
'''    Set objRs = objGrupoDespesa.SelecionarGrupoDespesa(lngRELBALANCOID)
'''    '
'''    If Not objRs.EOF Then
'''      INCLUIR_VALOR_NO_MASK mskGrupo, objRs.Fields("CODIGO").Value & "", TpMaskOutros
'''      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
'''      If objRs.Fields("TIPO").Value & "" = "D" Then
'''        cboTipo.Text = "DÉBITO"
'''      ElseIf objRs.Fields("TIPO").Value & "" = "C" Then
'''        cboTipo.Text = "CRÉDITO"
'''      End If
'''    End If
'''    Set objGrupoDespesa = Nothing
    cmdInserir.Enabled = False
    cmdExcluir.Enabled = False
    cmdAlterar.Enabled = False
    cmdRelatorio.Enabled = True
  End If
  AmpN
  Exit Sub
trata:
  AmpN
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
  If Not bFechar Then Cancel = True
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
  TratarErro Err.Number, Err.Description, "[frmUserRelBalancoInc.grdGeral_UnboundReadDataEx]"
End Sub

Private Sub mskData_GotFocus(Index As Integer)
  Seleciona_Conteudo_Controle mskData(Index)
End Sub

Private Sub mskData_LostFocus(Index As Integer)
  Pintar_Controle mskData(Index), tpCorContr_Normal
End Sub


Private Sub tabDetalhes_Click(PreviousTab As Integer)
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    cmdInserir.Enabled = False
    cmdExcluir.Enabled = False
    cmdAlterar.Enabled = False
    cmdRelatorio.Enabled = True
  Case 1
    COLUNASMATRIZ = grdGeral.Columns.Count
    LINHASMATRIZ = 0
  
    MontaMatriz
    grdGeral.ApproxCount = LINHASMATRIZ
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    '
    cmdInserir.Enabled = True
    cmdExcluir.Enabled = True
    cmdAlterar.Enabled = True
    cmdRelatorio.Enabled = False
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

