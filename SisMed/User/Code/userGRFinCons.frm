VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserGRFinCons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "V"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboPeriodo 
      Height          =   315
      Left            =   6570
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   30
      Width           =   5325
   End
   Begin VB.Frame fraUnidade 
      Caption         =   "GR´s"
      Height          =   6015
      Left            =   60
      TabIndex        =   11
      Top             =   330
      Width           =   11835
      Begin TrueDBGrid60.TDBGrid grdGeral 
         Height          =   5730
         Left            =   90
         OleObjectBlob   =   "userGRFinCons.frx":0000
         TabIndex        =   3
         Top             =   210
         Width           =   11580
      End
   End
   Begin VB.CommandButton cmdSairSelecao 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   855
      Left            =   11070
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6510
      Width           =   900
   End
   Begin VB.Frame Frame8 
      Caption         =   "Selecione a opção"
      Height          =   1725
      Left            =   60
      TabIndex        =   10
      Top             =   6420
      Width           =   10935
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&C - GR Expirada       "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3000
         TabIndex        =   6
         ToolTipText     =   "Incluir GR"
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&B - Liberar para atend "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1530
         TabIndex        =   5
         ToolTipText     =   "Incluir GR"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&A - GR não atendida   "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   4
         ToolTipText     =   "Atender GR"
         Top             =   240
         Width           =   1455
      End
      Begin ComctlLib.StatusBar StatusBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1350
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   450
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   5
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   6
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   1940
               MinWidth        =   1940
               TextSave        =   "17/5/2012"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   5
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   1235
               MinWidth        =   1235
               TextSave        =   "00:00"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   1
               Alignment       =   1
               Bevel           =   2
               Enabled         =   0   'False
               Object.Width           =   1235
               MinWidth        =   1235
               TextSave        =   "CAPS"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   2
               Alignment       =   1
               Bevel           =   2
               Enabled         =   0   'False
               Object.Width           =   1235
               MinWidth        =   1235
               TextSave        =   "NUM"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   3
               Alignment       =   1
               AutoSize        =   2
               Bevel           =   2
               Enabled         =   0   'False
               Object.Width           =   1244
               MinWidth        =   1235
               TextSave        =   "INS"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtUsuario 
      BackColor       =   &H00E0E0E0&
      Height          =   288
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "txtUsuario"
      Top             =   30
      Width           =   1695
   End
   Begin MSMask.MaskEdBox mskDataPrinc 
      Height          =   255
      Left            =   3990
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
      MaxLength       =   10
      Format          =   "DD/MMM/YYYY"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   "Últimos"
      Height          =   255
      Left            =   5340
      TabIndex        =   20
      Top             =   60
      Width           =   1155
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
      Index           =   5
      Left            =   8430
      TabIndex        =   19
      Top             =   8220
      Width           =   1785
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
      Index           =   4
      Left            =   6120
      TabIndex        =   18
      Top             =   8220
      Width           =   2265
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Expirada"
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
      Left            =   5160
      TabIndex        =   17
      Top             =   8220
      Width           =   915
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
      Index           =   2
      Left            =   2820
      TabIndex        =   16
      Top             =   8220
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
      Index           =   8
      Left            =   1590
      TabIndex        =   15
      Top             =   8220
      Width           =   1185
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
      Left            =   870
      TabIndex        =   14
      Top             =   8220
      Width           =   675
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
      TabIndex        =   13
      Top             =   8220
      Width           =   765
   End
   Begin VB.Label Label16 
      Caption         =   "Data"
      Height          =   255
      Left            =   3150
      TabIndex        =   9
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label22 
      Caption         =   "Usuário Logado"
      Height          =   255
      Left            =   30
      TabIndex        =   8
      Top             =   30
      Width           =   1215
   End
End
Attribute VB_Name = "frmUserGRFinCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''
Option Explicit

Public nGrupo                   As Integer
Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public blnPrimeiraVez           As Boolean 'Propósito: Preencher lista no combo

Private COLUNASMATRIZ           As Long
Private LINHASMATRIZ            As Long
Private Matriz()                As String

Private datDataIniAtual         As Date
Private datDataFimAtual         As Date

Private Sub cboPeriodo_Click()
  On Error GoTo trata
  If blnPrimeiraVez = False Then
    DoEvents
    '
    COLUNASMATRIZ = grdGeral.Columns.Count
    LINHASMATRIZ = 0

    MontaMatriz cboPeriodo.Text
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = LINHASMATRIZ
    blnPrimeiraVez = False
    SetarFoco grdGeral
  End If
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source

End Sub

Private Sub cmdSairSelecao_Click()
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

Public Sub cmdSelecao_Click(Index As Integer)
  On Error GoTo trata
  nGrupo = Index
  'strNumeroAptoPrinc = optUnidade
  'If Not ValiCamposPrinc Then Exit Sub
  VerificaQuemChamou
  'Atualiza Valores
  '
  COLUNASMATRIZ = grdGeral.Columns.Count
  LINHASMATRIZ = 0

  'MontaMatriz
  MontaMatriz cboPeriodo.Text
  grdGeral.Bookmark = Null
  grdGeral.ReBind
  grdGeral.ApproxCount = LINHASMATRIZ
  blnPrimeiraVez = False
  SetarFoco grdGeral
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[cmdSelecao_Click]"
  frmMDI.tmrUnidade.Enabled = True
End Sub


Public Sub VerificaQuemChamou()
  Dim objGR As busSisMed.clsGR

  On Error GoTo trata
  '
  Select Case nGrupo

  Case 0
  
    'GR Não atendida
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione uma GR para atendimento.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") <> "F" Then
      MsgBox "Apenas uma GR fechada pode não ser atendida.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    'Confirmação
    If MsgBox("Confirma não atendimento da GR " & grdGeral.Columns("Seq.").Value & " de " & grdGeral.Columns("Prontuário").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
      SetarFoco grdGeral
      Exit Sub
    End If
    
    Set objGR = New busSisMed.clsGR
    '
    objGR.AlterarGRNaoAtendida grdGeral.Columns("ID").Value & ""
    Set objGR = Nothing
  
  
  Case 1
    'GR LIBERADA para atendimento
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione uma GR para atendimento.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") <> "N" Then
      If Trim(grdGeral.Columns("Status").Value & "") <> "Z" Then
        MsgBox "Apenas uma GR não atendida pode pode ser liberada para atendimento.", vbExclamation, TITULOSISTEMA
        SetarFoco grdGeral
        Exit Sub
      End If
    End If
    'Confirmação
    If MsgBox("Confirma liberação para atendimento da GR " & grdGeral.Columns("Seq.").Value & " de " & grdGeral.Columns("Prontuário").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
      SetarFoco grdGeral
      Exit Sub
    End If
    
    Set objGR = New busSisMed.clsGR
    '
    objGR.AlterarGRLiberarAtend grdGeral.Columns("ID").Value & ""
    Set objGR = Nothing
  
  
  Case 2
    'GR Expirada
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione uma GR para atendimento.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") <> "Z" Then
      MsgBox "Apenas uma GR que já tenha expirado o prazo pode ser lançada.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    'Confirmação
    If MsgBox("Confirma lançar expiração da GR " & grdGeral.Columns("Seq.").Value & " de " & grdGeral.Columns("Prontuário").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
      SetarFoco grdGeral
      Exit Sub
    End If
    
    Set objGR = New busSisMed.clsGR
    '
    objGR.AlterarGRExpirarAtend grdGeral.Columns("ID").Value & ""
    Set objGR = Nothing
  End Select
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  End
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim strSql            As String
  Dim datDataTurno      As Date
  '
  'OK Para turno
  datDataIniAtual = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now) & " 00:00:00")
  datDataFimAtual = DateAdd("d", 1, datDataIniAtual)
  '
  blnFechar = False
  blnRetorno = False
  blnPrimeiraVez = True
  AmpS
  If Me.ActiveControl Is Nothing Then
    Me.Top = 580
    Me.Left = 1
    Me.WindowState = 2 'Maximizado
  End If
  'Me.Height = 9195
  'Me.Width = 12090
  'CenterForm Me
  LerFigurasAvulsas cmdSairSelecao, "Sair.ico", "SairDown.ico", "Sair"
  cboPeriodo.AddItem "Data atual"
  cboPeriodo.AddItem "05 últimos dias"
  cboPeriodo.AddItem "10 últimos dias"
  cboPeriodo.AddItem "20 últimos dias"
  cboPeriodo.AddItem "30 últimos dias"
  cboPeriodo.AddItem "60 últimos dias"
  cboPeriodo.AddItem "90 últimos dias"
  cboPeriodo.Text = "Data atual"
  '
  txtUsuario.Text = gsNomeUsu
  mskDataPrinc.Text = Format(Date, "DD/MM/YYYY")

  'NOVO BOTÕES NOVOS
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
  TratarErro Err.Number, Err.Description, "[frmUserGRFinCons.grdGeral_UnboundReadDataEx]"
End Sub



Private Sub Form_Activate()
  If blnPrimeiraVez Then
    DoEvents
    '
    COLUNASMATRIZ = grdGeral.Columns.Count
    LINHASMATRIZ = 0

    'MontaMatriz
    MontaMatriz cboPeriodo.Text
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = LINHASMATRIZ
    blnPrimeiraVez = False
    SetarFoco grdGeral
  End If
End Sub

Public Sub MontaMatriz(strPeriodo As String)
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGR     As busSisMed.clsGR
  Dim strPerFinal As String
  '
  AmpS
  On Error GoTo trata
  '
  If Not IsNumeric(Left(strPeriodo, 2)) Then
    strPerFinal = "01"
  Else
    strPerFinal = Left(strPeriodo, 2)
  End If
  'A data inicial passa a ser calculada de acordo com o período informado
  datDataIniAtual = DateAdd("d", CInt(strPerFinal) * -1, datDataFimAtual)
  Set objGR = New busSisMed.clsGR
  '
  Set objRs = objGR.CapturaGRTurnoCorrenteFIN(Format(datDataIniAtual, "DD/MM/YYYY hh:mm"), _
                                              Format(datDataFimAtual, "DD/MM/YYYY hh:mm"), _
                                              giMaxDiasAtend)
  If Not objRs.EOF Then
    'objRs.Filter = "STATUS = 'F' OR STATUS = 'A'"
    'objRs.Filter = "STATUS = 'F'"
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
