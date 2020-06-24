VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmInsumoLis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   10080
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   6120
      Left            =   8220
      ScaleHeight     =   6120
      ScaleWidth      =   1860
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4725
         Left            =   90
         ScaleHeight     =   4665
         ScaleWidth      =   1635
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1695
         Begin VB.CommandButton cmdFiltrar 
            Caption         =   "&X"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   150
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1020
            Width           =   1335
         End
         Begin VB.CommandButton cmdInserir 
            Caption         =   "&Z"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1890
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   3630
            Width           =   1335
         End
      End
   End
   Begin TrueDBGrid60.TDBGrid grdGeral 
      Align           =   3  'Align Left
      Height          =   6120
      Left            =   0
      OleObjectBlob   =   "userInsumoLis.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   8160
   End
End
Attribute VB_Name = "frmInsumoLis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intTipoInsumo      As tpInsumo

Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String
'
'UTILIZADO PARA FILTRO
Public strNOMEFIL         As String
Public lngCORIDFIL        As Long

Private Sub cmdAlterar_Click()
  On Error GoTo trata
  Dim strMsgErro          As String
  strMsgErro = ""
  Select Case intTipoInsumo
  Case tpInsumo_Perfil: strMsgErro = "Selecione um perfil para alteração."
  Case tpInsumo_Acessorio: strMsgErro = "Selecione um acessório para alteração."
  Case tpInsumo_Produto: strMsgErro = "Selecione um produto para alteração."
  End Select
  '
  If Not IsNumeric(grdGeral.Columns("PKID").Value & "") Then
    MsgBox strMsgErro, vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If

  frmInsumoInc.intTipoInsumo = intTipoInsumo
  frmInsumoInc.Status = tpStatus_Alterar
  frmInsumoInc.lngPKID = grdGeral.Columns("PKID").Value
  frmInsumoInc.Show vbModal

  If frmInsumoInc.blnRetorno Then
    MontaMatriz strNOMEFIL, _
                lngCORIDFIL
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  End If
  SetarFoco grdGeral
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub


Private Sub cmdExcluir_Click()
  Dim objInsumo         As busSisMetal.clsInsumo
  Dim objRs               As ADODB.Recordset
  Dim strSql              As String
  Dim objGeral            As busSisMetal.clsGeral
  Dim strMsgErro          As String
  Dim strMsgConf          As String
  '
  On Error GoTo trata
  'Cadastro de Insumo
  strMsgErro = ""
  strMsgConf = ""
  Select Case intTipoInsumo
  Case tpInsumo_Perfil
      strMsgErro = "Selecione um perfil para exclusão."
      strMsgConf = "Confirma exclusão do perfil "
  Case tpInsumo_Acessorio
      strMsgErro = "Selecione um acessório para exclusão."
      strMsgConf = "Confirma exclusão do acessório "
  Case tpInsumo_Produto
      strMsgErro = "Selecione um produto para exclusão."
      strMsgConf = "Confirma exclusão do produto "
  End Select
  '
  If Not IsNumeric(grdGeral.Columns("PKID").Value & "") Then
    MsgBox strMsgErro, vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If
  '
  If MsgBox(strMsgConf & grdGeral.Columns("Nome").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
    SetarFoco grdGeral
    Exit Sub
  End If
  'OK
  Set objInsumo = New busSisMetal.clsInsumo
  objInsumo.ExcluirInsumo CLng(grdGeral.Columns("PKID").Value)
  Set objInsumo = Nothing
  '
  MontaMatriz strNOMEFIL, _
              lngCORIDFIL
  grdGeral.Bookmark = Null
  grdGeral.ReBind

  SetarFoco grdGeral
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdFiltrar_Click()
  Dim objInsumoFil As frmInsumoFil
  On Error GoTo trata
  
  Set objInsumoFil = New frmInsumoFil
  objInsumoFil.intTipoInsumo = intTipoInsumo
  objInsumoFil.Show vbModal
  '
  If objInsumoFil.blnRetorno = True Then
    strNOMEFIL = objInsumoFil.strNOME
    lngCORIDFIL = objInsumoFil.lngCORID
  
    MontaMatriz strNOMEFIL, _
                lngCORIDFIL
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  End If
  Set objInsumoFil = Nothing
  '
  SetarFoco grdGeral
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub cmdInserir_Click()
  On Error GoTo trata
  frmInsumoInc.intTipoInsumo = intTipoInsumo
  frmInsumoInc.Status = tpStatus_Incluir
  frmInsumoInc.lngPKID = 0
  frmInsumoInc.Show vbModal

  If frmInsumoInc.blnRetorno Then
    MontaMatriz strNOMEFIL, _
                lngCORIDFIL
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  End If
  SetarFoco grdGeral
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  AmpS
  Me.Height = 6600
  Me.Width = 10170

  CenterForm Me

  LerFiguras Me, tpBmp_Vazio, , , cmdFechar, cmdExcluir, , cmdInserir, cmdAlterar
  LerFigurasAvulsas cmdFiltrar, "FILTRAR.ICO", "filtrarDown.ico", "Filtrar"
  Select Case intTipoInsumo
  Case tpInsumo_Perfil: Me.Caption = "Lista de Perfil"
  Case tpInsumo_Acessorio: Me.Caption = "Lista de Acessório"
  Case tpInsumo_Produto: Me.Caption = "Lista de Produto"
  End Select
  'FILTRO
  strNOMEFIL = ""
  lngCORIDFIL = 0
  '
  COLUNASMATRIZ = grdGeral.Columns.Count
  LINHASMATRIZ = 0

  MontaMatriz strNOMEFIL, _
              lngCORIDFIL
  grdGeral.ApproxCount = LINHASMATRIZ

  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Public Sub MontaMatriz(strNOME As String, _
                       lngCORID As Long)
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMetal.clsGeral
  '
  On Error GoTo trata

  Set clsGer = New busSisMetal.clsGeral
  '
  'Chamada do Cadastro de Insumos
  strSql = "SELECT INSUMO.PKID, "
  Select Case intTipoInsumo
  Case tpInsumo_Perfil: strSql = strSql & " TIPO_LINHA.NOME + ' - ' + LINHA.CODIGO + ' - ' + COR.NOME AS NOME_INSUMO  "
  Case tpInsumo_Acessorio: strSql = strSql & " ACESSORIO.NOME AS NOME_INSUMO  "
  Case tpInsumo_Produto: strSql = strSql & " PRODUTO.NOME AS NOME_INSUMO "
  End Select
  strSql = strSql & " , INSUMO.CODIGO  FROM INSUMO "
  Select Case intTipoInsumo
  Case tpInsumo_Perfil: strSql = strSql & " INNER JOIN PERFIL ON INSUMO.PKID = PERFIL.INSUMOID LEFT JOIN COR ON COR.PKID = PERFIL.CORID INNER JOIN LINHA ON LINHA.PKID = PERFIL.LINHAID INNER JOIN TIPO_LINHA ON TIPO_LINHA.PKID = LINHA.TIPO_LINHAID "
  Case tpInsumo_Acessorio: strSql = strSql & " INNER JOIN ACESSORIO ON INSUMO.PKID = ACESSORIO.INSUMOID "
  Case tpInsumo_Produto: strSql = strSql & " INNER JOIN PRODUTO ON INSUMO.PKID = PRODUTO.INSUMOID "
  End Select
  'FILTRO
  strSql = strSql & " WHERE 0 = 0 "
  Select Case intTipoInsumo
  Case tpInsumo_Perfil
    If strNOME <> "" Then
      strSql = strSql & " AND (LINHA.CODIGO LIKE " & Formata_Dados("%" & strNOME & "%", tpDados_Texto) & _
        " OR TIPO_LINHA.NOME LIKE " & Formata_Dados("%" & strNOME & "%", tpDados_Texto) & ")"
    End If
    If lngCORID <> 0 Then
      strSql = strSql & " AND COR.PKID = " & Formata_Dados(lngCORID, tpDados_Longo)
    End If
  Case tpInsumo_Acessorio
    If strNOME <> "" Then
      strSql = strSql & " AND ACESSORIO.NOME LIKE " & Formata_Dados("%" & strNOME & "%", tpDados_Texto)
    End If
  Case tpInsumo_Produto
    If strNOME <> "" Then
      strSql = strSql & " AND PRODUTO.NOME LIKE " & Formata_Dados("%" & strNOME & "%", tpDados_Texto)
    End If
  End Select
  '
  strSql = strSql & " ORDER BY NOME_INSUMO"
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
  TratarErro Err.Number, Err.Description, "[frmInsumoLis.grdGeral_UnboundReadDataEx]"
End Sub




