VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserAssociadoLis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de associados"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   11865
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   11865
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   11865
      Begin VB.CommandButton cmdNormal 
         Caption         =   "&Consultar"
         Height          =   255
         Index           =   0
         Left            =   8550
         TabIndex        =   12
         Top             =   330
         Width           =   1215
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Text            =   "txtNome"
         Top             =   330
         Width           =   7575
      End
      Begin MSMask.MaskEdBox mskCpf 
         Height          =   255
         Left            =   840
         TabIndex        =   0
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   450
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   14
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskMatricula 
         Height          =   255
         Left            =   4050
         TabIndex        =   1
         Top             =   60
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,###;($#,###)"
         PromptChar      =   "_"
      End
      Begin VB.Label Matricula 
         Caption         =   "Matrícula"
         Height          =   195
         Index           =   0
         Left            =   2790
         TabIndex        =   15
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "CPF"
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Nome"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   13
         Top             =   330
         Width           =   795
      End
   End
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7440
      Left            =   10005
      ScaleHeight     =   7440
      ScaleWidth      =   1860
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   705
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   3885
         Left            =   90
         ScaleHeight     =   3825
         ScaleWidth      =   1635
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   750
         Width           =   1695
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   150
            Width           =   1335
         End
         Begin VB.CommandButton cmdInserir 
            Caption         =   "&Z"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1020
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1890
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   2760
            Width           =   1335
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo = D - Dependente; T - Titular"
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   60
         TabIndex        =   10
         Top             =   30
         Width           =   1725
      End
   End
   Begin TrueDBGrid60.TDBGrid grdGeral 
      Align           =   3  'Align Left
      Height          =   7440
      Left            =   0
      OleObjectBlob   =   "userAssociadoLis.frx":0000
      TabIndex        =   3
      Top             =   705
      Width           =   9930
   End
End
Attribute VB_Name = "frmUserAssociadoLis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String


Private Sub cmdAlterar_Click()
  On Error GoTo trata
  If Not IsNumeric(grdGeral.Columns("PKID").Value & "") Then
    MsgBox "Selecione um associado !", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If
  If grdGeral.Columns("Tipo").Value & "" <> "T" Then
    MsgBox "Selecione um associado do tipo titular!", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If
  
  frmUserAssociadoInc.Status = tpStatus_Alterar
  frmUserAssociadoInc.lngPKID = grdGeral.Columns("PKID").Value
  frmUserAssociadoInc.strIcAssociado = "T"
  frmUserAssociadoInc.Show vbModal
  
  If frmUserAssociadoInc.blnRetorno Then
    MontaMatriz
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
  Dim objAssociado        As busApler.clsAssociado
  Dim objRs               As ADODB.Recordset
  Dim strSql              As String
  Dim objGeral            As busApler.clsGeral
  '
  On Error GoTo trata
  'Exclusão de associado
  If Not IsNumeric(grdGeral.Columns("PKID").Value & "") Then
    MsgBox "Selecione um associado para exclusão.", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If
  '
  If MsgBox("ATENÇÃO: A exclusão do associado removerá todas associações de pagamento e convênios." & vbCrLf & "Caso queira você pode apenas alterá-lo e selecionar a opção excluído, isso irá excluílo logicamente, mantendo suas informações na base de dados." & vbCrLf & "Confirma exclusão do associado " & grdGeral.Columns("Nome").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
    SetarFoco grdGeral
    Exit Sub
  End If
  'OK
  Set objAssociado = New busApler.clsAssociado
  objAssociado.ExcluirAssociado CLng(grdGeral.Columns("PKID").Value), _
                                grdGeral.Columns("ICASSOCIADO").Value
  Set objAssociado = Nothing
  '
  MontaMatriz
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



Private Sub cmdInserir_Click()
  On Error GoTo trata
  frmUserAssociadoInc.Status = tpStatus_Incluir
  frmUserAssociadoInc.strIcAssociado = "T"
  frmUserAssociadoInc.lngPKID = 0
  frmUserAssociadoInc.Show vbModal

  If frmUserAssociadoInc.blnRetorno Then
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  End If
  SetarFoco grdGeral
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCampos = False
  If Not Valida_Moeda(mskMatricula, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a matrícula válida" & vbCrLf
  End If
'''  If Len(Trim(mskCpf.ClipText)) > 0 Then
'''    If Not TestaCPF(mskCpf.ClipText) Then
'''      strMsg = strMsg & "Preencher o CPF válido" & vbCrLf
'''      Pintar_Controle mskCpf, tpCorContr_Erro
'''      SetarFoco mskCpf
'''      tabDetalhes.Tab = 0
'''      blnSetarFocoControle = False
'''    End If
'''  End If
  
  If mskCpf.Text = "___.___.___-__" And txtNome.Text = "" And mskMatricula.Text = "" Then
    strMsg = strMsg & "Preencher o nome, cpf ou matrícula" & vbCrLf
    SetarFoco mskCpf
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserAssociadoLis.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserAssociadoLis.ValidaCampos]", _
            Err.Description
End Function

Private Sub cmdNormal_Click(Index As Integer)
  On Error GoTo trata
  '
  If Index = 0 Then
    If Not ValidaCampos Then
      Exit Sub
    End If
    COLUNASMATRIZ = grdGeral.Columns.Count
    LINHASMATRIZ = 0
    MontaMatriz mskCpf.ClipText, _
                txtNome.Text, _
                mskMatricula
                   
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = LINHASMATRIZ
  
    grdGeral.SetFocus
  End If
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source, _
            Err.Description
End Sub

Private Sub Form_Activate()
  SetarFoco grdGeral
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  AmpS
  If Me.ActiveControl Is Nothing Then
    'Tela
    Me.Height = Screen.Height - 1450
    Me.Width = 11955
    CenterForm Me
  End If
  'Me.Height = 5355
  'Me.Width = 10170
  
  CenterForm Me
  
  LerFiguras Me, tpBmp_Vazio, , , cmdFechar, cmdExcluir, , cmdInserir, cmdAlterar
  
  'Limpar campos
  LimparCampoMask mskCpf
  LimparCampoMask mskMatricula
  LimparCampoTexto txtNome
  '
  COLUNASMATRIZ = grdGeral.Columns.Count
  LINHASMATRIZ = 0

  MontaMatriz
  grdGeral.ApproxCount = LINHASMATRIZ
  
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Public Sub MontaMatriz(Optional strCpf As String, _
                       Optional strNome As String, _
                       Optional strMatricula As String)
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busApler.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busApler.clsGeral
  '
  'Chamada do Cadastro de CLientes
  strSql = "SELECT ASSOCIADO.PKID, ASSOCIADO.ICASSOCIADO, ASSOCIADO.NOME, ASSOCIADO.ICASSOCIADO, ASSOCIADO.CPF, CONVERT(VARCHAR(50), ISNULL(DEPENDENTE.MATRICULADEP,'')) + CONVERT(VARCHAR(50), ISNULL(TITULAR.MATRICULA, '')), ASSOCIADO.DATANASCIMENTO "
  strSql = strSql & " FROM ASSOCIADO "
  strSql = strSql & " LEFT JOIN DEPENDENTE ON ASSOCIADO.PKID = DEPENDENTE.ASSOCIADOID "
  strSql = strSql & " LEFT JOIN TITULAR ON ASSOCIADO.PKID = TITULAR.ASSOCIADOID "
  strSql = strSql & " WHERE 1 = 1 "
  If strCpf & "" <> "" Then
    strSql = strSql & " AND CPF = " & Formata_Dados(strCpf, tpDados_Texto)
  End If
  If strNome & "" <> "" Then
    strSql = strSql & " AND NOME LIKE " & Formata_Dados(strNome & "%", tpDados_Texto)
  End If
  If strMatricula & "" <> "" Then
    strSql = strSql & " AND (DEPENDENTE.MATRICULADEP = " & Formata_Dados(strMatricula, tpDados_Texto)
    strSql = strSql & " OR TITULAR.MATRICULA = " & Formata_Dados(strMatricula, tpDados_Texto) & ")"
  End If
  strSql = strSql & " ORDER BY ASSOCIADO.NOME"
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
  TratarErro Err.Number, Err.Description, "[frmUserAssociado.grdGeral_UnboundReadDataEx]"
End Sub




Private Sub mskCpf_GotFocus()
  Seleciona_Conteudo_Controle mskCpf
End Sub
Private Sub mskCpf_LostFocus()
  Pintar_Controle mskCpf, tpCorContr_Normal
End Sub

Private Sub mskMatricula_GotFocus()
  Seleciona_Conteudo_Controle mskMatricula
End Sub
Private Sub mskMatricula_LostFocus()
  Pintar_Controle mskMatricula, tpCorContr_Normal
End Sub

Private Sub txtNome_GotFocus()
  Seleciona_Conteudo_Controle txtNome
End Sub
Private Sub txtNome_LostFocus()
  Pintar_Controle txtNome, tpCorContr_Normal
End Sub

