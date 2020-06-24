VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAjusteInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Ajustes"
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
      TabIndex        =   11
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
      TabCaption(0)   =   "Controle de ajustes"
      TabPicture(0)   =   "userAjusteInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "grdItemAjuste"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraFiltro"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraAjuste"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame fraAjuste 
         Caption         =   "Pedido"
         Height          =   1095
         Left            =   90
         TabIndex        =   15
         Top             =   1200
         Width           =   9195
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   735
            Index           =   0
            Left            =   90
            ScaleHeight     =   735
            ScaleWidth      =   8895
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   180
            Width           =   8895
            Begin VB.TextBox txtUsuario 
               BackColor       =   &H00E0E0E0&
               Height          =   288
               Left            =   1230
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   3
               TabStop         =   0   'False
               Text            =   "txtUsuario"
               Top             =   30
               Width           =   1815
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   255
               Left            =   4320
               ScaleHeight     =   255
               ScaleWidth      =   3855
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   30
               Width           =   3855
               Begin MSMask.MaskEdBox mskData 
                  Height          =   255
                  Index           =   0
                  Left            =   1200
                  TabIndex        =   4
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  _Version        =   393216
                  BackColor       =   14737632
                  AutoTab         =   -1  'True
                  MaxLength       =   16
                  Mask            =   "##/##/#### ##:##"
                  PromptChar      =   "_"
               End
               Begin VB.Label Label2 
                  Caption         =   "Data"
                  Height          =   255
                  Index           =   1
                  Left            =   450
                  TabIndex        =   18
                  Top             =   0
                  Width           =   615
               End
            End
            Begin VB.ComboBox cboAjuste 
               Height          =   315
               Left            =   1230
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   360
               Width           =   3435
            End
            Begin VB.Label Label2 
               Caption         =   "Usuário"
               Height          =   255
               Index           =   0
               Left            =   30
               TabIndex        =   20
               Top             =   30
               Width           =   1155
            End
            Begin VB.Label Label5 
               Caption         =   "Ajuste"
               Height          =   195
               Index           =   1
               Left            =   30
               TabIndex        =   19
               Top             =   390
               Width           =   1215
            End
         End
      End
      Begin VB.Frame fraFiltro 
         Caption         =   "Filtro"
         Height          =   885
         Left            =   90
         TabIndex        =   14
         Top             =   330
         Width           =   9195
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Left            =   1290
            MaxLength       =   50
            TabIndex        =   0
            Text            =   "txtCodigo"
            Top             =   180
            Width           =   5865
         End
         Begin VB.TextBox txtLinhaFim 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   3660
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   2
            TabStop         =   0   'False
            Text            =   "txtLinhaFim"
            Top             =   540
            Width           =   3495
         End
         Begin VB.TextBox txtCodigoFim 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   1290
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   1
            TabStop         =   0   'False
            Text            =   "txtCodigoFim"
            Top             =   540
            Width           =   2355
         End
         Begin VB.Label Label1 
            Caption         =   "Nome da Linha/Código Perfil"
            Height          =   615
            Index           =   3
            Left            =   90
            TabIndex        =   21
            Top             =   210
            Width           =   1095
         End
      End
      Begin TrueDBGrid60.TDBGrid grdItemAjuste 
         Height          =   3660
         Left            =   90
         OleObjectBlob   =   "userAjusteInc.frx":001C
         TabIndex        =   6
         Top             =   2370
         Width           =   9210
      End
      Begin VB.Label Label1 
         Caption         =   $"userAjusteInc.frx":6F83
         ForeColor       =   &H000000FF&
         Height          =   435
         Index           =   2
         Left            =   120
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   6090
         Width           =   8715
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7245
      Left            =   9660
      ScaleHeight     =   7245
      ScaleWidth      =   1860
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2115
         Left            =   120
         ScaleHeight     =   2055
         ScaleWidth      =   1545
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   4980
         Width           =   1605
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Height          =   885
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   150
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   885
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1020
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmAjusteInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public lngAJUSTEID              As Long
Public lngLINHAID               As Long

Public lngOSFINALID             As Long
Public lngANODIZACAOITEMID      As Long
Public lngITEMOSFINALID         As Long
Public lngOSID                  As Long
Public lngCORID                 As Long

Public strOSNumero              As String
Public strCor                   As String
Dim blnAlterouPeso              As Boolean

Dim blnFechar                   As Boolean
Public blnRetorno               As Boolean
Public blnPrimeiraVez           As Boolean
'
Dim ITEMAJU_COLUNASMATRIZ        As Long
Dim ITEMAJU_LINHASMATRIZ         As Long
Private ITEMAJU_Matriz()         As String


Public Sub ITEMAJU_MontaMatriz(lngLINHASELID As Long)
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGeral  As busSisMetal.clsGeral
  '
  On Error GoTo trata

  Set objGeral = New busSisMetal.clsGeral
  '
  strSql = "SELECT " & _
            IIf(Status = tpStatus_Incluir, "0,", " ITEM_AJUSTE.PKID, ") & _
            " PERFIL.INSUMOID, " & _
            " PERFIL.INSUMOID, " & _
            Formata_Dados(IIf(Status = tpStatus_Incluir, "0", "0"), tpDados_Texto) & ", " & _
            " TIPO_LINHA.NOME, LINHA.CODIGO, COR.NOME, " & _
            " PERFIL.PESO_ESTOQUE, "
  
  If Status = tpStatus_Incluir Then
    strSql = strSql & " '' AS QUANTIDADE, '' AS PESO "
  Else
    strSql = strSql & " ITEM_AJUSTE.QUANTIDADE, DBO.UFN_CALCULA_PESO(PERFIL.LINHAID, ITEM_AJUSTE.QUANTIDADE) AS PESO "
  End If
  
  strSql = strSql & " From PERFIL "
  strSql = strSql & " INNER JOIN LINHA ON LINHA.PKID = PERFIL.LINHAID "
  strSql = strSql & " INNER JOIN TIPO_LINHA ON TIPO_LINHA.PKID = LINHA.TIPO_LINHAID "
  strSql = strSql & " INNER JOIN COR ON COR.PKID = PERFIL.CORID "
  If Status = tpStatus_Incluir Then
    If lngLINHASELID <> 0 Then
      strSql = strSql & " AND PERFIL.LINHAID = " & Formata_Dados(lngLINHASELID, tpDados_Longo)
    End If
  Else
    strSql = strSql & " INNER JOIN ITEM_AJUSTE ON ITEM_AJUSTE.PERFILID = PERFIL.INSUMOID "
    strSql = strSql & " WHERE ITEM_AJUSTE.AJUSTEID = " & Formata_Dados(lngAJUSTEID, tpDados_Longo)
  End If
  strSql = strSql & " ORDER BY TIPO_LINHA.NOME, LINHA.CODIGO, COR.NOME"
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    ITEMAJU_LINHASMATRIZ = objRs.RecordCount
  Else
    ITEMAJU_LINHASMATRIZ = 0
  End If
  If Not objRs.EOF Then
    ReDim ITEMAJU_Matriz(0 To ITEMAJU_COLUNASMATRIZ - 1, 0 To ITEMAJU_LINHASMATRIZ - 1)
  Else
    ReDim ITEMAJU_Matriz(0 To ITEMAJU_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To ITEMAJU_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To ITEMAJU_COLUNASMATRIZ - 1  'varre as colunas
          If intJ = ITEMAJU_COLUNASMATRIZ - 1 Then
            ITEMAJU_Matriz(intJ, intI) = intI & ""
          Else
            ITEMAJU_Matriz(intJ, intI) = objRs(intJ) & ""
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

Private Sub cboAjuste_LostFocus()
  Pintar_Controle cboAjuste, tpCorContr_Normal
End Sub
Private Sub cmdOK_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim objAjuste               As busSisMetal.clsAjuste
  Dim objRs                   As ADODB.Recordset
  Dim objGer                  As busSisMetal.clsGeral
  Dim lngTIPO_AJUSTEID        As Long
  Dim objItemAjuste           As busSisMetal.clsItemAjuste
  Dim intI      As Integer
  '
  Select Case tabDetalhes.Tab
  Case 0 'Gravar Ajuste
    If Not ValidaCampos Then Exit Sub

    If ValidaCamposAnodOrigemAll Then
      SetarFoco grdItemAjuste
      grdItemAjuste.Col = 8
      grdItemAjuste.Row = 0
      Exit Sub
    End If
    'OK procede com o cadastro
    'CADASTRO DE AJUSTE
    '-------------------------
    Set objGer = New busSisMetal.clsGeral
    'TIPO_AJUSTE
    lngTIPO_AJUSTEID = 0
    strSql = "SELECT TIPO_AJUSTE.PKID FROM TIPO_AJUSTE " & _
      " WHERE TIPO_AJUSTE.DESCRICAO = " & Formata_Dados(cboAjuste.Text, tpDados_Texto)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngTIPO_AJUSTEID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGer = Nothing
    '
    Set objAjuste = New busSisMetal.clsAjuste
    'Altera ou incluiu pedido
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objAjuste.AlterarAjuste lngAJUSTEID, _
                              lngTIPO_AJUSTEID, _
                              txtUsuario.Text
      '
      blnRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Código para inclusão
      '
      objAjuste.InserirAjuste lngAJUSTEID, _
                              lngTIPO_AJUSTEID, _
                              txtUsuario.Text
      '
      blnRetorno = True
    End If
    Set objAjuste = Nothing
    '
    Set objItemAjuste = New busSisMetal.clsItemAjuste
    For intI = 0 To ITEMAJU_LINHASMATRIZ - 1
      grdItemAjuste.Bookmark = CLng(intI)
      'If grdItemAjuste.Columns("Branco").Text & "" <> "" Or _
        grdItemAjuste.Columns("Brilho").Text & "" <> "" Or _
        grdItemAjuste.Columns("Bronze").Text & "" <> "" Or _
        grdItemAjuste.Columns("Natural").Text & "" <> "" Then
      If grdItemAjuste.Columns("*").Text & "" = "-1" Then
        'Propósito: Cadastrar pedido
        '
        objItemAjuste.InserirItemAjuste grdItemAjuste.Columns("ITEM_AJUSTEID").Text & "", _
                                        lngAJUSTEID, _
                                        grdItemAjuste.Columns("PERFILID").Text & "", _
                                        IIf(grdItemAjuste.Columns("Quantidade").Text & "" = "", "0", grdItemAjuste.Columns("Quantidade").Text & "")

        blnRetorno = True
      End If
    Next
    Set objItemAjuste = Nothing
    '
    blnFechar = True
    Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  '
  grdItemAjuste.Bookmark = Null
  grdItemAjuste.ReBind
  SetarFoco grdItemAjuste
  If grdItemAjuste.Row <> -1 Then
    grdItemAjuste.Col = 8
    grdItemAjuste.Row = 0
  End If
End Sub

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Filtro
  LimparCampoTexto txtCodigo
  LimparCampoTexto txtCodigoFim
  LimparCampoTexto txtLinhaFim
  'Ajuste
  LimparCampoTexto txtUsuario
  LimparCampoMask mskData(0)
  LimparCampoCombo cboAjuste
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmAjusteInc.LimparCampos]", _
            Err.Description
End Sub

Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  '
  If Not Valida_String(cboAjuste, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o ajuste" & vbCrLf
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmAjusteInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[frmAjusteInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Montar RecordSet
    ITEMAJU_COLUNASMATRIZ = grdItemAjuste.Columns.Count
    ITEMAJU_LINHASMATRIZ = 0
    ITEMAJU_MontaMatriz (lngLINHAID)
    grdItemAjuste.Bookmark = Null
    grdItemAjuste.ReBind
    grdItemAjuste.ApproxCount = ITEMAJU_LINHASMATRIZ
    '
    If Status = tpStatus_Incluir Then
      SetarFoco txtCodigo
    ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
      SetarFoco cboAjuste
    End If
    'SetarFoco grdItemAjuste
    'grdItemAjuste.Col = 8
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmAjusteInc.Form_Activate]"
End Sub


Private Sub grdItemAjuste_BeforeUpdate(Cancel As Integer)
  On Error GoTo trata
  'Atualiza Matriz
  If blnAlterouPeso = True Then
    ITEMAJU_Matriz(3, grdItemAjuste.Columns("ROWNUM").Value) = "-1"
  Else
    ITEMAJU_Matriz(3, grdItemAjuste.Columns("ROWNUM").Value) = grdItemAjuste.Columns(3).Text
  End If
  ITEMAJU_Matriz(8, grdItemAjuste.Columns("ROWNUM").Value) = grdItemAjuste.Columns(8).Text
  ITEMAJU_Matriz(9, grdItemAjuste.Columns("ROWNUM").Value) = grdItemAjuste.Columns(9).Text
  blnAlterouPeso = False
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmAjusteInc.grdItemAjuste_BeforeRowColChange]"
End Sub

Private Sub grdItemAjuste_ColEdit(ByVal ColIndex As Integer)
  On Error GoTo trata
  '
  If grdItemAjuste.Col = 8 Or grdItemAjuste.Col = 9 Or grdItemAjuste.Col = 10 Then
    blnAlterouPeso = True
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmAjusteInc.grdItemAjuste_ColEdit]"
End Sub

Private Sub grdItemAjuste_GotFocus()
  On Error Resume Next
  If Status = tpStatus_Incluir Or gsNivel <> gsCompra Then
    grdItemAjuste.Col = 8
  Else
    grdItemAjuste.Col = 9
  End If
End Sub


Private Sub grdItemAjuste_UnboundReadDataEx( _
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
               Offset + intI, ITEMAJU_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, ITEMAJU_COLUNASMATRIZ, ITEMAJU_LINHASMATRIZ, ITEMAJU_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ITEMAJU_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmAjusteInc.grdItemAjuste_UnboundReadDataEx]"
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  Dim objAjuste     As busSisMetal.clsAjuste
  '
  blnFechar = False 'Não Pode Fechar pelo X
  blnRetorno = False
  blnPrimeiraVez = True
  blnAlterouPeso = False
  lngLINHAID = 0
  '
  AmpS
  Me.Height = 7620
  Me.Width = 11610
  CenterForm Me
  '
  LimparCampos
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdFechar
  '
  'Ajuste
  strSql = "SELECT TIPO_AJUSTE.DESCRICAO FROM TIPO_AJUSTE " & _
      "ORDER BY TIPO_AJUSTE.DESCRICAO"
  PreencheCombo cboAjuste, strSql, False, True
  '
  '--- Selecionar inventário
  INCLUIR_VALOR_NO_COMBO "INVENTÁRIO", cboAjuste
  cboAjuste.Enabled = False
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Ajuste
    '
    fraFiltro.Enabled = True
    fraAjuste.Enabled = True
    cmdOk.Enabled = True
    grdItemAjuste.Enabled = True
    txtUsuario.Text = gsNomeUsu
    'No evento de inclusão deve ser habilitado a coluna peso
    'grdItemAjuste.Columns(8).Locked = False
    'grdItemAjuste.Columns(9).Visible = False
    'grdItemAjuste.Columns(10).Visible = False
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objAjuste = New busSisMetal.clsAjuste
    Set objRs = objAjuste.ListarAjuste(lngAJUSTEID)
    '
    If Not objRs.EOF Then
      'Campos fixos
      txtUsuario = objRs.Fields("USUARIO").Value & ""
      INCLUIR_VALOR_NO_MASK mskData(0), objRs.Fields("DATA").Value, TpMaskData
      'Campos inserts
      INCLUIR_VALOR_NO_COMBO objRs.Fields("DESCR_AJUSTE").Value & "", cboAjuste
    End If
    Set objAjuste = Nothing
    '
    fraFiltro.Enabled = False
    If Status = tpStatus_Alterar Then
      fraAjuste.Enabled = True
      cmdOk.Enabled = True
      grdItemAjuste.Enabled = True
    Else
      fraAjuste.Enabled = False
      cmdOk.Enabled = False
      grdItemAjuste.Enabled = False
    End If

    'No evento de alteração deve ser habilitado as colunas anod e fabrica
    'If gsNivel <> gsCompra Then
    '  grdItemAjuste.Columns(8).Locked = False
    '  grdItemAjuste.Columns(9).Visible = False
    '  grdItemAjuste.Columns(10).Visible = False
    'Else
    '  grdItemAjuste.Columns(8).Locked = True
    '  grdItemAjuste.Columns(9).Visible = True
    '  grdItemAjuste.Columns(10).Visible = True
    'End If
  End If
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
  Dim objItemAjuste         As busSisMetal.clsItemAjuste
  '
  Dim lngQtdIni             As Long
  Dim lngQtdAnod            As Long
  Dim lngQtdFab             As Long

  '
  blnSetarFocoControle = True
  '
  strMsg = ""
  'Validção dos ítens do pedido
  If Not Valida_Moeda(grdItemAjuste.Columns("Quantidade"), TpnaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
    strMsg = strMsg & "Quantidade inválida na linha " & intLinha + 1 & vbCrLf
  End If
''''  '
''''  If Len(strMsg) = 0 Then
''''    'NOVO - Validações de cálculo de peças (quantidade)
''''    Set objItemAjuste = New busSisMetal.clsItemAjuste
''''    lngQtdIni = objItemAjuste.CalculoQuantidadeAjuste(grdItemAjuste.Columns("LINHAID").Value, _
''''                                                      grdItemAjuste.Columns("Peso").Value)
''''    If lngQtdIni = 0 Then
''''      strMsg = strMsg & "A quantidade calculada para o perfil deve ser maior que zero na linha " & intLinha + 1 & vbCrLf
''''    End If
''''    If grdItemAjuste.Columns("Anod.").Value & "" <> "" Then
''''      'Lançou peso para anodização
''''      lngQtdAnod = objItemAjuste.CalculoQuantidadeAjuste(grdItemAjuste.Columns("LINHAID").Value, _
''''                                                        grdItemAjuste.Columns("Anod.").Value)
''''      If lngQtdAnod = 0 Then
''''        strMsg = strMsg & "A quantidade calculada para anodização para o perfil deve ser maior que zero na linha " & intLinha + 1 & vbCrLf
''''      End If
''''    End If
''''    If grdItemAjuste.Columns("Fábrica").Value & "" <> "" Then
''''      'Lançou peso para anodização
''''      lngQtdFab = objItemAjuste.CalculoQuantidadeAjuste(grdItemAjuste.Columns("LINHAID").Value, _
''''                                                        grdItemAjuste.Columns("Fábrica").Value)
''''      If lngQtdFab = 0 Then
''''        strMsg = strMsg & "A quantidade calculada para fábrica para o perfil deve ser maior que zero na linha " & intLinha + 1 & vbCrLf
''''      End If
''''    End If
''''
''''
'''''''    lngTotal = 0
'''''''    lngTotalANOD = 0
'''''''    '
'''''''    lngTotal = CLng(grdItemAjuste.Columns("Qtd. Total")) - CLng(grdItemAjuste.Columns("Qtd. Baixa"))
'''''''    lngTotalANOD = CLng(IIf(Not IsNumeric(grdItemAjuste.Columns("Quantidade")), 0, grdItemAjuste.Columns("Quantidade")))
''''    '
''''    Set objItemAjuste = Nothing
''''  End If

  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmAjusteInc.ValidaCamposAnodOrigemLinha]"
    ValidaCamposAnodOrigemLinha = False
  Else
    ValidaCamposAnodOrigemLinha = True
  End If
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmAjusteInc.ValidaCamposAnodOrigemLinha]"
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
  
  
  For intRows = 0 To ITEMAJU_LINHASMATRIZ - 1
    grdItemAjuste.Bookmark = CLng(intRows)
    '
    If grdItemAjuste.Columns("*").Text & "" = "-1" Then
      'Somente válida se preencheu algo, sneão considera ok
      If grdItemAjuste.Columns("Quantidade").Text & "" <> "" Then
        If Not ValidaCamposAnodOrigemLinha(grdItemAjuste.Row) Then
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
    'não ouve erro
    If Status = tpStatus_Incluir Then
      TratarErroPrevisto "Selecione no mínimo 1 item para cadastro", "[frmAjusteInc.ValidaCamposAnodOrigemAll]"
    Else
      'No caso da alteração não é obrigatório cadastrar ou alterar ítem
      'pode estar apenas alterando dados do pedido
      blnRetorno = False
    End If
    
  End If
  grdItemAjuste.ReBind
  grdItemAjuste.SetFocus
  ValidaCamposAnodOrigemAll = blnRetorno
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserAjusteInc.ValidaCamposAnodOrigemAll]"
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
    If Not Valida_Moeda(grdItemAjuste.Columns(5), TpnaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
      strMsg = strMsg & "Quantidade de perfil branco inválida" & vbCrLf
    End If
  Case 6
    'Validção da quantidade brilho
    If Not Valida_Moeda(grdItemAjuste.Columns(6), TpnaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
      strMsg = strMsg & "Quantidade de perfil brilho inválida" & vbCrLf
    End If
  Case 7
    'Validção da quantidade bronze
    If Not Valida_Moeda(grdItemAjuste.Columns(7), TpnaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
      strMsg = strMsg & "Quantidade de perfil bronze inválida" & vbCrLf
    End If
  Case 8
    'Validção da quantidade natural
    If Not Valida_Moeda(grdItemAjuste.Columns(8), TpnaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
      strMsg = strMsg & "Quantidade de perfil natural inválida" & vbCrLf
    End If
  End Select
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmAjusteInc.ValidaCamposAnodOrigem]"
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
  Dim intUltimaColuna   As Integer
  If Me.ActiveControl.Name <> "grdItemAjuste" Then
    If KeyAscii = 13 Then
      SendKeys "{tab}"
    End If
  Else
    'If Status = tpStatus_Incluir Or gsNivel <> gsCompra Then
      intUltimaColuna = 8
    'Else
    '  intUltimaColuna = 10
    'End If
    If KeyAscii = 13 And IsNumeric(grdItemAjuste.Columns("ROWNUM").Value & "") = True Then
      If grdItemAjuste.Col >= intUltimaColuna Then
        If grdItemAjuste.Columns("ROWNUM").Value + 1 = ITEMAJU_LINHASMATRIZ Then
          cmdOK_Click
        Else
          grdItemAjuste.Col = intUltimaColuna
          '
          grdItemAjuste.MoveNext
        End If
      Else
        grdItemAjuste.Col = grdItemAjuste.Col + 1
      End If
    ElseIf (KeyAscii = 8) Or (KeyAscii = 44) Then
    ElseIf (KeyAscii < 48 Or KeyAscii > 57) Then
      KeyAscii = 0
    End If
  End If
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmAjusteInc.Form_KeyPress]"
End Sub



Private Sub txtCodigo_LostFocus()
  On Error GoTo trata
  Dim objLinhaCons    As Form
  Dim objLinhaPerfil  As busSisMetal.clsLinhaPerfil
  Dim objRs           As ADODB.Recordset
  Dim blnSelecionou   As Boolean
  If Me.ActiveControl.Name = "cmdFechar" Then Exit Sub
  If Me.ActiveControl.Name = "cmdOk" Then Exit Sub

  Pintar_Controle txtCodigo, tpCorContr_Normal
  If Len(txtCodigo.Text) = 0 Then
    If Len(txtCodigoFim.Text) <> 0 And Len(txtLinhaFim.Text) <> 0 Then
      lngLINHAID = 0
      '
      ITEMAJU_MontaMatriz (lngLINHAID)
      grdItemAjuste.Bookmark = Null
      grdItemAjuste.ReBind
      grdItemAjuste.ApproxCount = ITEMAJU_LINHASMATRIZ
      '
      Exit Sub
    Else
      'TratarErroPrevisto "Entre com o código ou descrição da linha."
      'Pintar_Controle txtCodigo, tpCorContr_Erro
      'SetarFoco txtCodigo
      'Exit Sub
      lngLINHAID = 0
      '
      ITEMAJU_MontaMatriz (lngLINHAID)
      grdItemAjuste.Bookmark = Null
      grdItemAjuste.ReBind
      grdItemAjuste.ApproxCount = ITEMAJU_LINHASMATRIZ
      '
      Exit Sub
      
    End If
  End If
  blnSelecionou = False
  Set objLinhaPerfil = New busSisMetal.clsLinhaPerfil
  '
  Set objRs = objLinhaPerfil.CapturaItemLinha(txtCodigo.Text)
  If objRs.EOF Then
    LimparCampoTexto txtCodigoFim
    LimparCampoTexto txtLinhaFim
    TratarErroPrevisto "Descrição/Código da linha não cadastrado"
    Pintar_Controle txtCodigo, tpCorContr_Erro
    SetarFoco txtCodigo
    Exit Sub
  Else
    If objRs.RecordCount = 1 Then
      txtCodigoFim.Text = objRs.Fields("CODIGO").Value & ""
      txtLinhaFim.Text = objRs.Fields("NOME").Value & ""
      lngLINHAID = objRs.Fields("PKID").Value & ""
      blnSelecionou = True
    Else
      'Novo : apresentar tela para seleção da linha
      Set objLinhaCons = New frmLinhaCons
      objLinhaCons.intIcOrigemLn = 6
      objLinhaCons.strCodigoDescricao = txtCodigo.Text
      objLinhaCons.Show vbModal
      blnSelecionou = True
      Set objLinhaCons = Nothing
    End If
    '
    objRs.Close
    Set objRs = Nothing
    Set objLinhaPerfil = Nothing
  End If
  ''''recarregar GRID
  If blnSelecionou = True Then
    '
    ITEMAJU_MontaMatriz (lngLINHAID)
    grdItemAjuste.Bookmark = Null
    grdItemAjuste.ReBind
    grdItemAjuste.ApproxCount = ITEMAJU_LINHASMATRIZ
  End If
  
  
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  KeyAscii = TRANSFORMA_MAIUSCULA(KeyAscii)
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtCodigo_GotFocus()
  Seleciona_Conteudo_Controle txtCodigo
End Sub

