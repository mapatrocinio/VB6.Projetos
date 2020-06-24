VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSaidaMaterialInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de saída de material"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5610
      Left            =   8520
      ScaleHeight     =   5610
      ScaleWidth      =   1860
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4695
         Left            =   30
         ScaleHeight     =   4635
         ScaleWidth      =   1605
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   720
         Width           =   1665
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&X"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdPedido 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   3600
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2730
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5295
      Left            =   120
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da saída de material"
      TabPicture(0)   =   "userSaidaMaterialInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraProf"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Inclusão de itens"
      TabPicture(1)   =   "userSaidaMaterialInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Composição dos Itens da saída de material"
      TabPicture(2)   =   "userSaidaMaterialInc.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdGeral"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraProf 
         Height          =   4335
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   7935
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   3975
            Index           =   0
            Left            =   120
            ScaleHeight     =   3975
            ScaleWidth      =   7695
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   240
            Width           =   7695
            Begin VB.Frame Frame3 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Cadastro"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3615
               Left            =   120
               TabIndex        =   25
               Top             =   360
               Width           =   7575
               Begin VB.TextBox txtMotivo 
                  Height          =   285
                  Left            =   1320
                  MaxLength       =   100
                  TabIndex        =   5
                  Text            =   "txtDescricao"
                  Top             =   2670
                  Width           =   6135
               End
               Begin VB.ComboBox cboFilial 
                  Height          =   315
                  Left            =   1320
                  Style           =   2  'Dropdown List
                  TabIndex        =   1
                  Top             =   1020
                  Width           =   3975
               End
               Begin VB.ComboBox cboDocumento 
                  Height          =   315
                  Left            =   1320
                  Style           =   2  'Dropdown List
                  TabIndex        =   0
                  Top             =   660
                  Width           =   3975
               End
               Begin VB.TextBox txtDescricao 
                  Height          =   765
                  Left            =   1320
                  MaxLength       =   250
                  MultiLine       =   -1  'True
                  TabIndex        =   2
                  Text            =   "userSaidaMaterialInc.frx":0054
                  Top             =   1380
                  Width           =   6135
               End
               Begin VB.PictureBox Picture2 
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   120
                  ScaleHeight     =   255
                  ScaleWidth      =   3855
                  TabIndex        =   26
                  TabStop         =   0   'False
                  Top             =   360
                  Width           =   3855
                  Begin MSMask.MaskEdBox mskData 
                     Height          =   255
                     Index           =   0
                     Left            =   1200
                     TabIndex        =   27
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
                     Left            =   0
                     TabIndex        =   28
                     Top             =   0
                     Width           =   615
                  End
               End
               Begin MSMask.MaskEdBox mskData 
                  Height          =   255
                  Index           =   1
                  Left            =   1320
                  TabIndex        =   3
                  Top             =   2220
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   450
                  _Version        =   393216
                  BackColor       =   16777215
                  AutoTab         =   -1  'True
                  MaxLength       =   10
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskData 
                  Height          =   255
                  Index           =   2
                  Left            =   3930
                  TabIndex        =   4
                  Top             =   2220
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   450
                  _Version        =   393216
                  BackColor       =   16777215
                  AutoTab         =   -1  'True
                  MaxLength       =   10
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin VB.Label Label6 
                  Caption         =   "Motivo"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   35
                  Top             =   2670
                  Width           =   1095
               End
               Begin VB.Label Label3 
                  Caption         =   "Data de Transação"
                  Height          =   375
                  Left            =   2730
                  TabIndex        =   34
                  Top             =   2220
                  Width           =   855
               End
               Begin VB.Label Label6 
                  Caption         =   "Descrição"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   32
                  Top             =   1380
                  Width           =   1095
               End
               Begin VB.Label Label6 
                  Caption         =   "Filial"
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   31
                  Top             =   1020
                  Width           =   1215
               End
               Begin VB.Label Label6 
                  Caption         =   "Documento"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   30
                  Top             =   660
                  Width           =   1215
               End
               Begin VB.Label Label5 
                  Caption         =   "Data de Aquisição"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   29
                  Top             =   2220
                  Width           =   855
               End
            End
            Begin VB.TextBox txtCodigo 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   2040
               Locked          =   -1  'True
               MaxLength       =   5
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   0
               Width           =   1455
            End
            Begin VB.Label Label44 
               Caption         =   "Sequencial"
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
               Left            =   120
               TabIndex        =   33
               Top             =   0
               Width           =   1935
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2715
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   6945
         Begin VB.TextBox txtCodigoProduto 
            Height          =   285
            Left            =   1560
            MaxLength       =   15
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   690
            Width           =   1455
         End
         Begin VB.TextBox txtProduto 
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1080
            Width           =   5205
         End
         Begin MSMask.MaskEdBox mskQuantidade 
            Height          =   255
            Left            =   1560
            TabIndex        =   11
            Top             =   360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,###;($#,###)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Produto"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   20
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Quantidade"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Código"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   18
            Top             =   720
            Width           =   1095
         End
      End
      Begin TrueDBGrid60.TDBGrid grdGeral 
         Height          =   3795
         Left            =   -74760
         OleObjectBlob   =   "userSaidaMaterialInc.frx":0061
         TabIndex        =   21
         Top             =   480
         Width           =   7335
      End
   End
End
Attribute VB_Name = "frmSaidaMaterialInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                 As tpStatus
Public lngSAIDAMATERIALID     As Long
Public blnRetorno             As Boolean
Public blnFechar              As Boolean
Public intQuemChamou          As Integer
Private blnPrimeiraVez        As Boolean

'Variáveis para Grid ObraEngenheiro
Dim GER_COLUNASMATRIZ         As Long
Dim GER_LINHASMATRIZ          As Long
Private GER_Matriz()          As String

Public Sub GER_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMetal.clsGeral
  '
  On Error GoTo trata

  Set clsGer = New busSisMetal.clsGeral
  '
  strSql = "SELECT TAB_SAIDAMATERIAL.QUANTIDADE, INSUMO.CODIGO, PRODUTO.NOME, TAB_SAIDAMATERIAL.PKID, TAB_SAIDAMATERIAL.SAIDAMATERIALID, TAB_SAIDAMATERIAL.PRODUTOID " & _
          "FROM TAB_SAIDAMATERIAL INNER JOIN  PRODUTO ON TAB_SAIDAMATERIAL.PRODUTOID =  PRODUTO.INSUMOID " & _
          "INNER JOIN INSUMO ON PRODUTO.INSUMOID =  INSUMO.PKID " & _
          "WHERE TAB_SAIDAMATERIAL.SAIDAMATERIALID = " & lngSAIDAMATERIALID & _
          " ORDER BY TAB_SAIDAMATERIAL.PKID"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    GER_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim GER_Matriz(0 To GER_COLUNASMATRIZ - 1, 0 To GER_LINHASMATRIZ - 1)
  Else
    ReDim GER_Matriz(0 To GER_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To GER_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To GER_COLUNASMATRIZ - 1  'varre as colunas
          GER_Matriz(intJ, intI) = objRs(intJ) & ""
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


Private Sub cboDocumento_Click()
  On Error GoTo trata
  If UCase(cboDocumento.Text & "") = "TRANSFERÊNCIA DE SAÍDA" Or _
    UCase(cboDocumento.Text & "") = "TRANSFERÊNCIA DE ENTRADA" Then
    Label6(3).Enabled = True
    cboFilial.Enabled = True
  Else
    cboFilial.ListIndex = -1
    Label6(3).Enabled = False
    cboFilial.Enabled = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmSaidaMaterialInc.cboDocumento_Click]"
End Sub

Private Sub cboDocumento_LostFocus()
  Pintar_Controle cboDocumento, tpCorContr_Normal
End Sub

Private Sub cmdExcluir_Click()
  '
  Dim objSaiMat           As busSisMetal.clsSaidaMaterial
  Dim lngNOVAQUANTIDADE   As Long
  Dim strMsgErro          As String
  '
  On Error GoTo trata
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração do cardápio
  Case 2
    If grdGeral.Columns(1).Value = "" Then
      MsgBox "Selecione um item da saída de material para excluí-lo.", vbExclamation, TITULOSISTEMA
      Exit Sub
    End If
    If MsgBox("Deseja excluir o item da saída de material " & grdGeral.Columns("Produto").Value & " ?", vbYesNo, TITULOSISTEMA) = vbYes Then
      Set objSaiMat = New busSisMetal.clsSaidaMaterial

      'Validações avançadas - CAPTURA TB A NOVA QUANTIDADE
      objSaiMat.ValidarExclusaoTab_SaidaMaterial strMsgErro, _
                                                 grdGeral.Columns("Código").Value, _
                                                 grdGeral.Columns("Qtd.").Value, _
                                                 lngNOVAQUANTIDADE, _
                                                 cboDocumento.Text
      If Len(Trim(strMsgErro)) <> 0 Then
        Set objSaiMat = Nothing
        TratarErroPrevisto strMsgErro, "[frmSaidaMaterialInc.ValidarExclusaoItemSaidaMaterial]"
        SetarFoco grdGeral
        Exit Sub
      End If

      '
      objSaiMat.ExcluirTAB_SAIDAMATERIAL CLng(grdGeral.Columns("PKID").Value), _
                                         CLng(grdGeral.Columns("PRODUTOID").Value), _
                                         lngNOVAQUANTIDADE


      '
      Set objSaiMat = Nothing
      blnRetorno = True
      'Montar RecordSet
      GER_COLUNASMATRIZ = grdGeral.Columns.Count
      GER_LINHASMATRIZ = 0
      GER_MontaMatriz
      grdGeral.Bookmark = Null
      grdGeral.ReBind

      SetarFoco grdGeral

    End If
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmSaidaMaterialInc.cmdExcluir_Click]"
End Sub


Private Sub cboFilial_LostFocus()
  Pintar_Controle cboFilial, tpCorContr_Normal
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
               Offset + intI, GER_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, GER_COLUNASMATRIZ, GER_LINHASMATRIZ, GER_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, GER_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmSaidaMaterialInc.grdGeral_UnboundReadDataEx]"
End Sub


Private Sub cmdCancelar_Click()
  On Error GoTo trata
  blnFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim strSql                    As String
  Dim objRs                     As ADODB.Recordset
  Dim objSaiMat                 As busSisMetal.clsSaidaMaterial
  Dim objGeral                  As busSisMetal.clsGeral
  '
  Dim lngDOCUMENTOSAIDAID       As Long
  Dim lngPRODUTOID              As Long
  Dim lngFILAILID               As Long
  Dim lngQUANTIDADE             As Long
  Dim strData                   As String
  Dim strCodigo                 As String
  '
  Dim strMsgErro                As String
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração Saida de material
    If Not ValidaCamposSaidaMaterial Then Exit Sub
    Set objSaiMat = New busSisMetal.clsSaidaMaterial
    Set objGeral = New busSisMetal.clsGeral
    '
    'Obter campos
    'lngDOCUMENTOSAIDAID
    Set objRs = objSaiMat.ListarDocumentoSaidaPelaDesc(cboDocumento.Text)
    If objRs.EOF Then
      lngDOCUMENTOSAIDAID = 0
    Else
      lngDOCUMENTOSAIDAID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    '
    'lngFILAILID
    lngFILAILID = 0
    strSql = "SELECT PKID FROM FILIAL INNER JOIN LOJA ON LOJA.PKID = FILIAL.LOJAID " & _
          " WHERE LOJA.NOME = " & Formata_Dados(cboFilial.Text, tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngFILAILID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    '
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objSaiMat.AlterarSaidaMaterial lngSAIDAMATERIALID, _
                                     txtDescricao.Text, _
                                     lngDOCUMENTOSAIDAID, _
                                     lngFILAILID, _
                                      mskData(1).Text, _
                                      mskData(2).Text, _
                                      txtMotivo.Text

    ElseIf Status = tpStatus_Incluir Then
      '
      objSaiMat.InserirSaidaMaterial txtDescricao.Text, _
                                     gsNomeUsu, _
                                     lngDOCUMENTOSAIDAID, _
                                     lngFILAILID, _
                                     mskData(1).Text, _
                                     lngSAIDAMATERIALID, _
                                     strData, _
                                     strCodigo, _
                                      mskData(2).Text, _
                                      txtMotivo.Text
      '
      txtCodigo.Text = Format(strCodigo, "0000")
      INCLUIR_VALOR_NO_MASK mskData(0), strData, TpMaskData
      '
      Status = tpStatus_Alterar
      '
      tabDetalhes.TabEnabled(1) = True
      tabDetalhes.TabEnabled(2) = True
      '
      tabDetalhes.Tab = 1
    End If
    Set objSaiMat = Nothing
    blnRetorno = True
    '
    tabDetalhes.TabEnabled(1) = True
    tabDetalhes.TabEnabled(2) = True
    '
    tabDetalhes.Tab = 1
    'blnFechar = True
    'Unload Me
  Case 1 'Itens da entrada de material
    'Código para Inclusão
    'Validações Básicas
    If Not ValidaCamposItens Then Exit Sub
    Set objSaiMat = New busSisMetal.clsSaidaMaterial
    'Validações avançadas - APENAS P/ CAPTURAR ESTOQUEID E QUANTIDADE
    
    objSaiMat.ValidarInclusaoProduto lngPRODUTOID, _
                                     lngSAIDAMATERIALID, _
                                     strMsgErro, _
                                     txtCodigoProduto.Text, _
                                     lngQUANTIDADE, _
                                     mskQuantidade.Text, _
                                     cboDocumento.Text
    If Len(Trim(strMsgErro)) <> 0 Then
      Set objSaiMat = Nothing
      TratarErroPrevisto strMsgErro, "[frmSaidaMaterialInc.ValidarExclusaoItemSaidaMaterial]"
      Exit Sub
    End If
    'Alterar PRODUTO
    objSaiMat.AlterarProdutoPelaSaidaMaterial lngPRODUTOID, _
                                              lngQUANTIDADE, _
                                              mskQuantidade.Text, _
                                              cboDocumento.Text
    'InserirTAB_SAIDAMATERIAL
    objSaiMat.InserirTAB_SAIDAMATERIAL lngSAIDAMATERIALID, _
                                       lngPRODUTOID, _
                                       mskQuantidade.Text


    'Limpa campos para Próxima inserção
    TratarErroPrevisto "Item do produto cadastrado com sucesso.", "[frmEntradaMaterialInc.cmdOk_Click]"
    
    mskQuantidade.Text = ""
    txtCodigoProduto.Text = ""
    '
    mskQuantidade.SetFocus
    Set objSaiMat = Nothing
    blnRetorno = True
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCamposItens() As Boolean
  Dim strMsg              As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  '
  If Not Valida_Moeda(mskQuantidade, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "A Quantidade do produto é inválida" & vbCrLf
    Pintar_Controle mskQuantidade, tpCorContr_Erro
  ElseIf CLng(mskQuantidade.Text) = 0 Then
    strMsg = strMsg & "A Quantidade do produto não pode ser igual a zero" & vbCrLf
    Pintar_Controle mskQuantidade, tpCorContr_Erro
    blnSetarFocoControle = True
    SetarFoco mskQuantidade
  End If
  If Len(txtProduto.Text) = 0 Then
    strMsg = strMsg & "Entrar com o produto" & vbCrLf
    Pintar_Controle txtCodigoProduto, tpCorContr_Erro
    blnSetarFocoControle = True
    SetarFoco txtCodigoProduto
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmSaidaMaterialInc.ValidaCamposItens]"
    ValidaCamposItens = False
  Else
    ValidaCamposItens = True
  End If
End Function



Private Function ValidaCamposSaidaMaterial() As Boolean
  Dim strMsg              As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  '
  If Not Valida_String(cboDocumento, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o documento" & vbCrLf
  End If
  If cboDocumento.Text = "TRANSFERÊNCIA DE SAÍDA" Or _
    cboDocumento.Text = "TRANSFERÊNCIA DE ENTRADA" Then
    If Not Valida_String(cboFilial, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Selecionar a filial" & vbCrLf
    End If
  End If
  If Not Valida_String(txtDescricao, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Informar a descrição válida" & vbCrLf
  End If
  If Not Valida_Data(mskData(1), TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Informar a Data de Aquisição do Documento válida" & vbCrLf
  End If
  If Not Valida_Data(mskData(2), TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Informar a Data de Transação do Documento válida" & vbCrLf
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmSaidaMaterialInc.ValidaCamposSaidaMaterial]"
    ValidaCamposSaidaMaterial = False
  Else
    ValidaCamposSaidaMaterial = True
  End If
End Function


Private Sub cmdPedido_Click()
  On Error GoTo trata
  frmProdutoCons.QuemChamou = 1
  frmProdutoCons.strCodigoProduto = txtCodigoProduto.Text
  frmProdutoCons.Show vbModal

  SetarFoco txtCodigoProduto
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    If Status = tpStatus_Incluir Then
      tabDetalhes.Tab = 0
      'txtNome.SetFocus
    Else
      tabDetalhes.Tab = 0
    End If
    blnPrimeiraVez = False

  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmSaidaMaterialInc.Form_Activate]"
End Sub




Private Sub mskData_GotFocus(Index As Integer)
  Selecionar_Conteudo mskData(Index)
End Sub

Private Sub mskData_LostFocus(Index As Integer)
  Pintar_Controle mskData(Index), tpCorContr_Normal
End Sub

Private Sub mskQuantidade_Change()
  'If Len(Trim(mskQuantidade.ClipText)) = 2 Then SetarFoco txtCodigoProduto
End Sub

Private Sub txtCodigoProduto_Change()
  On Error GoTo trata
  Dim objProduto    As busSisMetal.clsInsumo
  Dim objRs     As ADODB.Recordset
  'If Len(txtCodigoProduto.Text) < 3 Then
  If Len(txtCodigoProduto.Text) = 0 Then
    txtProduto.Text = ""
    Exit Sub
  End If
  Set objProduto = New busSisMetal.clsInsumo
  '
  Set objRs = objProduto.SelecionarProdutoPeloCodigo(txtCodigoProduto.Text)
  If objRs.EOF Then
    txtProduto.Text = ""
  Else
    If objRs.RecordCount = 1 Then
      txtProduto.Text = objRs.Fields("NOME").Value
      SetarFoco cmdOk
    Else
      txtProduto.Text = ""
    End If
  End If
  '
  objRs.Close
  Set objRs = Nothing
  Set objProduto = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtCodigoProduto_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  KeyAscii = TRANSFORMA_MAIUSCULA(KeyAscii)
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub


Private Sub txtCodigoProduto_LostFocus()
  Dim objProdutoCons  As SisMetal.frmProdutoCons
  'Dim objProduto      As busSisMetal.clsInsumo
  Dim objGeral        As busSisMetal.clsGeral
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  On Error GoTo trata
  txtCodigoProduto.Text = UCase(txtCodigoProduto.Text)
  Pintar_Controle txtCodigoProduto, tpCorContr_Normal
  If Me.ActiveControl.Name = "cmdCancelar" Then Exit Sub
  If Me.ActiveControl.Name = "grdGeral" Then Exit Sub
  If Me.ActiveControl.Name = "cmdPedido" Then Exit Sub
  If Len(txtCodigoProduto.Text) = 0 Then
    TratarErroPrevisto "Entre com o código do produto."
    Pintar_Controle txtCodigoProduto, tpCorContr_Erro
    SetarFoco txtCodigoProduto
    Exit Sub
  End If
  Set objGeral = New busSisMetal.clsGeral
  '
  strSql = "SELECT INSUMO.CODIGO, PRODUTO.NOME, PRODUTO.SALDOESTOQUE " & _
    " FROM PRODUTO INNER JOIN INSUMO ON INSUMO.PKID = PRODUTO.INSUMOID " & _
    "WHERE (NOME LIKE '%" & txtCodigoProduto.Text & "%' " & _
        " OR CODIGO LIKE '%" & txtCodigoProduto.Text & "%') " & _
   " ORDER BY PRODUTO.NOME;"
  Set objRs = objGeral.ExecutarSQL(strSql)
  If objRs.EOF Then
    TratarErroPrevisto "Entre com o código do produto."
    Pintar_Controle txtCodigoProduto, tpCorContr_Erro
    SetarFoco txtCodigoProduto
    Exit Sub
  Else
    If objRs.RecordCount = 1 Then
      txtProduto.Text = objRs.Fields("NOME").Value
      If Len(mskQuantidade.ClipText) = 0 Then
        SetarFoco mskQuantidade
      ElseIf Len(txtCodigoProduto.Text & "") = 0 Then
        SetarFoco txtCodigoProduto
      Else
        SetarFoco cmdOk
      End If
    Else
      'Novo : apresentar tela para seleção do produto
      Set objProdutoCons = New frmProdutoCons
      objProdutoCons.QuemChamou = 1
      objProdutoCons.strCodigoProduto = txtCodigoProduto.Text
      objProdutoCons.Show vbModal
      
      If txtProduto.Text = "" Then
        txtProduto.Text = ""
        TratarErroPrevisto "Selecione um produto."
        Pintar_Controle txtCodigoProduto, tpCorContr_Erro
        SetarFoco txtCodigoProduto
        Exit Sub
      Else
        If Len(mskQuantidade.ClipText) = 0 Then
          SetarFoco mskQuantidade
        ElseIf Len(txtCodigoProduto.Text & "") = 0 Then
          SetarFoco txtCodigoProduto
        Else
          SetarFoco cmdOk
        End If
        
      End If
      Set objProdutoCons = Nothing
    End If
  End If
  '
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  'cmdOk.Default = True
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Load()
On Error GoTo trata
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  Dim objSaiMat     As busSisMetal.clsSaidaMaterial
  '
  blnFechar = False
  blnRetorno = False
  AmpS
  Me.Height = 5985
  Me.Width = 10470
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, cmdExcluir, , , , , cmdImprimir
  LerFigurasAvulsas cmdPedido, "Cardapio.ico", "CardapioDown.ico", "Visualização de itens do cardápio"
  'LIMPAR CAMPOS
  LimparCampoTexto txtCodigo
  'INCLUIR_VALOR_NO_MASK mskData(0), "", TpMaskData
  LimparCampoTexto txtDescricao
  'INCLUIR_VALOR_NO_MASK mskData(1), "", TpMaskData
  LimparCampoCombo cboDocumento
  LimparCampoCombo cboFilial
  LimparCampoMask mskData(1)
  LimparCampoMask mskData(2)
  LimparCampoTexto txtMotivo
  '
  'DOCUMENTO DE SAIDA
  strSql = "Select NOME from DOCUMENTOSAIDA ORDER BY NOME"
  PreencheCombo cboDocumento, strSql, False, True
  'FILAIL
  strSql = "Select LOJA.NOME FROM FILIAL INNER JOIN LOJA ON LOJA.PKID = FILIAL.LOJAID ORDER BY LOJA.NOME"
  PreencheCombo cboFilial, strSql, False, True
  '
  tabDetalhes_Click 0
  '
  If Status = tpStatus_Incluir Then
    '
    tabDetalhes.TabEnabled(1) = False
    tabDetalhes.TabEnabled(2) = False
    cboDocumento.Enabled = True
    Label6(0).Enabled = True
  ElseIf Status = tpStatus_Alterar Then
    'Pega Dados do Banco de dados
    Set objSaiMat = New busSisMetal.clsSaidaMaterial
    Set objRs = objSaiMat.ListarSaidaMaterial(lngSAIDAMATERIALID)
    '
    If Not objRs.EOF Then
      '
      INCLUIR_VALOR_NO_MASK mskData(0), objRs.Fields("DATA").Value, TpMaskData
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
      txtCodigo.Text = Format(IIf(Not IsNumeric(objRs.Fields("CODIGO").Value), 0, objRs.Fields("CODIGO").Value), "0000")
      If Not IsNull(objRs.Fields("DOCSAIDA").Value) Then cboDocumento.Text = objRs.Fields("DOCSAIDA").Value
      INCLUIR_VALOR_NO_COMBO objRs.Fields("DOCSAIDA").Value, cboDocumento
      If Not IsNull(objRs.Fields("NOME_FILIAL").Value) Then cboFilial.Text = objRs.Fields("NOME_FILIAL").Value
      INCLUIR_VALOR_NO_MASK mskData(1), objRs.Fields("DATAAQUISICAO").Value, TpMaskData
      INCLUIR_VALOR_NO_MASK mskData(2), objRs.Fields("DATA_TRANSACAO").Value, TpMaskData
      txtMotivo.Text = objRs.Fields("MOTIVO").Value & ""
      '
    End If
    objRs.Close
    Set objRs = Nothing
    Set objSaiMat = Nothing
    '
    cboDocumento.Enabled = False
    Label6(0).Enabled = False
    tabDetalhes.TabEnabled(1) = True
    tabDetalhes.TabEnabled(2) = True
  End If
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then Cancel = True
End Sub


Private Sub tabDetalhes_Click(PreviousTab As Integer)
  Dim strMsgErro    As String
  Dim strCobranca   As String
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    'dados principais da venda
    picTrava(0).Enabled = True
    Frame1.Enabled = False
    grdGeral.Enabled = False
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdPedido.Enabled = False
    cmdExcluir.Enabled = False
    cmdImprimir.Enabled = False
    SetarFoco cboDocumento
  Case 1
    picTrava(0).Enabled = False
    Frame1.Enabled = True
    grdGeral.Enabled = False
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdPedido.Enabled = True
    cmdExcluir.Enabled = False
    cmdImprimir.Enabled = False
    '
    SetarFoco mskQuantidade
  Case 2
    'Vizualização dos Itens do cardápio
    picTrava(0).Enabled = False
    Frame1.Enabled = False
    grdGeral.Enabled = True
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdPedido.Enabled = False
    cmdExcluir.Enabled = True
    cmdImprimir.Enabled = False
    'Montar RecordSet
    GER_COLUNASMATRIZ = grdGeral.Columns.Count
    GER_LINHASMATRIZ = 0
    GER_MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    SetarFoco grdGeral
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "SisMetal.frmSaidaMaterialInc.tabDetalhes"
  AmpN
End Sub


Private Sub mskQuantidade_GotFocus()
  Selecionar_Conteudo mskQuantidade
End Sub

Private Sub mskQuantidade_LostFocus()
  Pintar_Controle mskQuantidade, tpCorContr_Normal
End Sub


Private Sub txtCodigoProduto_GotFocus()
  Selecionar_Conteudo txtCodigoProduto
End Sub

Private Sub txtDescricao_GotFocus()
  Selecionar_Conteudo txtDescricao
End Sub

Private Sub txtDescricao_LostFocus()
  Pintar_Controle txtDescricao, tpCorContr_Normal
End Sub


Private Sub txtMotivo_GotFocus()
  Selecionar_Conteudo txtMotivo
End Sub

Private Sub txtMotivo_LostFocus()
  Pintar_Controle txtMotivo, tpCorContr_Normal
End Sub

