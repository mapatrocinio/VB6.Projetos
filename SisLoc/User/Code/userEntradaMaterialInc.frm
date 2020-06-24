VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserEntradaMaterialInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de entrada de material"
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4695
         Left            =   30
         ScaleHeight     =   4635
         ScaleWidth      =   1605
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   720
         Width           =   1665
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&X"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdPedido 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   3600
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
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
      Height          =   5295
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da entrada de material"
      TabPicture(0)   =   "userEntradaMaterialInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraProf"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Inclusão de itens"
      TabPicture(1)   =   "userEntradaMaterialInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Composição dos Itens da entrada de material"
      TabPicture(2)   =   "userEntradaMaterialInc.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdGeral"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame fraProf 
         Height          =   4335
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   7935
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   3975
            Index           =   0
            Left            =   120
            ScaleHeight     =   3975
            ScaleWidth      =   7695
            TabIndex        =   22
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
               TabIndex        =   24
               Top             =   360
               Width           =   7575
               Begin VB.ComboBox cboDocumento 
                  Height          =   315
                  Left            =   1320
                  Style           =   2  'Dropdown List
                  TabIndex        =   1
                  Top             =   1080
                  Width           =   3975
               End
               Begin VB.TextBox txtDescricao 
                  Height          =   765
                  Left            =   1320
                  MaxLength       =   250
                  MultiLine       =   -1  'True
                  TabIndex        =   3
                  Text            =   "userEntradaMaterialInc.frx":0054
                  Top             =   1800
                  Width           =   6135
               End
               Begin VB.TextBox txtFornecedor 
                  Height          =   285
                  Left            =   1320
                  MaxLength       =   50
                  TabIndex        =   2
                  Text            =   "txtFornecedor"
                  Top             =   1440
                  Width           =   6135
               End
               Begin VB.PictureBox Picture2 
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   120
                  ScaleHeight     =   255
                  ScaleWidth      =   3855
                  TabIndex        =   25
                  TabStop         =   0   'False
                  Top             =   360
                  Width           =   3855
                  Begin MSMask.MaskEdBox mskData 
                     Height          =   255
                     Index           =   0
                     Left            =   1200
                     TabIndex        =   26
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
                     TabIndex        =   27
                     Top             =   0
                     Width           =   615
                  End
               End
               Begin VB.TextBox txtNumero 
                  Height          =   285
                  Left            =   1320
                  MaxLength       =   15
                  MultiLine       =   -1  'True
                  TabIndex        =   0
                  Text            =   "userEntradaMaterialInc.frx":0061
                  Top             =   720
                  Width           =   1455
               End
               Begin MSMask.MaskEdBox mskData 
                  Height          =   255
                  Index           =   1
                  Left            =   1320
                  TabIndex        =   4
                  Top             =   2640
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
                  Caption         =   "Descrição"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   32
                  Top             =   1800
                  Width           =   1095
               End
               Begin VB.Label Label6 
                  Caption         =   "Fornecedor"
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   31
                  Top             =   1440
                  Width           =   1215
               End
               Begin VB.Label Label6 
                  Caption         =   "Documento"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   30
                  Top             =   1080
                  Width           =   1215
               End
               Begin VB.Label Label5 
                  Caption         =   "Data de Aquisição"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   29
                  Top             =   2640
                  Width           =   855
               End
               Begin VB.Label Label6 
                  Caption         =   "Número"
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   28
                  Top             =   720
                  Width           =   1095
               End
            End
            Begin VB.TextBox txtCodigo 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   2040
               Locked          =   -1  'True
               MaxLength       =   5
               TabIndex        =   23
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
         TabIndex        =   16
         Top             =   360
         Width           =   6945
         Begin VB.TextBox mskCodigo 
            Height          =   285
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   11
            Top             =   690
            Width           =   5175
         End
         Begin VB.TextBox txtProduto 
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1080
            Width           =   5205
         End
         Begin MSMask.MaskEdBox mskQuantidade 
            Height          =   255
            Left            =   1560
            TabIndex        =   10
            Top             =   360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0;($#,##0)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Produto"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   19
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Quantidade"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Código"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   720
            Width           =   1095
         End
      End
      Begin TrueDBGrid60.TDBGrid grdGeral 
         Height          =   3795
         Left            =   -74760
         OleObjectBlob   =   "userEntradaMaterialInc.frx":006D
         TabIndex        =   20
         Top             =   480
         Width           =   7335
      End
   End
End
Attribute VB_Name = "frmUserEntradaMaterialInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                 As tpStatus
Public lngENTRADAMATERIALID   As Long
Public bRetorno               As Boolean
Public bFechar                As Boolean
Public sTitulo                As String
Public strNumeroSuiteApto     As String
Public strNomeSuiteApto       As String
Public intQuemChamou          As Integer
Public strMotivo              As String
Public lngQtdRestanteProdReal As Long
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
  Dim clsGer    As busSisLoc.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisLoc.clsGeral
  '
  strSql = "SELECT ITEMENTRADA.QUANTIDADE, ESTOQUE.CODIGO, ESTOQUE.DESCRICAO, ITEMENTRADA.PKID, ITEMENTRADA.ENTRADAMATERIALID, ITEMENTRADA.ESTOQUEID " & _
          "FROM ITEMENTRADA INNER JOIN  ESTOQUE ON (ITEMENTRADA.ESTOQUEID =  ESTOQUE.PKID) " & _
          "WHERE ITEMENTRADA.ENTRADAMATERIALID = " & Formata_Dados(lngENTRADAMATERIALID, tpDados_Longo) & _
          " ORDER BY ITEMENTRADA.PKID"
          

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






Private Sub cboDocumento_LostFocus()
  cboDocumento.BackColor = vbWhite
End Sub

Private Sub cmdExcluir_Click()
  '
  Dim clsEntMat           As busSisLoc.clsEntradaMaterial
  Dim strMsgErro          As String
  '
  On Error GoTo trata
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração do cardápio
  Case 2
    If grdGeral.Columns(1).Value & "" = "" Then
      MsgBox "Selecione um item da entrada de material para exclui-la.", vbExclamation, TITULOSISTEMA
      Exit Sub
    End If
    If MsgBox("Deseja excluir o item da entrada de material " & grdGeral.Columns("Produto").Value & " ?", vbYesNo, TITULOSISTEMA) = vbYes Then
      Set clsEntMat = New busSisLoc.clsEntradaMaterial
      'Obter campos
      '
      clsEntMat.ExcluirITEMENTRADA CLng(grdGeral.Columns("PKID").Value), _
                                   CLng(grdGeral.Columns("ESTOQUEID").Value), _
                                   grdGeral.Columns("Qtd.").Value
      '
      Set clsEntMat = Nothing
      bRetorno = True
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
  TratarErro Err.Number, Err.Description, "[frmUserEntradaMaterialInc.cmdExcluir_Click]"
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
  TratarErro Err.Number, Err.Description, "[frmUserEntradaMaterialInc.grdGeral_UnboundReadDataEx]"
End Sub


Private Sub cmdCancelar_Click()
  '
  bFechar = True
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub



Private Sub cmdImprimir_Click()
  On Error GoTo trata
  'IMP_COMP_VENDA lngENTRADAMATERIALID, gsNomeEmpresa
  Exit Sub
  
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim strSql                    As String
  Dim objRs                     As ADODB.Recordset
  Dim objEntradaMaterial        As busSisLoc.clsEntradaMaterial
  Dim objDocumento              As busSisLoc.clsDocumento
  Dim objEstoque                As busSisLoc.clsEstoque
  '
  Dim lngDOCUMENTOENTRADAID     As Long
  Dim lngESTOQUEID              As Long
  Dim lngQUANTIDADE             As Long
  Dim strData                   As String
  Dim strCodigo                 As String
  '
  Dim lngESTOQUEINTERMEDIARIOID As Long
  Dim lngGRUPOESTQOUEID         As Long
  Dim lngQUANTIDADEESTINTER     As Long
  Dim strMsgErro                As String
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração Entrada de material
    If Not ValidaCamposEntradaMaterial Then Exit Sub
    Set objEntradaMaterial = New busSisLoc.clsEntradaMaterial
    '
    Set objDocumento = New busSisLoc.clsDocumento
    'Pegar campos
    '
    Set objRs = objDocumento.ListarDocumentoPelaDesc(cboDocumento.Text)
    If objRs.EOF Then
      lngDOCUMENTOENTRADAID = 0
    Else
      lngDOCUMENTOENTRADAID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    Set objDocumento = Nothing
    '
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objEntradaMaterial.AlterarEntradaMaterial lngENTRADAMATERIALID, _
                                       txtDescricao.Text, _
                                       lngDOCUMENTOENTRADAID, _
                                       txtFornecedor.Text, _
                                       txtNumero.Text, _
                                       mskData(1).Text

    ElseIf Status = tpStatus_Incluir Then
      '
      objEntradaMaterial.InserirEntradaMaterial txtDescricao.Text, _
                                       gsNomeUsu, _
                                       lngDOCUMENTOENTRADAID, _
                                       txtFornecedor.Text, _
                                       mskData(1).Text, _
                                       txtNumero.Text, _
                                       lngENTRADAMATERIALID, _
                                       strData, _
                                       strCodigo
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
    Set objEntradaMaterial = Nothing
    bRetorno = True
    '
    tabDetalhes.TabEnabled(1) = True
    tabDetalhes.TabEnabled(2) = True
    '
    tabDetalhes.Tab = 1
    'bFechar = True
    'Unload Me
  Case 1 'Itens da entrada de material
    'Código para Inclusão
    'Validações Básicas
    If Not ValidaCamposItens Then Exit Sub
    Set objEstoque = New busSisLoc.clsEstoque
    '
    'Obter campos
    '
    Set objRs = objEstoque.ListarEstoquePeloCodigo(mskCodigo.Text)
    If objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objEstoque = Nothing
      TratarErroPrevisto "Código do Produto não cadastrado no estoque", "[frmUserEntradaMaterialInc.ValidarExclusaoItemEntradaMaterial]"
      Exit Sub
    Else
      lngESTOQUEID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    Set objEstoque = Nothing
    '
    Set objEntradaMaterial = New busSisLoc.clsEntradaMaterial
    'Alterar ESTOQUE
    objEntradaMaterial.AlterarEstoquePelaEntradaMaterial lngESTOQUEID, _
                                                         mskCodigo.Text, _
                                                         mskQuantidade.Text
                                                
    'InserirITEMENTRADA
    objEntradaMaterial.InserirITEMENTRADA lngENTRADAMATERIALID, _
                                          lngESTOQUEID, _
                                          mskQuantidade.Text
                                         
                                         
    'Limpa campos para Próxima inserção
    mskQuantidade.Text = ""
    mskCodigo.Text = "     "
    '
    mskQuantidade.SetFocus
    Set objEntradaMaterial = Nothing
    bRetorno = True
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCamposItens() As Boolean
  Dim strMsg     As String
  
  If Not Valida_Moeda(mskQuantidade, TpObrigatorio) Then
    strMsg = strMsg & "A Quantidade do produto é inválida" & vbCrLf
    Pintar_Controle mskQuantidade, tpCorContr_Erro
  ElseIf CLng(mskQuantidade.Text) = 0 Then
    strMsg = strMsg & "A Quantidade do produto não pode ser igual a zero" & vbCrLf
    Pintar_Controle mskQuantidade, tpCorContr_Erro
  End If
  '
  If Len(mskCodigo.Text) = 0 Then
    strMsg = strMsg & "Digitar o código do produto" & vbCrLf
    Pintar_Controle mskCodigo, tpCorContr_Erro
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserEntradaMaterialInc.ValidaCamposItens]"
    ValidaCamposItens = False
  Else
    ValidaCamposItens = True
  End If
End Function

Private Function ValidaCamposEntradaMaterial() As Boolean
  Dim strMsg      As String
  Dim strSql      As String
  'Dim objRs       As ADODB.Recordset
  'Dim clsGer  As busSisLoc.clsGeral
  '
  If Len(cboDocumento.Text) = 0 Then
    strMsg = strMsg & "Selecionar o Documento" & vbCrLf
    Pintar_Controle cboDocumento, tpCorContr_Erro
  End If
  If Len(txtFornecedor.Text) = 0 Then
    strMsg = strMsg & "Informar o Fornecedor" & vbCrLf
    Pintar_Controle txtFornecedor, tpCorContr_Erro
  End If
  If Len(txtDescricao.Text) = 0 Then
    strMsg = strMsg & "Informar a Descrição do Documento" & vbCrLf
    Pintar_Controle txtDescricao, tpCorContr_Erro
  End If
  If Not Valida_Data(mskData(1), TpNaoObrigatorio) Then
    strMsg = strMsg & "Informar a Data de Aquisição do Documento válida" & vbCrLf
    Pintar_Controle mskData(1), tpCorContr_Erro
  End If
    
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserEntradaMaterialInc.ValidaCamposEntradaMaterial]"
    ValidaCamposEntradaMaterial = False
  Else
    ValidaCamposEntradaMaterial = True
  End If
End Function

Private Sub cmdPedido_Click()
  On Error GoTo trata
  frmUserEstoqueCons.QuemChamou = 0
  frmUserEstoqueCons.Show vbModal
  
  SetarFoco mskCodigo
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
  TratarErro Err.Number, Err.Description, "[frmUserEntradaMaterialInc.Form_Activate]"
End Sub




Private Sub mskQuantidade_Change()
  'If Len(Trim(mskQuantidade.Text)) = 2 Then mskCodigo.SetFocus
End Sub



Private Sub mskCodigo_Change()
  On Error GoTo trata
  Dim clsEst    As busSisLoc.clsEstoque
  Dim objRs     As ADODB.Recordset
  If Len(mskCodigo.Text) = 0 Then
    txtProduto.Text = ""
    Exit Sub
  End If
  Set clsEst = New busSisLoc.clsEstoque
  '
  Set objRs = clsEst.ListarEstoquePeloCodigo(mskCodigo.Text)
  If objRs.EOF Then
    txtProduto.Text = ""
  Else
    txtProduto.Text = objRs.Fields("DESCRICAO").Value
  End If
  '
  objRs.Close
  Set objRs = Nothing
  Set clsEst = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub mskCodigo_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  KeyAscii = TRANSFORMA_MAIUSCULA(KeyAscii)
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub mskCodigo_LostFocus()
  On Error GoTo trata
  mskCodigo.Text = UCase(mskCodigo.Text)
  Pintar_Controle mskCodigo, tpCorContr_Normal
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Load()
On Error GoTo trata
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  Dim clsEntMat     As busSisLoc.clsEntradaMaterial
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 5985
  Me.Width = 10470
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, cmdExcluir, , , , , cmdImprimir
  LerFigurasAvulsas cmdPedido, "FILTRAR.ICO", "FILTRARDown.ICO", "Visualização de itens do cardápio"
  '
  'DOCUMENTO DE ENTRADA
  strSql = "Select NOME from DOCUMENTO ORDER BY NOME"
  PreencheCombo cboDocumento, strSql, False
  '
  tabDetalhes_Click 0
  If Status = tpStatus_Incluir Then
    txtCodigo.Text = ""
    'INCLUIR_VALOR_NO_MASK mskData(0), "", TpMaskData
    txtNumero.Text = ""
    txtFornecedor.Text = ""
    txtDescricao.Text = ""
    'INCLUIR_VALOR_NO_MASK mskData(1), "", TpMaskData
    '
    tabDetalhes.TabEnabled(1) = False
    tabDetalhes.TabEnabled(2) = False
  ElseIf Status = tpStatus_Alterar Then
    'Pega Dados do Banco de dados
    Set clsEntMat = New busSisLoc.clsEntradaMaterial
    Set objRs = clsEntMat.ListarEntradaMaterial(lngENTRADAMATERIALID)
    '
    If Not objRs.EOF Then
      '
      INCLUIR_VALOR_NO_MASK mskData(0), objRs.Fields("DATA").Value, TpMaskData
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
      txtCodigo.Text = Format(IIf(Not IsNumeric(objRs.Fields("CODIGO").Value), 0, objRs.Fields("CODIGO").Value), "0000")
      If Not IsNull(objRs.Fields("DOCENTRADA").Value) Then cboDocumento.Text = objRs.Fields("DOCENTRADA").Value
      txtFornecedor.Text = objRs.Fields("FORNECEDOR").Value & ""
      INCLUIR_VALOR_NO_MASK mskData(1), objRs.Fields("DATAAQUISICAO").Value, TpMaskData
      txtNumero.Text = objRs.Fields("NUMERO").Value & ""
      '
    End If
    objRs.Close
    Set objRs = Nothing
    Set clsEntMat = Nothing
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
  If Not bFechar Then Cancel = True
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
    SetarFoco txtNumero
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
  TratarErro Err.Number, Err.Description, "SisLoc.frmUserEntradaMaterialInc.tabDetalhes"
  AmpN
End Sub


Private Sub mskQuantidade_GotFocus()
  Selecionar_Conteudo mskQuantidade
End Sub

Private Sub mskQuantidade_LostFocus()
  Pintar_Controle mskQuantidade, tpCorContr_Normal
End Sub


Private Sub mskCodigo_GotFocus()
  Selecionar_Conteudo mskCodigo
End Sub

Private Sub txtDescricao_GotFocus()
  Selecionar_Conteudo txtDescricao
End Sub

Private Sub txtDescricao_LostFocus()
  Pintar_Controle txtDescricao, tpCorContr_Normal
End Sub

Private Sub txtFornecedor_LostFocus()
  txtFornecedor.BackColor = vbWhite
End Sub

Private Sub txtFornecedor_GotFocus()
  Selecionar_Conteudo txtFornecedor
End Sub



