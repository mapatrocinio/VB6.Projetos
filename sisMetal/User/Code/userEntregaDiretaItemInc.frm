VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEntregaDiretaItemInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Entrega Direta"
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
      TabCaption(0)   =   "Controle de entrega direta"
      TabPicture(0)   =   "userEntregaDiretaItemInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "grdItemEntregaDireta"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraFiltro"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraEntregaDireta"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.Frame fraEntregaDireta 
         Caption         =   "Entrega Direta"
         Height          =   1245
         Left            =   90
         TabIndex        =   16
         Top             =   1410
         Width           =   9105
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   975
            Index           =   0
            Left            =   90
            ScaleHeight     =   975
            ScaleWidth      =   8895
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   180
            Width           =   8895
            Begin VB.TextBox txtNumeroOS 
               BackColor       =   &H00E0E0E0&
               Height          =   288
               Left            =   1230
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   3
               TabStop         =   0   'False
               Text            =   "txtNumeroOS"
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
               TabIndex        =   18
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
                  TabIndex        =   19
                  Top             =   0
                  Width           =   615
               End
            End
            Begin VB.ComboBox cboFornecedor 
               Height          =   315
               Left            =   1230
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   330
               Width           =   3435
            End
            Begin VB.Label Label2 
               Caption         =   "Ano-Número OS"
               Height          =   255
               Index           =   0
               Left            =   30
               TabIndex        =   21
               Top             =   30
               Width           =   1155
            End
            Begin VB.Label Label5 
               Caption         =   "Fornecedor"
               Height          =   195
               Index           =   0
               Left            =   30
               TabIndex        =   20
               Top             =   360
               Width           =   1215
            End
         End
      End
      Begin VB.Frame fraFiltro 
         Caption         =   "Filtro"
         Height          =   1095
         Left            =   90
         TabIndex        =   15
         Top             =   330
         Width           =   9105
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
            TabIndex        =   22
            Top             =   210
            Width           =   1095
         End
      End
      Begin TrueDBGrid60.TDBGrid grdItemEntregaDireta 
         Height          =   3090
         Left            =   90
         OleObjectBlob   =   "userEntregaDiretaItemInc.frx":001C
         TabIndex        =   6
         Top             =   2730
         Width           =   9210
      End
      Begin VB.Label Label1 
         Caption         =   $"userEntregaDiretaItemInc.frx":6082
         ForeColor       =   &H000000FF&
         Height          =   435
         Index           =   2
         Left            =   120
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   6090
         Width           =   8715
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "* Clique no botão SALVAR para registrar todas as definições de anodização para os perfis ou tecle ESC para sair"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   5820
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
Attribute VB_Name = "frmEntregaDiretaItemInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''Option Explicit
'''
'''Public Status                   As tpStatus
'''Public lngENTREGADIRETAID       As Long
'''Public lngLINHAID               As Long
'''
'''Public lngOSFINALID             As Long
'''Public lngANODIZACAOITEMID      As Long
'''Public lngITEMOSFINALID         As Long
'''Public lngOSID                  As Long
'''Public lngCORID                 As Long
'''
'''Public strOSNumero              As String
'''Public strCor                   As String
'''Dim blnAlterouPeso              As Boolean
'''
'''Dim blnFechar                   As Boolean
'''Public blnRetorno               As Boolean
'''Public blnPrimeiraVez           As Boolean
''''
'''Dim ITEMENTDIR_COLUNASMATRIZ        As Long
'''Dim ITEMENTDIR_LINHASMATRIZ         As Long
'''Private ITEMENTDIR_Matriz()         As String
'''
'''
'''Public Sub ITEMENTDIR_MontaMatriz(lngLINHASELID As Long)
'''  Dim strSql    As String
'''  Dim objRs     As ADODB.Recordset
'''  Dim intI      As Integer
'''  Dim intJ      As Integer
'''  Dim objGeral  As busSisMetal.clsGeral
'''  '
'''  On Error GoTo trata
'''
'''  Set objGeral = New busSisMetal.clsGeral
'''  '
'''  strSql = "SELECT " & _
'''            IIf(Status = tpStatus_Incluir, "0,", " ITEM_PEDIDO.PKID, ") & _
'''            " ESTOQUE.LINHAID, " & _
'''            " ESTOQUE.LINHAID, " & _
'''            Formata_Dados(IIf(Status = tpStatus_Incluir, "0", "0"), tpDados_Texto) & ", " & _
'''            " ESTOQUE.NOME, ESTOQUE.CODIGO, " & _
'''            " ESTOQUE.PESO_MINIMO, " & _
'''            "  (ESTOQUE.PESO_ESTOQUE + ESTOQUE.PEDIDO_PESO_RESTA + ESTOQUE.OS_PESO_RESTA + ESTOQUE.ANOD_PESO_RESTA) AS PESO_REAL, "
'''
'''  If Status = tpStatus_Incluir Then
'''    strSql = strSql & " (ESTOQUE.PESO_MINIMO) - (ESTOQUE.PESO_ESTOQUE + ESTOQUE.PEDIDO_PESO_RESTA + ESTOQUE.OS_PESO_RESTA + ESTOQUE.ANOD_PESO_RESTA) AS PESO_APEDIR, '' AS PESO_ANOD, '' AS PESO_FAB "
'''  Else
'''    strSql = strSql & " ITEM_PEDIDO.PESO_INI, ITEM_PEDIDO.PESO, ITEM_PEDIDO.PESO_FAB "
'''  End If
'''
'''  strSql = strSql & " From VW_CONS_ESTOQUE_PERFIL AS ESTOQUE "
'''  If Status = tpStatus_Incluir Then
'''    strSql = strSql & " WHERE (ESTOQUE.PESO_ESTOQUE + ESTOQUE.PEDIDO_PESO_RESTA + ESTOQUE.OS_PESO_RESTA + ESTOQUE.ANOD_PESO_RESTA) < ESTOQUE.PESO_MINIMO "
'''    If lngLINHASELID <> 0 Then
'''      strSql = strSql & " AND ESTOQUE.LINHAID = " & Formata_Dados(lngLINHASELID, tpDados_Longo)
'''    End If
'''  Else
'''    strSql = strSql & " INNER JOIN ITEM_PEDIDO ON ITEM_PEDIDO.LINHAID = ESTOQUE.LINHAID "
'''    strSql = strSql & " WHERE ITEM_PEDIDO.PEDIDOID = " & Formata_Dados(lngENTREGADIRETAID, tpDados_Longo)
'''  End If
'''  strSql = strSql & " ORDER BY ESTOQUE.NOME, ESTOQUE.CODIGO"
'''  Set objRs = objGeral.ExecutarSQL(strSql)
'''  If Not objRs.EOF Then
'''    ITEMENTDIR_LINHASMATRIZ = objRs.RecordCount
'''  Else
'''    ITEMENTDIR_LINHASMATRIZ = 0
'''  End If
'''  If Not objRs.EOF Then
'''    ReDim ITEMENTDIR_Matriz(0 To ITEMENTDIR_COLUNASMATRIZ - 1, 0 To ITEMENTDIR_LINHASMATRIZ - 1)
'''  Else
'''    ReDim ITEMENTDIR_Matriz(0 To ITEMENTDIR_COLUNASMATRIZ - 1, 0 To 0)
'''  End If
'''  '
'''  If Not objRs.EOF Then   'se já houver algum item
'''    For intI = 0 To ITEMENTDIR_LINHASMATRIZ - 1  'varre as linhas
'''      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
'''        For intJ = 0 To ITEMENTDIR_COLUNASMATRIZ - 1  'varre as colunas
'''          If intJ = ITEMENTDIR_COLUNASMATRIZ - 1 Then
'''            ITEMENTDIR_Matriz(intJ, intI) = intI & ""
'''          Else
'''            ITEMENTDIR_Matriz(intJ, intI) = objRs(intJ) & ""
'''          End If
'''        Next
'''        objRs.MoveNext
'''      End If
'''    Next  'próxima linha matriz
'''  End If
'''  Set objGeral = Nothing
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''End Sub
'''
'''
'''Private Sub cboFornecedor_LostFocus()
'''  Pintar_Controle cboFornecedor, tpCorContr_Normal
'''End Sub
'''
'''Private Sub cmdOK_Click()
'''  On Error GoTo trata
'''  Dim strSql                  As String
'''  Dim objEntregaDireta               As busSisMetal.clsEntregaDireta
'''  Dim objRs                   As ADODB.Recordset
'''  Dim objGer                  As busSisMetal.clsGeral
'''  Dim lngFORNECEDORID         As Long
'''  Dim objItemEntregaDireta           As busSisMetal.clsItemEntregaDireta
'''  Dim intI      As Integer
'''  '
'''  Select Case tabDetalhes.Tab
'''  Case 0 'Gravar Anodização
'''    If Not ValidaCampos Then Exit Sub
'''
'''    If ValidaCamposAnodOrigemAll Then
'''      SetarFoco grdItemEntregaDireta
'''      If Status = tpStatus_Incluir Then
'''        grdItemEntregaDireta.Col = 8
'''      Else
'''        grdItemEntregaDireta.Col = 9
'''      End If
'''      grdItemEntregaDireta.Row = 0
'''      Exit Sub
'''    End If
'''    'OK procede com o cadastro
'''    'CADASTRO DE PEDIDO
'''    '-------------------------
'''    Set objGer = New busSisMetal.clsGeral
'''    'FORNECEDOR
'''    lngFORNECEDORID = 0
'''    strSql = "SELECT LOJA.PKID FROM LOJA " & _
'''      " INNER JOIN FORNECEDOR ON FORNECEDOR.LOJAID = LOJA.PKID " & _
'''      " WHERE LOJA.NOME = " & Formata_Dados(cboFornecedor.Text, tpDados_Texto)
'''    Set objRs = objGer.ExecutarSQL(strSql)
'''    If Not objRs.EOF Then
'''      lngFORNECEDORID = objRs.Fields("PKID").Value
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
'''    Set objGer = Nothing
'''    '
'''    Set objEntregaDireta = New busSisMetal.clsEntregaDireta
'''    'Altera ou incluiu pedido
'''    If Status = tpStatus_Alterar Then
'''      'Código para alteração
'''      '
'''      objEntregaDireta.AlterarEntregaDireta lngENTREGADIRETAID, _
'''                                            lngFORNECEDORID
'''      '
'''      blnRetorno = True
'''    ElseIf Status = tpStatus_Incluir Then
'''      'Código para inclusão
'''      '
'''      objEntregaDireta.InserirEntregaDireta lngENTREGADIRETAID, _
'''                                            lngFORNECEDORID
'''
'''      '
'''      blnRetorno = True
'''    End If
'''    Set objEntregaDireta = Nothing
'''    '
'''    Set objItemEntregaDireta = New busSisMetal.clsItemEntregaDireta
'''    For intI = 0 To ITEMENTDIR_LINHASMATRIZ - 1
'''      grdItemEntregaDireta.Bookmark = CLng(intI)
'''      'If grdItemEntregaDireta.Columns("Branco").Text & "" <> "" Or _
'''        grdItemEntregaDireta.Columns("Brilho").Text & "" <> "" Or _
'''        grdItemEntregaDireta.Columns("Bronze").Text & "" <> "" Or _
'''        grdItemEntregaDireta.Columns("Natural").Text & "" <> "" Then
'''      If grdItemEntregaDireta.Columns("*").Text & "" = "-1" Then
'''        'Propósito: Cadastrar pedido
'''        '
'''        objItemEntregaDireta.InserirItemEntregaDiretaItem grdItemEntregaDireta.Columns("ITEM_PEDIDOID").Text & "", _
'''                                                          lngENTREGADIRETAID, _
'''                                                          grdItemEntregaDireta.Columns("LINHAID").Text & "", _
'''                                                          IIf(grdItemEntregaDireta.Columns("Peso").Text & "" = "", "0", grdItemEntregaDireta.Columns("Peso").Text & "")
'''
'''        blnRetorno = True
'''      End If
'''    Next
'''    Set objItemEntregaDireta = Nothing
'''    '
'''    blnFechar = True
'''    Unload Me
'''  End Select
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''  '
'''  grdItemEntregaDireta.Bookmark = Null
'''  grdItemEntregaDireta.ReBind
'''  SetarFoco grdItemEntregaDireta
'''  If Status = tpStatus_Incluir Then
'''    grdItemEntregaDireta.Col = 8
'''  Else
'''    grdItemEntregaDireta.Col = 9
'''  End If
'''  grdItemEntregaDireta.Row = 0
'''
'''End Sub
'''
'''Private Sub LimparCampos()
'''  Dim sMask As String
'''
'''  On Error GoTo trata
'''  'Filtro
'''  LimparCampoTexto txtCodigo
'''  LimparCampoTexto txtCodigoFim
'''  LimparCampoTexto txtLinhaFim
'''  'EntregaDireta
'''  LimparCampoTexto txtNumeroOS
'''  LimparCampoMask mskData(0)
'''  LimparCampoCombo cboFornecedor
'''  LimparCampoMask mskValor
'''  LimparCampoCombo cboAnodizadora
'''  LimparCampoCombo cboFabrica
'''  '
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, _
'''            "[frmEntregaDiretaItemEntregaDiretaInc.LimparCampos]", _
'''            Err.Description
'''End Sub
'''
'''Private Function ValidaCampos() As Boolean
'''  On Error GoTo trata
'''  Dim strMsg                As String
'''  Dim blnSetarFocoControle  As Boolean
'''  '
'''  blnSetarFocoControle = True
'''  '
'''  If Not Valida_String(cboFornecedor, TpObrigatorio, blnSetarFocoControle) Then
'''    strMsg = strMsg & "Selecionar o fornecedor" & vbCrLf
'''  End If
'''  If Not Valida_Moeda(mskValor, TpObrigatorio, blnSetarFocoControle) Then
'''    strMsg = strMsg & "Valor do Alumínio inválido" & vbCrLf
'''  End If
'''  If Not Valida_String(cboAnodizadora, IIf(Status = tpStatus_Incluir, TpNaoObrigatorio, TpObrigatorio), blnSetarFocoControle) Then
'''    strMsg = strMsg & "Selecionar a anodizadora" & vbCrLf
'''  End If
'''  If Not Valida_String(cboFabrica, IIf(Status = tpStatus_Incluir, TpNaoObrigatorio, TpObrigatorio), blnSetarFocoControle) Then
'''    strMsg = strMsg & "Selecionar a fábrica" & vbCrLf
'''  End If
'''  If Len(strMsg) <> 0 Then
'''    TratarErroPrevisto strMsg, "[frmEntregaDiretaInc.ValidaCampos]"
'''    ValidaCampos = False
'''  Else
'''    ValidaCampos = True
'''  End If
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            "[frmEntregaDiretaItemEntregaDiretaInc.ValidaCampos]", _
'''            Err.Description
'''End Function
'''
'''Private Sub Form_Activate()
'''  On Error GoTo trata
'''  If blnPrimeiraVez Then
'''    'Montar RecordSet
'''    ITEMENTDIR_COLUNASMATRIZ = grdItemEntregaDireta.Columns.Count
'''    ITEMENTDIR_LINHASMATRIZ = 0
'''    ITEMENTDIR_MontaMatriz (lngLINHAID)
'''    grdItemEntregaDireta.Bookmark = Null
'''    grdItemEntregaDireta.ReBind
'''    grdItemEntregaDireta.ApproxCount = ITEMENTDIR_LINHASMATRIZ
'''    '
'''    If Status = tpStatus_Incluir Then
'''      SetarFoco txtCodigo
'''    ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
'''      SetarFoco cboFornecedor
'''    End If
'''    'SetarFoco grdItemEntregaDireta
'''    'grdItemEntregaDireta.Col = 8
'''    blnPrimeiraVez = False
'''  End If
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, "[frmEntregaDiretaItemEntregaDiretaInc.Form_Activate]"
'''End Sub
'''
'''
'''Private Sub grdItemEntregaDireta_BeforeUpdate(Cancel As Integer)
'''  On Error GoTo trata
'''  'Atualiza Matriz
'''  If blnAlterouPeso = True Then
'''    ITEMENTDIR_Matriz(3, grdItemEntregaDireta.Row) = "-1"
'''  Else
'''    ITEMENTDIR_Matriz(3, grdItemEntregaDireta.Row) = grdItemEntregaDireta.Columns(3).Text
'''  End If
'''  If Status = tpStatus_Incluir Then
'''    ITEMENTDIR_Matriz(8, grdItemEntregaDireta.Row) = grdItemEntregaDireta.Columns(8).Text
'''  Else
'''    ITEMENTDIR_Matriz(9, grdItemEntregaDireta.Row) = grdItemEntregaDireta.Columns(9).Text
'''    ITEMENTDIR_Matriz(10, grdItemEntregaDireta.Row) = grdItemEntregaDireta.Columns(10).Text
'''  End If
'''  blnAlterouPeso = False
'''  '
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, "[frmEntregaDiretaItemEntregaDiretaInc.grdItemEntregaDireta_BeforeRowColChange]"
'''End Sub
'''
'''Private Sub grdItemEntregaDireta_ColEdit(ByVal ColIndex As Integer)
'''  On Error GoTo trata
'''  '
'''  If grdItemEntregaDireta.Col = 8 Or grdItemEntregaDireta.Col = 9 Or grdItemEntregaDireta.Col = 10 Then
'''    blnAlterouPeso = True
'''  End If
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, "[frmEntregaDiretaItemEntregaDiretaInc.grdItemEntregaDireta_ColEdit]"
'''End Sub
'''
'''Private Sub grdItemEntregaDireta_GotFocus()
'''  On Error Resume Next
'''  If Status = tpStatus_Incluir Then
'''    grdItemEntregaDireta.Col = 8
'''  Else
'''    grdItemEntregaDireta.Col = 9
'''  End If
'''End Sub
'''
'''
'''Private Sub grdItemEntregaDireta_UnboundReadDataEx( _
'''     ByVal RowBuf As TrueDBGrid60.RowBuffer, _
'''    StartLocation As Variant, ByVal Offset As Long, _
'''    ApproximatePosition As Long)
'''  ' UnboundReadData is fired by an unbound grid whenever
'''  ' it requires data for display. This event will fire
'''  ' when the grid is first shown, when Refresh or ReBind
'''  ' is used, when the grid is scrolled, and after a
'''  ' record in the grid is modified and the user commits
'''  ' the change by moving off of the current row. The
'''  ' grid fetches data in "chunks", and the number of rows
'''  ' the grid is asking for is given by RowBuf.RowCount.
'''  ' RowBuf is the row buffer where you place the data
'''  ' the bookmarks for the rows that the grid is
'''  ' requesting to display. It will also hold the number
'''  ' of rows that were successfully supplied to the grid.
'''  ' StartLocation is a vrtBookmark which, together with
'''  ' Offset, specifies the row for the programmer to start
'''  ' transferring data. A StartLocation of Null indicates
'''  ' a request for data from BOF or EOF.
'''  ' Offset specifies the relative position (from
'''  ' StartLocation) of the row for the programmer to start
'''  ' transferring data. A positive number indicates a
'''  ' forward relative position while a negative number
'''  ' indicates a backward relative position. Regardless
'''  ' of whether the rows to be read are before or after
'''  ' StartLocation, rows are always fetched going forward
'''  ' (this is why there is no ReadPriorRows parameter to
'''  ' the procedure).
'''  ' If you page down on the grid, for instance, the new
'''  ' top row of the grid will have an index greater than
'''  ' the StartLocation (Offset > 0). If you page up on
'''  ' the grid, the new index is less than that of
'''  ' StartLocation, so Offset < 0. If StartLocation is
'''  ' a vrtBookmark to row N, the grid always asks for row
'''  ' data in the following order:
'''  '   (N + Offset), (N + Offset + 1), (N + Offset + 2)...
'''  ' ApproximatePosition is a value you can set to indicate
'''  ' the ordinal position of (StartLocation + Offset).
'''  ' Setting this variable will enhance the ability of the
'''  ' grid to display its vertical scroll bar accurately.
'''  ' If the exact ordinal position of the new location is
'''  ' not known, you can set it to a reasonable,
'''  ' approximate value, or just ignore this parameter.
'''
'''  On Error GoTo trata
'''  '
'''  Dim intColIndex      As Integer
'''  Dim intJ             As Integer
'''  Dim intRowsFetched   As Integer
'''  Dim intI             As Long
'''  Dim lngNewPosition   As Long
'''  Dim vrtBookmark      As Variant
'''  '
'''  intRowsFetched = 0
'''  For intI = 0 To RowBuf.RowCount - 1
'''    ' Get the vrtBookmark of the next available row
'''    vrtBookmark = GetRelativeBookmarkGeral(StartLocation, _
'''               Offset + intI, ITEMENTDIR_LINHASMATRIZ)
'''
'''    ' If the next row is BOF or EOF, then stop fetching
'''    ' and return any rows fetched up to this point.
'''    If IsNull(vrtBookmark) Then Exit For
'''
'''    ' Place the record data into the row buffer
'''    For intJ = 0 To RowBuf.ColumnCount - 1
'''      intColIndex = RowBuf.ColumnIndex(intI, intJ)
'''      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
'''                           intColIndex, ITEMENTDIR_COLUNASMATRIZ, ITEMENTDIR_LINHASMATRIZ, ITEMENTDIR_Matriz)
'''    Next intJ
'''
'''    ' Set the vrtBookmark for the row
'''    RowBuf.Bookmark(intI) = vrtBookmark
'''
'''    ' Increment the count of fetched rows
'''    intRowsFetched = intRowsFetched + 1
'''  Next intI
'''
'''' Tell the grid how many rows were fetched
'''  RowBuf.RowCount = intRowsFetched
'''
'''' Set the approximate scroll bar position. Only
'''' nonnegative values of IndexFromBookmark() are valid.
'''  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ITEMENTDIR_LINHASMATRIZ)
'''  If lngNewPosition >= 0 Then _
'''     ApproximatePosition = lngNewPosition
'''
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, "[frmAnodizadoraInc.grdItemEntregaDireta_UnboundReadDataEx]"
'''End Sub
'''
'''Private Sub Form_Load()
'''  On Error GoTo trata
'''  Dim objRs         As ADODB.Recordset
'''  Dim strSql        As String
'''  Dim objEntregaDireta     As busSisMetal.clsEntregaDireta
'''  '
'''  blnFechar = False 'Não Pode Fechar pelo X
'''  blnRetorno = False
'''  blnPrimeiraVez = True
'''  blnAlterouPeso = False
'''  lngLINHAID = 0
'''  '
'''  AmpS
'''  Me.Height = 7620
'''  Me.Width = 11610
'''  CenterForm Me
'''  '
'''  LimparCampos
'''  '
'''  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdFechar
'''  '
'''  'Fornecedor
'''  strSql = "SELECT LOJA.NOME FROM LOJA " & _
'''      " INNER JOIN FORNECEDOR ON LOJA.PKID = FORNECEDOR.LOJAID " & _
'''      " WHERE LOJA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
'''      "ORDER BY LOJA.NOME"
'''  PreencheCombo cboFornecedor, strSql, False, True
'''  'Anodizadora
'''  strSql = "SELECT LOJA.NOME FROM LOJA " & _
'''      " INNER JOIN ANODIZADORA ON LOJA.PKID = ANODIZADORA.LOJAID " & _
'''      " WHERE LOJA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
'''      "ORDER BY LOJA.NOME"
'''  PreencheCombo cboAnodizadora, strSql, False, True
'''  'Fabrica
'''  strSql = "SELECT LOJA.NOME FROM LOJA " & _
'''      " INNER JOIN FABRICA ON LOJA.PKID = FABRICA.LOJAID " & _
'''      " WHERE LOJA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
'''      "ORDER BY LOJA.NOME"
'''  PreencheCombo cboFabrica, strSql, False, True
'''  '
'''  If Status = tpStatus_Incluir Then
'''    'Caso esteja em um evento de Inclusão, Inclui o EntregaDireta
'''    '
'''    fraFiltro.Enabled = True
'''    'No evento de inclusão deve ser habilitado a coluna peso
'''    grdItemEntregaDireta.Columns(8).Locked = False
'''    grdItemEntregaDireta.Columns(9).Visible = False
'''    grdItemEntregaDireta.Columns(10).Visible = False
'''  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
'''    'Pega Dados do Banco de dados
'''    Set objEntregaDireta = New busSisMetal.clsEntregaDireta
'''    Set objRs = objEntregaDireta.ListarEntregaDireta(lngENTREGADIRETAID)
'''    '
'''    If Not objRs.EOF Then
'''      'Campos fixos
'''      txtNumeroOS = objRs.Fields("OS_ANO").Value & "-" & Format(objRs.Fields("OS_NUMERO").Value & "", "0000")
'''      INCLUIR_VALOR_NO_MASK mskData(0), objRs.Fields("DATA").Value, TpMaskData
'''      'Campos inserts
'''      INCLUIR_VALOR_NO_COMBO objRs.Fields("NOME_FORNECEDOR").Value & "", cboFornecedor
'''      INCLUIR_VALOR_NO_COMBO objRs.Fields("NOME_ANODIZADORA").Value & "", cboAnodizadora
'''      INCLUIR_VALOR_NO_COMBO objRs.Fields("NOME_FABRICA").Value & "", cboFabrica
'''      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VALOR_ALUMINIO").Value, TpMaskMoeda
'''    End If
'''    Set objEntregaDireta = Nothing
'''    '
'''    fraFiltro.Enabled = False
'''    'No evento de alteração deve ser habilitado as colunas anod e fabrica
'''    grdItemEntregaDireta.Columns(8).Locked = True
'''    grdItemEntregaDireta.Columns(9).Visible = True
'''    grdItemEntregaDireta.Columns(10).Visible = True
'''  End If
'''  '
'''  AmpN
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''  AmpN
'''End Sub
'''
'''Private Function ValidaCamposAnodOrigemLinha(intLinha As Integer) As Boolean
'''  On Error GoTo trata
'''  Dim strMsg                As String
'''  Dim blnSetarFocoControle  As Boolean
'''  Dim strSql                As String
'''  Dim objRs                 As ADODB.Recordset
'''  Dim objGeral              As busSisMetal.clsGeral
'''  Dim objItemEntregaDireta         As busSisMetal.clsItemEntregaDireta
'''  '
'''  Dim lngQtdIni             As Long
'''  Dim lngQtdAnod            As Long
'''  Dim lngQtdFab             As Long
'''
'''  '
'''  blnSetarFocoControle = True
'''  '
'''  strMsg = ""
'''  'Validção dos ítens do pedido
'''  If Not Valida_Moeda(grdItemEntregaDireta.Columns("Peso"), TpNaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
'''    strMsg = strMsg & "Peso inválido na linha " & intLinha + 1 & vbCrLf
'''  End If
'''  If Not Valida_Moeda(grdItemEntregaDireta.Columns("Anod."), TpNaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
'''    strMsg = strMsg & "Peso para anodizadora inválido na linha " & intLinha + 1 & vbCrLf
'''  End If
'''  If Not Valida_Moeda(grdItemEntregaDireta.Columns("Fábrica"), TpNaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
'''    strMsg = strMsg & "Peso inválido para fábrica na linha " & intLinha + 1 & vbCrLf
'''  End If
'''  '
'''  If Len(strMsg) = 0 Then
'''    'NOVO - Validações de cálculo de peças (quantidade)
'''    Set objItemEntregaDireta = New busSisMetal.clsItemEntregaDireta
'''    lngQtdIni = objItemEntregaDireta.CalculoQuantidadeEntregaDireta(grdItemEntregaDireta.Columns("LINHAID").Value, _
'''                                                      grdItemEntregaDireta.Columns("Peso").Value)
'''    If lngQtdIni = 0 Then
'''      strMsg = strMsg & "A quantidade calculada para o perfil deve ser maior que zero na linha " & intLinha + 1 & vbCrLf
'''    End If
'''    If grdItemEntregaDireta.Columns("Anod.").Value & "" <> "" Then
'''      'Lançou peso para anodização
'''      lngQtdAnod = objItemEntregaDireta.CalculoQuantidadeEntregaDireta(grdItemEntregaDireta.Columns("LINHAID").Value, _
'''                                                        grdItemEntregaDireta.Columns("Anod.").Value)
'''      If lngQtdAnod = 0 Then
'''        strMsg = strMsg & "A quantidade calculada para anodização para o perfil deve ser maior que zero na linha " & intLinha + 1 & vbCrLf
'''      End If
'''    End If
'''    If grdItemEntregaDireta.Columns("Fábrica").Value & "" <> "" Then
'''      'Lançou peso para anodização
'''      lngQtdFab = objItemEntregaDireta.CalculoQuantidadeEntregaDireta(grdItemEntregaDireta.Columns("LINHAID").Value, _
'''                                                        grdItemEntregaDireta.Columns("Fábrica").Value)
'''      If lngQtdFab = 0 Then
'''        strMsg = strMsg & "A quantidade calculada para fábrica para o perfil deve ser maior que zero na linha " & intLinha + 1 & vbCrLf
'''      End If
'''    End If
'''
'''
''''''    lngTotal = 0
''''''    lngTotalANOD = 0
''''''    '
''''''    lngTotal = CLng(grdItemEntregaDireta.Columns("Qtd. Total")) - CLng(grdItemEntregaDireta.Columns("Qtd. Baixa"))
''''''    lngTotalANOD = CLng(IIf(Not IsNumeric(grdItemEntregaDireta.Columns("Quantidade")), 0, grdItemEntregaDireta.Columns("Quantidade")))
'''    '
'''    Set objItemEntregaDireta = Nothing
'''  End If
'''
'''  If Len(strMsg) <> 0 Then
'''    TratarErroPrevisto strMsg, "[frmEntregaDiretaItemEntregaDiretaInc.ValidaCamposAnodOrigemLinha]"
'''    ValidaCamposAnodOrigemLinha = False
'''  Else
'''    ValidaCamposAnodOrigemLinha = True
'''  End If
'''  Exit Function
'''trata:
'''  TratarErro Err.Number, _
'''             Err.Description, _
'''             "[frmUserLeituraFechaInc.ValidaCamposAnodOrigemLinha]"
'''  ValidaCamposAnodOrigemLinha = False
'''End Function
'''
'''Private Function ValidaCamposAnodOrigemAll() As Boolean
'''  On Error GoTo trata
'''  Dim blnRetorno            As Boolean
'''  Dim blnCadastrou1Linha    As Boolean
'''  Dim blnEncontrouErro      As Boolean
'''  Dim blnEncontrouErroLinha As Boolean
'''  Dim intRows               As Integer
'''  'Validar todas as linhas da matriz
'''  blnEncontrouErro = False
'''  blnCadastrou1Linha = False
'''  blnEncontrouErroLinha = False
'''  blnRetorno = True
'''
'''
'''  For intRows = 0 To ITEMENTDIR_LINHASMATRIZ - 1
'''    grdItemEntregaDireta.Bookmark = CLng(intRows)
'''    '
'''    If grdItemEntregaDireta.Columns("*").Text & "" = "-1" Then
'''      'Somente válida se preencheu algo, sneão considera ok
'''      If grdItemEntregaDireta.Columns("Peso").Text & "" <> "" Then
'''        If Not ValidaCamposAnodOrigemLinha(grdItemEntregaDireta.Row) Then
'''          blnEncontrouErro = True
'''          blnEncontrouErroLinha = True
'''        Else
'''          blnCadastrou1Linha = True
'''        End If
'''      Else
'''        'tudo brnao, considera OK
'''        blnCadastrou1Linha = True
'''      End If
'''    Else
'''      'blnEncontrouErro = True
'''    End If
'''    If blnEncontrouErro = True Then Exit For
'''  Next
'''  '
'''  If blnEncontrouErro = False And blnCadastrou1Linha = True Then
'''    blnRetorno = False
'''  End If
'''  If blnEncontrouErroLinha = False And blnEncontrouErro = False And blnCadastrou1Linha = False Then
'''    'não ouve erro
'''    If Status = tpStatus_Incluir Then
'''      TratarErroPrevisto "Selecione no mínimo 1 perfil para cadastro", "[frmEntregaDiretaItemEntregaDiretaInc.ValidaCamposAnodOrigemAll]"
'''    Else
'''      'No caso da alteração não é obrigatório cadastrar ou alterar ítem
'''      'pode estar apenas alterando dados do pedido
'''      blnRetorno = False
'''    End If
'''
'''  End If
'''  grdItemEntregaDireta.ReBind
'''  grdItemEntregaDireta.SetFocus
'''  ValidaCamposAnodOrigemAll = blnRetorno
'''  Exit Function
'''trata:
'''  TratarErro Err.Number, _
'''             Err.Description, _
'''             "[frmUserLeituraFechaInc.ValidaCamposAnodOrigemAll]"
'''  ValidaCamposAnodOrigemAll = False
'''End Function
'''Private Function ValidaCamposAnodOrigem(intLinha As Integer, intColuna As Integer) As Boolean
'''  On Error GoTo trata
'''  Dim strMsg                As String
'''  Dim blnSetarFocoControle  As Boolean
'''  '
'''  blnSetarFocoControle = True
'''  '
'''  Select Case intColuna
'''  Case 5
'''    'Validção da quantidade branco
'''    If Not Valida_Moeda(grdItemEntregaDireta.Columns(5), TpNaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
'''      strMsg = strMsg & "Quantidade de perfil branco inválida" & vbCrLf
'''    End If
'''  Case 6
'''    'Validção da quantidade brilho
'''    If Not Valida_Moeda(grdItemEntregaDireta.Columns(6), TpNaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
'''      strMsg = strMsg & "Quantidade de perfil brilho inválida" & vbCrLf
'''    End If
'''  Case 7
'''    'Validção da quantidade bronze
'''    If Not Valida_Moeda(grdItemEntregaDireta.Columns(7), TpNaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
'''      strMsg = strMsg & "Quantidade de perfil bronze inválida" & vbCrLf
'''    End If
'''  Case 8
'''    'Validção da quantidade natural
'''    If Not Valida_Moeda(grdItemEntregaDireta.Columns(8), TpNaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
'''      strMsg = strMsg & "Quantidade de perfil natural inválida" & vbCrLf
'''    End If
'''  End Select
'''  '
'''  If Len(strMsg) <> 0 Then
'''    TratarErroPrevisto strMsg, "[frmEntregaDiretaItemEntregaDiretaInc.ValidaCamposAnodOrigem]"
'''    ValidaCamposAnodOrigem = False
'''  Else
'''    ValidaCamposAnodOrigem = True
'''  End If
'''  Exit Function
'''trata:
'''  TratarErro Err.Number, _
'''             Err.Description, _
'''             "[frmUserGrupoCln.ValidaCamposAnodOrigem]"
'''  ValidaCamposAnodOrigem = False
'''End Function
'''Private Sub cmdFechar_Click()
'''  blnFechar = True
'''  '
'''  Unload Me
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''  Unload Me
'''End Sub
'''
'''Private Sub Form_Unload(Cancel As Integer)
'''  If Not blnFechar Then Cancel = True
'''End Sub
'''Private Sub Form_KeyPress(KeyAscii As Integer)
'''  On Error GoTo trata
'''  '
'''  Dim intUltimaColuna   As Integer
'''  If Me.ActiveControl.Name <> "grdItemEntregaDireta" Then
'''    If KeyAscii = 13 Then
'''      SendKeys "{tab}"
'''    End If
'''  Else
'''    If Status = tpStatus_Incluir Then
'''      intUltimaColuna = 8
'''    Else
'''      intUltimaColuna = 10
'''    End If
'''    If KeyAscii = 13 And grdItemEntregaDireta.Row <> -1 Then
'''      If grdItemEntregaDireta.Col = intUltimaColuna Then
'''        If grdItemEntregaDireta.Columns("ROWNUM").Value + 1 = ITEMENTDIR_LINHASMATRIZ Then
'''          cmdOK_Click
'''        Else
'''          If Status = tpStatus_Incluir Then
'''            grdItemEntregaDireta.Col = intUltimaColuna
'''          Else
'''            grdItemEntregaDireta.Col = intUltimaColuna - 1
'''          End If
'''
'''          grdItemEntregaDireta.MoveNext
'''        End If
'''      Else
'''        grdItemEntregaDireta.Col = grdItemEntregaDireta.Col + 1
'''      End If
'''    ElseIf (KeyAscii = 8) Then
'''    ElseIf (KeyAscii < 48 Or KeyAscii > 57) Then
'''      KeyAscii = 0
'''    End If
'''  End If
'''  '
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, "[frmEntregaDiretaItemEntregaDiretaInc.Form_KeyPress]"
'''End Sub
'''
'''
'''Private Sub mskValor_GotFocus()
'''  Seleciona_Conteudo_Controle mskValor
'''End Sub
'''Private Sub mskValor_LostFocus()
'''  Pintar_Controle mskValor, tpCorContr_Normal
'''End Sub
'''
'''Private Sub txtCodigo_LostFocus()
'''  On Error GoTo trata
'''  Dim objLinhaCons    As Form
'''  Dim objLinhaPerfil  As busSisMetal.clsLinhaPerfil
'''  Dim objRs           As ADODB.Recordset
'''  Dim blnSelecionou   As Boolean
'''  If Me.ActiveControl.Name = "cmdFechar" Then Exit Sub
'''  If Me.ActiveControl.Name = "cmdOk" Then Exit Sub
'''
'''  Pintar_Controle txtCodigo, tpCorContr_Normal
'''  If Len(txtCodigo.Text) = 0 Then
'''    If Len(txtCodigoFim.Text) <> 0 And Len(txtLinhaFim.Text) <> 0 Then
'''      lngLINHAID = 0
'''      '
'''      ITEMENTDIR_MontaMatriz (lngLINHAID)
'''      grdItemEntregaDireta.Bookmark = Null
'''      grdItemEntregaDireta.ReBind
'''      grdItemEntregaDireta.ApproxCount = ITEMENTDIR_LINHASMATRIZ
'''      '
'''      Exit Sub
'''    Else
'''      'TratarErroPrevisto "Entre com o código ou descrição da linha."
'''      'Pintar_Controle txtCodigo, tpCorContr_Erro
'''      'SetarFoco txtCodigo
'''      'Exit Sub
'''      lngLINHAID = 0
'''      '
'''      ITEMENTDIR_MontaMatriz (lngLINHAID)
'''      grdItemEntregaDireta.Bookmark = Null
'''      grdItemEntregaDireta.ReBind
'''      grdItemEntregaDireta.ApproxCount = ITEMENTDIR_LINHASMATRIZ
'''      '
'''      Exit Sub
'''
'''    End If
'''  End If
'''  blnSelecionou = False
'''  Set objLinhaPerfil = New busSisMetal.clsLinhaPerfil
'''  '
'''  Set objRs = objLinhaPerfil.CapturaItemLinha(txtCodigo.Text)
'''  If objRs.EOF Then
'''    LimparCampoTexto txtCodigoFim
'''    LimparCampoTexto txtLinhaFim
'''    TratarErroPrevisto "Descrição/Código da linha não cadastrado"
'''    Pintar_Controle txtCodigo, tpCorContr_Erro
'''    SetarFoco txtCodigo
'''    Exit Sub
'''  Else
'''    If objRs.RecordCount = 1 Then
'''      txtCodigoFim.Text = objRs.Fields("CODIGO").Value & ""
'''      txtLinhaFim.Text = objRs.Fields("NOME").Value & ""
'''      lngLINHAID = objRs.Fields("PKID").Value & ""
'''      blnSelecionou = True
'''    Else
'''      'Novo : apresentar tela para seleção da linha
'''      Set objLinhaCons = New frmLinhaCons
'''      objLinhaCons.intIcOrigemLn = 4
'''      objLinhaCons.strCodigoDescricao = txtCodigo.Text
'''      objLinhaCons.Show vbModal
'''      blnSelecionou = True
'''      Set objLinhaCons = Nothing
'''    End If
'''    '
'''    objRs.Close
'''    Set objRs = Nothing
'''    Set objLinhaPerfil = Nothing
'''  End If
'''  ''''recarregar GRID
'''  If blnSelecionou = True Then
'''    '
'''    ITEMENTDIR_MontaMatriz (lngLINHAID)
'''    grdItemEntregaDireta.Bookmark = Null
'''    grdItemEntregaDireta.ReBind
'''    grdItemEntregaDireta.ApproxCount = ITEMENTDIR_LINHASMATRIZ
'''  End If
'''
'''
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''End Sub
'''
'''Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
'''  On Error GoTo trata
'''  KeyAscii = TRANSFORMA_MAIUSCULA(KeyAscii)
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''End Sub
'''
'''Private Sub txtCodigo_GotFocus()
'''  Seleciona_Conteudo_Controle txtCodigo
'''End Sub
'''
