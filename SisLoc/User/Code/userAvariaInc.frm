VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserAvariaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Avaria de �tens da NF"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   12465
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7470
      Left            =   10605
      ScaleHeight     =   7470
      ScaleWidth      =   1860
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   1275
         Left            =   90
         ScaleHeight     =   1215
         ScaleWidth      =   1605
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   6120
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   150
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   7245
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   150
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   12779
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&�tens"
      TabPicture(0)   =   "userAvariaInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Dados da avaria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2715
         Left            =   120
         TabIndex        =   17
         Top             =   390
         Width           =   7695
         Begin VB.TextBox txtDescricao 
            Height          =   735
            Left            =   1350
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   5
            Text            =   "userAvariaInc.frx":001C
            Top             =   1590
            Width           =   6135
         End
         Begin VB.TextBox txtSequencial 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   5130
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   4
            TabStop         =   0   'False
            Text            =   "txtSequencial"
            Top             =   1260
            Width           =   2385
         End
         Begin VB.ComboBox cboAvaria 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   960
            Width           =   2385
         End
         Begin VB.TextBox txtContrato 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   1350
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   1
            TabStop         =   0   'False
            Text            =   "txtContrato"
            Top             =   660
            Width           =   2385
         End
         Begin VB.TextBox txtEmpresa 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   1350
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   0
            TabStop         =   0   'False
            Text            =   "txtEmpresa"
            Top             =   360
            Width           =   6165
         End
         Begin MSMask.MaskEdBox mskData 
            Height          =   255
            Left            =   1350
            TabIndex        =   3
            Top             =   1290
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
            Caption         =   "Descri��o"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   23
            Top             =   1590
            Width           =   1125
         End
         Begin VB.Label Label3 
            Caption         =   "Sequencial"
            Height          =   255
            Left            =   3930
            TabIndex        =   22
            Top             =   1260
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Avaria"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   21
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Contrato"
            Height          =   255
            Left            =   150
            TabIndex        =   20
            Top             =   660
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Data"
            Height          =   225
            Left            =   150
            TabIndex        =   19
            Top             =   1290
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Empresa"
            Height          =   255
            Left            =   150
            TabIndex        =   18
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Pacotes a serem criados / alterados "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   120
         TabIndex        =   16
         Top             =   3150
         Width           =   10215
         Begin VB.CommandButton cmdCadastraItem 
            Caption         =   ">"
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   5190
            TabIndex        =   7
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdCadastraItem 
            Caption         =   ">>"
            Height          =   375
            Index           =   1
            Left            =   5190
            TabIndex        =   8
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton cmdCadastraItem 
            Caption         =   "<"
            Height          =   375
            Index           =   2
            Left            =   5190
            TabIndex        =   9
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton cmdCadastraItem 
            Caption         =   "<<"
            Height          =   375
            Index           =   3
            Left            =   5190
            TabIndex        =   10
            Top             =   1320
            Width           =   375
         End
         Begin TrueDBGrid60.TDBGrid grdNF 
            Height          =   3555
            Left            =   90
            OleObjectBlob   =   "userAvariaInc.frx":002B
            TabIndex        =   6
            Top             =   240
            Width           =   5055
         End
         Begin TrueDBGrid60.TDBGrid grdNFOrigem 
            Height          =   3555
            Left            =   5640
            OleObjectBlob   =   "userAvariaInc.frx":7B72
            TabIndex        =   11
            Top             =   240
            Width           =   4455
         End
      End
   End
End
Attribute VB_Name = "frmUserAvariaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                 As tpStatus
Public lngCONTRATOID          As Long
Public lngAVARIAID            As Long
Public blnRetorno             As Boolean
Public blnFechar              As Boolean
Public blnPrimeiraVez         As Boolean
Dim ITEMNF_COLUNASMATRIZ      As Long
Dim ITEMNF_LINHASMATRIZ       As Long
Private ITEMNF_Matriz()       As String

Dim ITEMNFLANC_COLUNASMATRIZ         As Long
Dim ITEMNFLANC_LINHASMATRIZ          As Long
Private ITEMNFLANC_Matriz()          As String

Public Sub ITEMNF_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim objRsConf As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisLoc.clsGeral
  Dim strListaConfiguracaoId As String
  '
  On Error GoTo trata

  Set clsGer = New busSisLoc.clsGeral
  '
  'strSql = "SELECT ITEMNF.PKID, ITEMNF.NFID, ITEMNF.ESTOQUEID, ESTOQUE.DESCRICAO, (ITEMNF.QUANTIDADE - ISNULL(ITENS_DEVOLVIDOS.QTD_DEVOL, 0) - ISNULL(ITENS_AVARIA.QTD_AVARIA, 0)) AS QTD_REAL, '' As ALancar, ESTOQUE.VALORINDENIZACAO "
  'strSql = strSql & " FROM ITEMNF INNER JOIN ESTOQUE ON ESTOQUE.PKID = ITEMNF.ESTOQUEID " & _
          "LEFT JOIN (SELECT ITEMNFID, ISNULL(SUM(QUANTIDADE), 0) AS QTD_DEVOL " & _
          "FROM ITEMDEVOLUCAO GROUP BY ITEMNFID) AS ITENS_DEVOLVIDOS ON ITEMNF.PKID = ITENS_DEVOLVIDOS.ITEMNFID " & _
          "LEFT JOIN (SELECT ITEMNFID, ISNULL(SUM(QUANTIDADE), 0) AS QTD_AVARIA " & _
          "FROM ITEMAVARIA GROUP BY ITEMNFID) AS ITENS_AVARIA ON ITEMNF.PKID = ITENS_AVARIA.ITEMNFID " & _
          "WHERE ITEMNF.NFID = " & Formata_Dados(lngCONTRATOID, tpDados_Longo) & _
          " AND (ITEMNF.QUANTIDADE - ISNULL(ITENS_DEVOLVIDOS.QTD_DEVOL, 0) - ISNULL(ITENS_AVARIA.QTD_AVARIA, 0)) > 0 " & _
          " ORDER BY ITEMNF.PKID;"
  '
  strSql = "SELECT '', '', ITEMNF.ESTOQUEID, MIN(ESTOQUE.DESCRICAO) AS DESCRICAO, (SUM(ITEMNF.QUANTIDADE) - ISNULL(MIN(ITENS_DEVOLVIDOS.QTD_DEVOL), 0) - ISNULL(MIN(ITENS_AVARIA.QTD_AVARIA), 0)) AS QTD_REAL, '' As ALancar, ESTOQUE.VALORINDENIZACAO  "
  strSql = strSql & " FROM ITEMNF INNER JOIN ESTOQUE ON ESTOQUE.PKID = ITEMNF.ESTOQUEID " & _
          " INNER JOIN NF ON NF.PKID = ITEMNF.NFID " & _
          "LEFT JOIN (SELECT ITEMNF.ESTOQUEID, ISNULL(SUM(ITEMDEVOLUCAO.QUANTIDADE), 0) AS QTD_DEVOL " & _
          "         FROM ITEMDEVOLUCAO INNER JOIN ITEMNF ON ITEMNF.PKID = ITEMDEVOLUCAO.ITEMNFID " & _
          "         INNER JOIN NF ON NF.PKID = ITEMNF.NFID " & _
          "         WHERE NF.CONTRATOID = " & Formata_Dados(lngCONTRATOID, tpDados_Longo) & _
          "         GROUP BY ITEMNF.ESTOQUEID) AS ITENS_DEVOLVIDOS ON ITEMNF.ESTOQUEID = ITENS_DEVOLVIDOS.ESTOQUEID " & _
          "LEFT JOIN (SELECT ITEMNF.ESTOQUEID, ISNULL(SUM(ITEMAVARIA.QUANTIDADE), 0) AS QTD_AVARIA " & _
          "         FROM ITEMAVARIA INNER JOIN ITEMNF ON ITEMNF.PKID = ITEMAVARIA.ITEMNFID " & _
          "         INNER JOIN NF ON NF.PKID = ITEMNF.NFID " & _
          "         WHERE NF.CONTRATOID = " & Formata_Dados(lngCONTRATOID, tpDados_Longo) & _
          "         GROUP BY ITEMNF.ESTOQUEID) AS ITENS_AVARIA ON ITEMNF.ESTOQUEID = ITENS_AVARIA.ESTOQUEID " & _
          "WHERE NF.CONTRATOID = " & Formata_Dados(lngCONTRATOID, tpDados_Longo) & _
          " AND NF.STATUS IN ('F', 'S') " & _
          " GROUP BY ITEMNF.ESTOQUEID " & _
          " HAVING (SUM(ITEMNF.QUANTIDADE) - ISNULL(MIN(ITENS_DEVOLVIDOS.QTD_DEVOL), 0) - ISNULL(MIN(ITENS_AVARIA.QTD_AVARIA), 0)) > 0 " & _
          " ORDER BY ESTOQUE.DESCRICAO;"
  
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    ITEMNF_LINHASMATRIZ = objRs.RecordCount
  Else
    ITEMNF_LINHASMATRIZ = 0
  End If
  If Not objRs.EOF Then
    ReDim ITEMNF_Matriz(0 To ITEMNF_COLUNASMATRIZ - 1, 0 To ITEMNF_LINHASMATRIZ - 1)
  Else
    ReDim ITEMNF_Matriz(0 To ITEMNF_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se j� houver algum item
    For intI = 0 To ITEMNF_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda n�o se atingiu fim do recordset
        For intJ = 0 To ITEMNF_COLUNASMATRIZ - 1  'varre as colunas
          ITEMNF_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'pr�xima linha matriz
  End If
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Public Sub ITEMNFLANC_MontaMatriz()
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim intI          As Integer
  Dim intJ          As Integer
  Dim intRows       As Integer
  Dim clsGer        As busSisLoc.clsGeral
  '
  On Error GoTo trata


  Set clsGer = New busSisLoc.clsGeral
  '
  strSql = "SELECT ITEMAVARIA.PKID, ITEMAVARIA.ITEMNFID, ITEMAVARIA.AVARIAID, ITEMNF.ESTOQUEID, ESTOQUE.DESCRICAO, ITEMAVARIA.QUANTIDADE "
  strSql = strSql & " FROM ITEMAVARIA INNER JOIN ITEMNF ON ITEMNF.PKID = ITEMAVARIA.ITEMNFID " & _
          " INNER JOIN ESTOQUE ON ESTOQUE.PKID = ITEMNF.ESTOQUEID " & _
          "WHERE ITEMAVARIA.AVARIAID = " & Formata_Dados(lngAVARIAID, tpDados_Longo) & _
          " ORDER BY ITEMAVARIA.PKID;"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  '
  If Not objRs.EOF Then
    objRs.MoveFirst
    ITEMNFLANC_LINHASMATRIZ = objRs.RecordCount
  Else
    ITEMNFLANC_LINHASMATRIZ = 0
  End If
  If Not objRs.EOF Then
    ReDim ITEMNFLANC_Matriz(0 To ITEMNFLANC_COLUNASMATRIZ - 1, 0 To ITEMNFLANC_LINHASMATRIZ - 1)
  Else
    ReDim ITEMNFLANC_Matriz(0 To ITEMNFLANC_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se j� houver algum item
    For intI = 0 To ITEMNFLANC_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda n�o se atingiu fim do recordset
        For intJ = 0 To ITEMNFLANC_COLUNASMATRIZ - 1  'varre as colunas
          ITEMNFLANC_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'pr�xima linha matriz
  End If
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub


Private Sub cboAvaria_Click()
  On Error GoTo trata
  Dim objAvaria As busSisLoc.clsAvaria
  Dim objRs As ADODB.Recordset
  'Alterna para status de altera��o/inclus�o
  If cboAvaria.Text = "" Then
    Status = tpStatus_Incluir
    lngAVARIAID = 0
    Form_Load
    SetarFoco mskData
    Exit Sub
  End If
  Set objAvaria = New busSisLoc.clsAvaria
  Set objRs = objAvaria.ListarAvariaPeloSeq(lngCONTRATOID, _
                                            Left(cboAvaria.Text, 3))
  If objRs.EOF Then
    TratarErroPrevisto "Avaria " & cboAvaria.Text & " n�o cadastrada!"
    Status = tpStatus_Incluir
    lngAVARIAID = 0
    Form_Load
  Else
    Status = tpStatus_Alterar
    lngAVARIAID = objRs.Fields("PKID").Value
    Form_Load
  End If
  'Montar RecordSet
  ITEMNFLANC_COLUNASMATRIZ = grdNFOrigem.Columns.Count
  ITEMNFLANC_LINHASMATRIZ = 0
  ITEMNFLANC_MontaMatriz
  grdNFOrigem.Bookmark = Null
  grdNFOrigem.ReBind
  grdNFOrigem.ApproxCount = ITEMNFLANC_LINHASMATRIZ
  '
  SetarFoco mskData
  objRs.Close
  Set objRs = Nothing
  Set objAvaria = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub cmdCadastraItem_Click(Index As Integer)
  On Error GoTo trata
  TratarAssociacao Index + 1
  SetarFoco grdNF
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub TratarAssociacao(pIndice As Integer)
  On Error GoTo trata
  Dim objAvaria As busSisLoc.clsAvaria
  Dim intI          As Long
  Dim blnRet        As Boolean
  Dim blnSel        As Boolean
  Dim intExc        As Long
  Dim strSequencial As String
  '
  blnRet = False
  intExc = 0
  '
  Select Case pIndice
'''  Case 1 'Cadastrar Selecionados
'''    For intI = 0 To grdUnidade.SelBookmarks.Count - 1
'''      grdUnidade.Bookmark = CLng(grdUnidade.SelBookmarks.Item(intI))
'''      'Verificar se item possui estoue suficiente
'''      clsEstInter.AssociarUnidadeAoGrpEstoque grdUnidade.Columns("APARTAMENTOID").Text, lngGRUPOESTOQUEID
'''      blnRet = True
'''    Next
  Case 2 'Cadastrar Todos
    If ValidaCamposItemNFDestinoAll Then
      Exit Sub
    End If
    'Avaria
    Set objAvaria = New busSisLoc.clsAvaria
    If Status = tpStatus_Alterar Then
      'Alterar Contrato
      objAvaria.AlterarAvaria lngAVARIAID, _
                              mskData.Text, _
                              txtDescricao.Text
      '
    ElseIf Status = tpStatus_Incluir Then
      'Inserir Avaria
      objAvaria.InserirAvaria lngAVARIAID, _
                              lngCONTRATOID, _
                              strSequencial, _
                              mskData.Text, _
                              txtDescricao.Text
    End If
    '
    For intI = 0 To ITEMNF_LINHASMATRIZ - 1
      grdNF.Bookmark = CLng(intI)
      If grdNF.Columns(5) <> "" Then
        'Retorna valor para o estoque
        'objAvaria.AlterarEstoquePelaAvaria grdNF.Columns("ESTOQUEID").Text, _
                                           grdNF.Columns("Devolver").Text, _
                                           "INC"
        '
        objAvaria.InserirItemAvaria grdNF.Columns("ITEMNFID").Text, _
                                    lngAVARIAID, _
                                    grdNF.Columns("Avaria").Text, _
                                    grdNF.Columns("VALORINDENIZACAO").Text
        blnRet = True
      End If
    Next
    Set objAvaria = Nothing
    If Status = tpStatus_Incluir Then
      blnPrimeiraVez = True
      Status = tpStatus_Alterar
      Form_Load
      blnPrimeiraVez = False
      INCLUIR_VALOR_NO_COMBO txtSequencial.Text & "-" & mskData.Text, cboAvaria
    End If
    'blnFechar = True
    'Unload Me
  Case 3 'Retirar Selecionados
    'Avaria
    Set objAvaria = New busSisLoc.clsAvaria
    blnSel = False
    For intI = 0 To grdNFOrigem.SelBookmarks.Count - 1
      grdNFOrigem.Bookmark = CLng(grdNFOrigem.SelBookmarks.Item(intI))
      'retornar quantidade ao itens no estoque
      'objAvaria.AlterarEstoquePelaAvaria grdNFOrigem.Columns("ESTOQUEID").Text, _
                                               grdNFOrigem.Columns("Avaria").Text, _
                                               "RET"
      objAvaria.ExcluirItemAvaria grdNFOrigem.Columns("ITEMAVARIAID").Text
      
      blnSel = True
      blnRet = True
    Next
    Set objAvaria = Nothing
    If blnSel = False Then
      TratarErroPrevisto "Nenhuma pe�a selecionada para exclus�o.", "[frmUserAvariaInc.TratarAssociacao]"
    End If
  Case 4 'retirar Todos
    'Avaria
    Set objAvaria = New busSisLoc.clsAvaria
    For intI = 0 To ITEMNFLANC_LINHASMATRIZ - 1
      grdNFOrigem.Bookmark = CLng(intI)
      If IsNull(grdNFOrigem.Bookmark) Then grdNFOrigem.Bookmark = CLng(intI)
      
      'retornar quantidade ao itens no estoque
      'objAvaria.AlterarEstoquePelaAvaria grdNFOrigem.Columns("ESTOQUEID").Text, _
                                               grdNFOrigem.Columns("Avaria").Text, _
                                               "RET"
      objAvaria.ExcluirItemAvaria grdNFOrigem.Columns("ITEMAVARIAID").Text
      blnRet = True
    Next
    Set objAvaria = Nothing
  End Select
'''  '
'''  Set clsEstInter = Nothing
'''    '
  If blnRet Then
    VerificaStatusConsolicacao lngCONTRATOID
  End If
  If blnRet Then 'Houve Autera��o, Atualiza grids
    blnRetorno = True
    '
    ITEMNF_COLUNASMATRIZ = grdNF.Columns.Count
    ITEMNF_LINHASMATRIZ = 0
    ITEMNF_MontaMatriz
    grdNF.Bookmark = Null
    grdNF.ReBind
    grdNF.ApproxCount = ITEMNF_LINHASMATRIZ
    '
    'Montar RecordSet
    ITEMNFLANC_COLUNASMATRIZ = grdNFOrigem.Columns.Count
    ITEMNFLANC_LINHASMATRIZ = 0
    ITEMNFLANC_MontaMatriz
    grdNFOrigem.Bookmark = Null
    grdNFOrigem.ReBind
    grdNFOrigem.ApproxCount = ITEMNFLANC_LINHASMATRIZ
    '
  End If
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
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

Private Function ValidaCamposItemNFDestino(intLinha As Integer, intColuna As Integer) As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  '
  Select Case intColuna
  Case 5
    'Valid��o de quantidade
    If Not Valida_Moeda(grdNF.Columns(5), TpNaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
      strMsg = strMsg & "Quantidade inv�lida" & vbCrLf
    End If
    If strMsg = "" And grdNF.Columns(5) <> "" Then
      If CLng(grdNF.Columns(5)) = 0 Then
        strMsg = strMsg & "Quantidade lan�ada n�o pode ser zero" & vbCrLf
      End If
    End If
    If strMsg = "" And grdNF.Columns(5) <> "" Then
      If CLng(grdNF.Columns(5)) > CLng(grdNF.Columns(4)) Then
        strMsg = strMsg & "Quantidade n�o pode ser maior que a quantidade restante da pe�a na NF." & vbCrLf
      End If
    End If
  End Select
  
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserAvariaInc.ValidaCamposItemNFDestino]"
    ValidaCamposItemNFDestino = False
  Else
    ValidaCamposItemNFDestino = True
  End If
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserAvariaInc.ValidaCamposItemNFDestino]"
  ValidaCamposItemNFDestino = False
End Function

Private Function ValidaCamposItemNFDestinoAllSel() As Boolean
  On Error GoTo trata
  Dim blnRetorno            As Boolean
  Dim intRows               As Integer
  'Validar todas as linhas da matriz
  blnRetorno = False
  For intRows = 0 To ITEMNF_LINHASMATRIZ - 1
    grdNF.Bookmark = CLng(intRows)
    If grdNF.Columns(5).Text & "" <> "" Then
      blnRetorno = True
    End If
    If blnRetorno Then Exit For
  Next
  '
  grdNF.ReBind
  grdNF.SetFocus
  ValidaCamposItemNFDestinoAllSel = blnRetorno
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserAvariaInc.ValidaCamposItemNFDestinoAllSel]"
  ValidaCamposItemNFDestinoAllSel = False
End Function

Private Function ValidaCamposItemNFDestinoAll() As Boolean
  On Error GoTo trata
  Dim blnRetorno            As Boolean
  Dim blnLancouItem         As Boolean
  Dim intRows               As Integer
  'Validar todas as linhas da matriz
  blnRetorno = True
  blnLancouItem = False
  For intRows = 0 To ITEMNF_LINHASMATRIZ - 1
    grdNF.Bookmark = CLng(intRows)
    blnRetorno = ValidaCamposItemNFDestino(intRows, _
                                          5)
    If Not blnRetorno Then Exit For
    If blnRetorno Then
      If grdNF.Columns(5).Text & "" <> "" Then
        blnLancouItem = True
      End If
    End If
  Next
  '
  If blnLancouItem = False Then
    blnRetorno = False
    TratarErroPrevisto "Nenhum item lan�ado para esta NF.", "[frmUserAvariaInc.ValidaCamposItemNFDestinoAll]"
  End If
  If blnRetorno = True Then
    'Nenhum erro encontrado
    If Not Valida_Data(mskData, TpObrigatorio, False) Then
      TratarErroPrevisto "Data inv�lida.", "[frmUserAvariaInc.ValidaCamposItemNFDestinoAll]"
      blnRetorno = False
    End If
  End If
  If blnRetorno = False Then
    grdNF.ReBind
    grdNF.SetFocus
  End If
  ValidaCamposItemNFDestinoAll = Not blnRetorno
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserAvariaInc.ValidaCamposItemNFDestinoAll]"
  ValidaCamposItemNFDestinoAll = False
End Function

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Empresa
  LimparCampoTexto txtEmpresa
  LimparCampoTexto txtContrato
  LimparCampoMask mskData
  LimparCampoTexto txtDescricao
  LimparCampoTexto txtSequencial
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserEmpresaInc.LimparCampos]", _
            Err.Description
End Sub

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Montar RecordSet
    ITEMNF_COLUNASMATRIZ = grdNF.Columns.Count
    ITEMNF_LINHASMATRIZ = 0
    ITEMNF_MontaMatriz
    grdNF.Bookmark = Null
    grdNF.ReBind
    grdNF.ApproxCount = ITEMNF_LINHASMATRIZ
    'Montar RecordSet
    ITEMNFLANC_COLUNASMATRIZ = grdNFOrigem.Columns.Count
    ITEMNFLANC_LINHASMATRIZ = 0
    ITEMNFLANC_MontaMatriz
    grdNFOrigem.Bookmark = Null
    grdNFOrigem.ReBind
    grdNFOrigem.ApproxCount = ITEMNFLANC_LINHASMATRIZ
    '
    SetarFoco mskData
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserAvariaInc.Form_Activate]"
End Sub

Private Sub Form_Load()
On Error GoTo trata
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  Dim objNF         As busSisLoc.clsNF
  Dim objAvaria  As busSisLoc.clsAvaria
  '
  blnFechar = False
  blnRetorno = False
  AmpS
  Me.Height = 7950
  Me.Width = 12555
  CenterForm Me
  'blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, , , cmdCancelar
  '
'''  tabDetalhes_Click 0
  'LimparCampos
  LimparCampos
  'DEVOLU��O
  If blnPrimeiraVez Then
    strSql = "Select right('000' + convert(varchar(20), SEQUENCIAL), 3) + '-' + right('00' + CONVERT(VARCHAR(2), DAY(DATA)), 2) + '/' + right('00' + CONVERT(VARCHAR(2), MONTH(DATA)), 2) + '/' + right('0000' + CONVERT(VARCHAR(4), YEAR(DATA)), 4) " & _
            " FROM AVARIA " & _
            " WHERE NFID = " & Formata_Dados(lngCONTRATOID, tpDados_Longo) & _
            " ORDER BY SEQUENCIAL ASC"
    PreencheCombo cboAvaria, strSql, False, True
  End If
  'Pega Dados do Banco de dados
  Set objNF = New busSisLoc.clsNF
  Set objRs = objNF.SelecionarNFPeloPkid(lngCONTRATOID)
  '
  If Not objRs.EOF Then
    txtEmpresa.Text = objRs.Fields("NOME_EMPRESA").Value & ""
    txtContrato.Text = objRs.Fields("CONTRATO_NUMERO").Value & ""
    '
  End If
  objRs.Close
  Set objRs = Nothing
  Set objNF = Nothing
  '
  If Status = tpStatus_Incluir Then
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objAvaria = New busSisLoc.clsAvaria
    Set objRs = objAvaria.ListarAvaria(lngAVARIAID)
    '
    If Not objRs.EOF Then
      txtSequencial.Text = Format(objRs.Fields("SEQUENCIAL").Value, "000") & ""
      INCLUIR_VALOR_NO_MASK mskData, objRs.Fields("DATA").Value & "", TpMaskData
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objAvaria = Nothing
  End If
  '
  'Montar RecordSet
  ITEMNF_COLUNASMATRIZ = grdNF.Columns.Count
  ITEMNF_LINHASMATRIZ = 0
  ITEMNF_MontaMatriz
  grdNF.Bookmark = Null
  grdNF.ReBind
  grdNF.ApproxCount = ITEMNF_LINHASMATRIZ
  'Montar RecordSet
  ITEMNFLANC_COLUNASMATRIZ = grdNFOrigem.Columns.Count
  ITEMNFLANC_LINHASMATRIZ = 0
  ITEMNFLANC_MontaMatriz
  grdNFOrigem.Bookmark = Null
  grdNFOrigem.ReBind
  grdNFOrigem.ApproxCount = ITEMNFLANC_LINHASMATRIZ
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub



Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then Cancel = True
End Sub

Private Sub grdNF_BeforeRowColChange(Cancel As Integer)
  On Error GoTo trata
  If Not ValidaCamposItemNFDestino(grdNF.Row, _
                                  grdNF.Col) Then Cancel = True
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserAvariaInc.grdNF_BeforeRowColChange]"
End Sub

Private Sub grdNF_BeforeUpdate(Cancel As Integer)
  On Error GoTo trata
  'Atualiza Matriz
  ITEMNF_Matriz(5, grdNF.Row) = grdNF.Columns(5).Text
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserAvariaInc.grdNF_BeforeRowColChange]"
End Sub

Private Sub grdNF_UnboundReadDataEx( _
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
               Offset + intI, ITEMNF_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, ITEMNF_COLUNASMATRIZ, ITEMNF_LINHASMATRIZ, ITEMNF_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ITEMNF_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserAvariaInc.grdGeral_UnboundReadDataEx]"
End Sub

'''Private Sub grdNFOrigem_BeforeRowColChange(Cancel As Integer)
'''  On Error GoTo trata
'''  If Not ValidaCamposGrupoOrigem(grdNFOrigem.Row, _
'''                                 grdNFOrigem.Col) Then Cancel = True
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, "[frmUserAvariaInc.grdNFOrigem_BeforeRowColChange]"
'''End Sub
'''
'''Private Sub grdNFOrigem_BeforeUpdate(Cancel As Integer)
'''  On Error GoTo trata
'''  'Atualiza Matriz
'''  ITEMNF_Matriz(7, grdNFOrigem.Row) = grdNFOrigem.Columns(7).Text
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, "[frmUserAvariaInc.grdNFOrigem_BeforeRowColChange]"
'''End Sub


Private Sub grdNFOrigem_UnboundReadDataEx( _
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
               Offset + intI, ITEMNFLANC_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, ITEMNFLANC_COLUNASMATRIZ, ITEMNFLANC_LINHASMATRIZ, ITEMNFLANC_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ITEMNFLANC_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserAvariaInc.grdGeral_UnboundReadDataEx]"
End Sub

Private Sub grdNF_Validate(Cancel As Boolean)
  'Fazer valida��es ao retirar do grid ou clicar em outro controle
  On Error GoTo trata
  'Cancel = ValidaCamposItemNFDestinoAll
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserAvariaInc.grdNF_Validate]"
End Sub

Private Sub mskData_GotFocus()
  Seleciona_Conteudo_Controle mskData
End Sub
Private Sub mskData_LostFocus()
  Pintar_Controle mskData, tpCorContr_Normal
End Sub

