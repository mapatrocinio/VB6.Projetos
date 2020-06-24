VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserBoletoCredInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ítens do Boleto"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   13110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   13110
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6450
      Width           =   13110
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   1155
         Left            =   11340
         ScaleHeight     =   1095
         ScaleWidth      =   1605
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   90
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   90
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   6225
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   150
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   10980
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Ítens do Boleto"
      TabPicture(0)   =   "userBoletoCredInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Dados do ítem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   120
         TabIndex        =   14
         Top             =   390
         Width           =   12705
         Begin VB.TextBox txtUsuario 
            Enabled         =   0   'False
            Height          =   312
            Left            =   1290
            MaxLength       =   30
            TabIndex        =   0
            Top             =   315
            Width           =   1452
         End
         Begin VB.TextBox txtSenha 
            Enabled         =   0   'False
            Height          =   312
            IMEMode         =   3  'DISABLE
            Left            =   1290
            MaxLength       =   6
            PasswordChar    =   "*"
            TabIndex        =   1
            Top             =   690
            Width           =   1452
         End
         Begin VB.TextBox txtBoleto 
            Height          =   285
            Left            =   1290
            MaxLength       =   100
            TabIndex        =   2
            Text            =   "txtBoleto"
            Top             =   1080
            Width           =   2385
         End
         Begin MSMask.MaskEdBox mskTotal 
            Height          =   255
            Left            =   5370
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label5 
            Caption         =   "Total"
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   4110
            TabIndex        =   19
            Top             =   1050
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Senha"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   17
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Usuário"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   16
            Top             =   330
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Boleto"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   15
            Top             =   1095
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ítens a serem lançados"
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
         TabIndex        =   13
         Top             =   2070
         Width           =   12705
         Begin VB.CommandButton cmdCadastraItem 
            Caption         =   ">"
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   5820
            TabIndex        =   4
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdCadastraItem 
            Caption         =   ">>"
            Height          =   375
            Index           =   1
            Left            =   5820
            TabIndex        =   5
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton cmdCadastraItem 
            Caption         =   "<"
            Height          =   375
            Index           =   2
            Left            =   5820
            TabIndex        =   6
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton cmdCadastraItem 
            Caption         =   "<<"
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   5820
            TabIndex        =   7
            Top             =   1320
            Width           =   375
         End
         Begin TrueDBGrid60.TDBGrid grdBL 
            Height          =   3555
            Left            =   90
            OleObjectBlob   =   "userBoletoCredInc.frx":001C
            TabIndex        =   3
            Top             =   240
            Width           =   5745
         End
         Begin TrueDBGrid60.TDBGrid grdBLOrigem 
            Height          =   3555
            Left            =   6180
            OleObjectBlob   =   "userBoletoCredInc.frx":735B
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   240
            Width           =   6435
         End
      End
   End
End
Attribute VB_Name = "frmUserBoletoCredInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                 As tpStatus
Public lngBOLETOARRECID         As Long
Public lngFUNCIONARIOID         As Long
Public lngTURNOARRECEPESQ         As Long
Public blnRetorno               As Boolean
Public blnFechar                As Boolean
Public blnPrimeiraVez          As Boolean
Dim ITEMBL_COLUNASMATRIZ         As Long
Dim ITEMBL_LINHASMATRIZ          As Long
Private ITEMBL_Matriz()          As String

Dim ITEMBLLANC_COLUNASMATRIZ         As Long
Dim ITEMBLLANC_LINHASMATRIZ          As Long
Private ITEMBLLANC_Matriz()          As String

Private blnSairRow                As Boolean
Private blnSairGrid               As Boolean

Public Sub ITEMBL_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim objRsInt  As ADODB.Recordset
  Dim objRsConf As ADODB.Recordset
  Dim objRsFabricado As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGeral  As busSisMaq.clsGeral
  Dim strListaConfiguracaoId As String
  Dim vetColumns()
  '
  On Error GoTo trata

  Set objGeral = New busSisMaq.clsGeral
  '
  'Fabrica recordset
  Set objRsFabricado = New ADODB.Recordset
  objRsFabricado.Fields.Append "NUMERO", adInteger
  objRsFabricado.Fields.Append "MAQUINA", adVarChar, 50
  objRsFabricado.Fields.Append "CREDITO", adVarChar, 50
  objRsFabricado.Fields.Append "MEDICAO", adVarChar, 50
  objRsFabricado.Fields.Append "VALOR", adVarChar, 50
  objRsFabricado.Open
  'Monta Rs
  If lngBOLETOARRECID <> 0 Then
    'Criar vetor de colunas
    vetColumns = Array("NUMERO", "MAQUINA", "CREDITO", "MEDICAO", "VALOR")
  
    For intI = 1 To 10  'varre as linhas
      'Valida número já lançado
      Set objRsInt = New ADODB.Recordset
      strSql = "SELECT CREDITO.PKID "
      strSql = strSql & " FROM CREDITO " & _
              "WHERE CREDITO.NUMERO = " & Formata_Dados(intI, tpDados_Longo) & _
              " AND CREDITO.BOLETOARRECID = " & Formata_Dados(lngBOLETOARRECID, tpDados_Longo)
      '
      Set objRsInt = objGeral.ExecutarSQL(strSql)
      If objRsInt.EOF Then
        'Linha não lançada
        objRsFabricado.AddNew vetColumns, _
                              Array(intI, _
                                    "", _
                                    "", _
                                    "", _
                                    "")
        
      End If
      objRsInt.Close
      Set objRsInt = Nothing
    Next intI
  End If
  '
  objRsFabricado.Sort = "NUMERO"
  Set objRs = objRsFabricado
  Set objRsFabricado.ActiveConnection = Nothing
  '
  If Not objRs.EOF Then
    ITEMBL_LINHASMATRIZ = objRs.RecordCount
  Else
    ITEMBL_LINHASMATRIZ = 0
  End If
  If Not objRs.EOF Then
    ReDim ITEMBL_Matriz(0 To ITEMBL_COLUNASMATRIZ - 1, 0 To ITEMBL_LINHASMATRIZ - 1)
  Else
    ReDim ITEMBL_Matriz(0 To ITEMBL_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To ITEMBL_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To ITEMBL_COLUNASMATRIZ - 1  'varre as colunas
          ITEMBL_Matriz(intJ, intI) = objRs(intJ) & ""
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

Public Sub ITEMBLLANC_MontaMatriz()
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim intI          As Integer
  Dim intJ          As Integer
  Dim intRows       As Integer
  Dim clsGer        As busSisMaq.clsGeral
  Dim curValor      As Currency
  '
  On Error GoTo trata


  Set clsGer = New busSisMaq.clsGeral
  'Cálculo Valor total
  strSql = "SELECT SUM(CREDITO.VALORPAGO) AS TOTAL "
  strSql = strSql & " FROM CREDITO " & _
          "WHERE CREDITO.BOLETOARRECID = " & Formata_Dados(lngBOLETOARRECID, tpDados_Longo)
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  '
  curValor = 0
  If Not objRs.EOF Then
    curValor = IIf(IsNull(objRs.Fields("TOTAL").Value), 0, objRs.Fields("TOTAL").Value)
  End If
  objRs.Close
  INCLUIR_VALOR_NO_MASK mskTotal, curValor, TpMaskMoeda
  '
  strSql = "SELECT CREDITO.PKID, CREDITO.BOLETOARRECID, CREDITO.NUMERO, EQUIPAMENTO.NUMERO, CREDITO.MEDICAO, CREDITO.VALORPAGO, (ISNULL(CREDITO.VALORPAGO,0) / ISNULL(CREDITO.COEFICIENTE,0)) AS CREDITO, CREDITO.DATA "
  strSql = strSql & " FROM CREDITO " & _
          " INNER JOIN BOLETOARREC ON BOLETOARREC.PKID = CREDITO.BOLETOARRECID " & _
          " INNER JOIN MAQUINA ON MAQUINA.PKID = CREDITO.MAQUINAID " & _
          " INNER JOIN EQUIPAMENTO ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
          "WHERE CREDITO.BOLETOARRECID = " & Formata_Dados(lngBOLETOARRECID, tpDados_Longo) & _
          " ORDER BY CREDITO.NUMERO;"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  '
  If Not objRs.EOF Then
    objRs.MoveFirst
    ITEMBLLANC_LINHASMATRIZ = objRs.RecordCount
  Else
    ITEMBLLANC_LINHASMATRIZ = 0
  End If
  If Not objRs.EOF Then
    ReDim ITEMBLLANC_Matriz(0 To ITEMBLLANC_COLUNASMATRIZ - 1, 0 To ITEMBLLANC_LINHASMATRIZ - 1)
  Else
    ReDim ITEMBLLANC_Matriz(0 To ITEMBLLANC_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To ITEMBLLANC_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To ITEMBLLANC_COLUNASMATRIZ - 1  'varre as colunas
          ITEMBLLANC_Matriz(intJ, intI) = objRs(intJ) & ""
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


'''Private Sub cboBoletoCred_Click()
'''  On Error GoTo trata
'''  Dim objBoletoCred As busSisMaq.clsBoletoCred
'''  Dim objRs As ADODB.Recordset
'''  'Alterna para status de alteração/inclusão
'''  If cboBoletoCred.Text = "" Then
'''    Status = tpStatus_Incluir
'''    lngBOLETOARRECID = 0
'''    Form_Load
'''    'Montar RecordSet
'''    ITEMBLLANC_COLUNASMATRIZ = grdBLOrigem.Columns.Count
'''    ITEMBLLANC_LINHASMATRIZ = 0
'''    ITEMBLLANC_MontaMatriz
'''    grdBLOrigem.Bookmark = Null
'''    grdBLOrigem.ReBind
'''    grdBLOrigem.ApproxCount = ITEMBLLANC_LINHASMATRIZ
'''    '
'''    SetarFoco txtNFCliente
'''    Exit Sub
'''  End If
'''  Set objBoletoCred = New busSisMaq.clsBoletoCred
'''  Set objRs = objBoletoCred.ListarBoletoCredPeloSeq(lngCONTRATOID, _
'''                                                  lngOBRAID, _
'''                                                  Left(cboBoletoCred.Text, 3))
'''  If objRs.EOF Then
'''    TratarErroPrevisto "Devolução " & cboBoletoCred.Text & " não cadastrada!"
'''    Status = tpStatus_Incluir
'''    lngBOLETOARRECID = 0
'''    Form_Load
'''  Else
'''    Status = tpStatus_Alterar
'''    lngBOLETOARRECID = objRs.Fields("PKID").Value
'''    Form_Load
'''  End If
'''  'Montar RecordSet
'''  ITEMBL_COLUNASMATRIZ = grdBL.Columns.Count
'''  ITEMBL_LINHASMATRIZ = 0
'''  ITEMBL_MontaMatriz
'''  grdBL.Bookmark = Null
'''  grdBL.ReBind
'''  grdBL.ApproxCount = ITEMBL_LINHASMATRIZ
'''  'Montar RecordSet
'''  ITEMBLLANC_COLUNASMATRIZ = grdBLOrigem.Columns.Count
'''  ITEMBLLANC_LINHASMATRIZ = 0
'''  ITEMBLLANC_MontaMatriz
'''  grdBLOrigem.Bookmark = Null
'''  grdBLOrigem.ReBind
'''  grdBLOrigem.ApproxCount = ITEMBLLANC_LINHASMATRIZ
'''  '
'''  SetarFoco txtNFCliente
'''  objRs.Close
'''  Set objRs = Nothing
'''  Set objBoletoCred = Nothing
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''  AmpN
'''End Sub
'''
Private Sub cmdCadastraItem_Click(Index As Integer)
  On Error GoTo trata
  TratarAssociacao Index + 1
  SetarFoco grdBL
  grdBL.Col = 1
  If grdBL.Row > -1 Then
    grdBL.Row = 0
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub TratarAssociacao(pIndice As Integer)
  On Error GoTo trata
  Dim objCredito       As busSisMaq.clsCredito
  Dim objGeral        As busSisMaq.clsGeral
  Dim lngMAQUINAID    As Long
  Dim curCOEFICIENTE  As Currency
  Dim strCOEFICIENTE  As String
  Dim strData         As String
'''  Dim objItemNF     As busSisMaq.clsItemNF
  Dim intI          As Long
  Dim blnRet        As Boolean
  Dim blnSel        As Boolean
  Dim intExc        As Long
  Dim strSql        As String
'''  Dim strSequencial As String
  Dim objRs         As ADODB.Recordset
'''  Dim lngQUANTIDADE As Long
'''  Dim lngQtdEstoque As Long
'''  Dim lngQtdALanc   As Long
'''  Dim lngQtdALancAva  As Long
'''  Dim lngQTDAVARIA  As Long
'''  Dim strQTDNF      As String

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
    If ValidaCamposItemBLDestinoAllSel Then
      Exit Sub
    End If
    '
    Set objCredito = New busSisMaq.clsCredito
    Set objGeral = New busSisMaq.clsGeral
    strData = Format(Now, "DD/MM/YYYY hh:mm")
    For intI = 0 To ITEMBL_LINHASMATRIZ - 1
      grdBL.Bookmark = CLng(intI)
      If grdBL.Columns("Máquina").Text & "" <> "" And _
          grdBL.Columns("Crédito").Text & "" <> "" And _
          grdBL.Columns("Medição").Text & "" <> "" And _
          grdBL.Columns("Valor").Text & "" <> "" Then
        'Propósito: Retornar todos os ítens
        '
        lngMAQUINAID = 0
        '
        strSql = "SELECT MAQUINA.PKID, ISNULL(EQUIPAMENTO.COEFICIENTE, ISNULL(SERIE.COEFICIENTE,0)) AS COEF FROM EQUIPAMENTO " & _
              " INNER JOIN MAQUINA ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
              " INNER JOIN SERIE ON SERIE.PKID = EQUIPAMENTO.SERIEID " & _
              " WHERE EQUIPAMENTO.NUMERO = " & Formata_Dados(grdBL.Columns("Máquina").Text, tpDados_Texto) & _
              " AND MAQUINA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
              " AND EQUIPAMENTO.STATUS = " & Formata_Dados("A", tpDados_Texto)
        Set objRs = objGeral.ExecutarSQL(strSql)
        If Not objRs.EOF Then
          lngMAQUINAID = objRs.Fields("PKID").Value
          curCOEFICIENTE = objRs.Fields("COEF").Value
          strCOEFICIENTE = Format(curCOEFICIENTE, "###,##0.0000")
        End If
        objRs.Close
        Set objRs = Nothing
        
        objCredito.InserirCredito lngMAQUINAID, _
                                lngBOLETOARRECID, _
                                grdBL.Columns("Nro.").Text & "", _
                                grdBL.Columns("Medição").Text & "", _
                                grdBL.Columns("Valor").Text & "", _
                                strCOEFICIENTE, _
                                strData, _
                                RetornaCodTurnoCorrente, _
                                grdBL.Columns("Crédito").Text & ""
                                
        blnRet = True
        'Verifica consolidação
        VerificaStatusConsolicacaoArrec lngBOLETOARRECID
        'Indica se quantidade restante fechou
      End If
    Next
    Set objCredito = Nothing
    Set objGeral = Nothing
    '
    blnFechar = True
    Unload Me
  Case 3 'Retirar Selecionados
    'Devolução
    'Pede liberação do gerente
    frmUserLoginLibera.lngFUNCIONARIOID = 0
    frmUserLoginLibera.strNivel = "'GER','ADM'"
    frmUserLoginLibera.Show vbModal
    If Len(Trim(gsNomeUsuLib)) = 0 Then
      TratarErroPrevisto "É necessário confirmação do gerente para executar esta ação.", "cmdConfirmar_Click"
      Exit Sub
    End If
    '
    Set objCredito = New busSisMaq.clsCredito
    blnSel = False
    For intI = 0 To grdBLOrigem.SelBookmarks.Count - 1
      grdBLOrigem.Bookmark = CLng(grdBLOrigem.SelBookmarks.Item(intI))
      'excluir debito
      objCredito.ExcluirCredito grdBLOrigem.Columns("CREDITOID").Text
      'Verifica consolidação
      VerificaStatusConsolicacaoArrec lngBOLETOARRECID

      blnSel = True
      blnRet = True
    Next
    Set objCredito = Nothing
    If blnSel = False Then
      TratarErroPrevisto "Nenhum ítem do boleto selecionado para exclusão.", "[frmUserBoletoCredInc.TratarAssociacao]"
    End If
'''  Case 4 'retirar Todos
'''    'Devolução
'''    Set objBoletoCred = New busSisMaq.clsBoletoCred
'''    For intI = 0 To ITEMBLLANC_LINHASMATRIZ - 1
'''      grdBLOrigem.Bookmark = CLng(intI)
'''      If IsNull(grdBLOrigem.Bookmark) Then grdBLOrigem.Bookmark = CLng(intI)
'''
'''      'retornar quantidade ao itens no estoque
'''      objBoletoCred.AlterarEstoquePelaBoletoCred grdBLOrigem.Columns("ESTOQUEID").Text, _
'''                                               grdBLOrigem.Columns("Devol.").Text, _
'''                                               "RET"
'''      objBoletoCred.ExcluirItemDeVolucao grdBLOrigem.Columns("ITEMDEVOLUCAOID").Text
'''      'Verifica consolidação
'''      VerificaStatusConsolicacao grdBLOrigem.Columns("NFID").Text
'''      blnRet = True
'''    Next
'''    Set objBoletoCred = Nothing
  End Select
'''  '
'''  Set clsEstInter = Nothing
'''    '
  If blnRet Then 'Houve Auteração, Atualiza grids
    blnRetorno = True
    '
    ITEMBL_COLUNASMATRIZ = grdBL.Columns.Count
    ITEMBL_LINHASMATRIZ = 0
    ITEMBL_MontaMatriz
    grdBL.Bookmark = Null
    grdBL.ReBind
    grdBL.ApproxCount = ITEMBL_LINHASMATRIZ
    '
    'Montar RecordSet
    ITEMBLLANC_COLUNASMATRIZ = grdBLOrigem.Columns.Count
    ITEMBLLANC_LINHASMATRIZ = 0
    ITEMBLLANC_MontaMatriz
    grdBLOrigem.Bookmark = Null
    grdBLOrigem.ReBind
    grdBLOrigem.ApproxCount = ITEMBLLANC_LINHASMATRIZ
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
'''Private Function ValidaCamposItemNFGeral() As Boolean
'''  On Error GoTo trata
'''  Dim strMsg                As String
'''  Dim blnSetarFocoControle  As Boolean
'''  '
'''  blnSetarFocoControle = True
'''  '
'''  If grdBL.Columns("Informado").Text = "" And grdBL.Columns("Avaria").Text = "" And grdBL.Columns("Recebido").Text = "" Then
'''    'Não lançou item
'''    ValidaCamposItemNFGeral = True
'''    Exit Function
'''  End If
'''  'Validção de quantidade Informada
'''  If Not Valida_Moeda(grdBL.Columns("Informado"), TpNaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
'''    strMsg = strMsg & "Quantidade informada inválida" & vbCrLf
'''  End If
'''  'Validção de quantidade avaria
'''  If Not Valida_Moeda(grdBL.Columns("Avaria"), TpNaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
'''    strMsg = strMsg & "Quantidade avaria inválida" & vbCrLf
'''  End If
'''  'Validção de quantidade avaria
'''  If Not Valida_Moeda(grdBL.Columns("Recebido"), TpObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
'''    strMsg = strMsg & "Quantidade recebido inválida" & vbCrLf
'''  End If
'''  If strMsg = "" Then
'''    'Avaria e Recebido não Recebido
'''    If grdBL.Columns("Avaria").Text = "" And grdBL.Columns("Recebido").Text = "" Then
'''      strMsg = strMsg & "Informar a quantidade de avaria ou recebido na NFSF." & vbCrLf
'''      SetarFoco grdBL
'''    End If
'''  End If
'''  If strMsg = "" Then
'''    'Quantidade informada > quantidade restante
'''    If (CLng(IIf(grdBL.Columns("Recebido").Text & "" = "", "0", grdBL.Columns("Recebido").Text))) > CLng(grdBL.Columns("Restante").Text) Then
'''      strMsg = strMsg & "Quantidade informada não pode ser maior que a quantidade restante da peça na NFSF." & vbCrLf
'''      SetarFoco grdBL
'''    End If
'''  End If
'''  If strMsg = "" Then
'''    'Quantidade informada > quantidade restante
'''    If (CLng(IIf(grdBL.Columns("Avaria").Text & "" = "", "0", grdBL.Columns("Avaria").Text))) > CLng(IIf(grdBL.Columns("Recebido").Text & "" = "", "0", grdBL.Columns("Recebido").Text)) Then
'''      strMsg = strMsg & "Quantidade de avaria não pode ser maior que a quantidade recebida da peça na NFSF." & vbCrLf
'''      SetarFoco grdBL
'''    End If
'''  End If
'''  '
'''  If Len(strMsg) <> 0 Then
'''    TratarErroPrevisto strMsg, "[frmUserBoletoCredInc.ValidaCamposItemNFGeral]"
'''    ValidaCamposItemNFGeral = False
'''  Else
'''    ValidaCamposItemNFGeral = True
'''  End If
'''  Exit Function
'''trata:
'''  TratarErro Err.Number, _
'''             Err.Description, _
'''             "[frmUserBoletoCredInc.ValidaCamposItemNFGeral]"
'''  ValidaCamposItemNFGeral = False
'''End Function
'''
Private Function ValidaCamposItemBLDestino(intLinha As Integer) As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  Dim strSql                As String
  Dim objRs                 As ADODB.Recordset
  Dim objGeral              As busSisMaq.clsGeral
  Dim lngMAQUINAID          As Long
  Dim curCOEFICIENTE        As Currency
  Dim curCREDITO            As Currency
  Dim curVALORCALCINI       As Currency
  Dim curVALORCALCFIM       As Currency
  Dim curValor              As Currency
  Dim curMEDICAO            As Currency
  Dim curMEDICAOANT         As Currency
  Dim strMEDICAODESCR       As String
  Dim strDATAINI            As String
  Dim strDATAFIM            As String
  Dim datTurnoCorrente      As Date
  Dim lngCorTurnoCorrente   As Long
  Dim lngPeriodo            As Long
  '
  blnSetarFocoControle = True
  '
  'Validção da Máquina
  If Not Valida_Moeda(grdBL.Columns("Máquina"), TpObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
    strMsg = strMsg & "Máquina informada inválida na linha " & intLinha + 1 & vbCrLf
  End If
  'Validção do Crédito
  If Not Valida_Moeda(grdBL.Columns("Crédito"), TpObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
    strMsg = strMsg & "Crédito inválido na linha " & intLinha + 1 & vbCrLf
  End If
  'Validção da Medição
  If Not Valida_Moeda(grdBL.Columns("Medição"), TpNaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
    strMsg = strMsg & "Medição inválida na linha " & intLinha + 1 & vbCrLf
  End If
  'Validção do Valor
  If Not Valida_Moeda(grdBL.Columns("Valor"), TpObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
    strMsg = strMsg & "Valor inválido na linha " & intLinha + 1 & vbCrLf
  End If
  If Len(strMsg) = 0 Then
    Set objGeral = New busSisMaq.clsGeral
    'strSql = "SELECT EQUIPAMENTO.PKID FROM EQUIPAMENTO " & _
          " INNER JOIN MAQUINA ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
          " WHERE EQUIPAMENTO.NUMERO = " & Formata_Dados(grdBL.Columns("Máquina").Text, tpDados_Texto) & _
          " AND MAQUINA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
          " AND EQUIPAMENTO.STATUS = " & Formata_Dados("A", tpDados_Texto)
    strSql = "SELECT MAQUINA.PKID, ISNULL(EQUIPAMENTO.COEFICIENTE, ISNULL(SERIE.COEFICIENTE,0)) AS COEF FROM EQUIPAMENTO " & _
          " INNER JOIN MAQUINA ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
          " INNER JOIN SERIE ON SERIE.PKID = EQUIPAMENTO.SERIEID " & _
          " WHERE EQUIPAMENTO.NUMERO = " & Formata_Dados(grdBL.Columns("Máquina").Text, tpDados_Texto) & _
          " AND MAQUINA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
          " AND EQUIPAMENTO.STATUS = " & Formata_Dados("A", tpDados_Texto)
          
    Set objRs = objGeral.ExecutarSQL(strSql)
    If objRs.EOF Then
      strMsg = strMsg & "Equipamento não cadastrado na linha " & intLinha + 1 & vbCrLf
    Else
      lngMAQUINAID = objRs.Fields("PKID").Value
      curCOEFICIENTE = objRs.Fields("COEF").Value
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
  End If
  If Len(strMsg) = 0 Then
    'Validações avançadas
    curCREDITO = CCur(grdBL.Columns("Crédito").Text)
    curVALORCALCINI = curCREDITO * curCOEFICIENTE
    'curVALORCALCFIM = curCREDITO / curCOEFICIENTE
    curVALORCALCFIM = curCREDITO
    '
    'Validação de valor
    curValor = CCur(grdBL.Columns("Valor").Text)
    strMEDICAODESCR = ""
    If curVALORCALCINI <> curValor Then
      strMsg = strMsg & "Valor lançado : " & Format(curValor, "###,##0.00") & vbCrLf
      strMsg = strMsg & "Difere do valor calculado : " & Format(curVALORCALCINI, "###,##0.00") & vbCrLf & vbCrLf
      strMsg = strMsg & "Favor informar o valor correto na linha " & intLinha + 1 & vbCrLf
    End If
    If strMsg = "" Then
      'Valida valor calculado final
      Set objGeral = New busSisMaq.clsGeral
      'strSql = "SELECT TOP 1 ISNULL(DEBITO.MEDICAO, 0) AS MED FROM DEBITO " & _
            " WHERE DEBITO.DATA < " & Formata_Dados(Format(Now, "DD/MM/YYYY hh:mm"), tpDados_DataHora) & _
            " AND DEBITO.MAQUINAID = " & Formata_Dados(lngMAQUINAID, tpDados_Longo) & _
            " ORDER BY DEBITO.DATA DESC "
      strSql = "SELECT TOP 1 ISNULL(CREDITO.MEDICAO, 0) AS MED FROM CREDITO " & _
            " WHERE CREDITO.MAQUINAID = " & Formata_Dados(lngMAQUINAID, tpDados_Longo) & _
            " AND CREDITO.TURNOID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo) & _
            " ORDER BY CREDITO.PKID DESC "
      Set objRs = objGeral.ExecutarSQL(strSql)
      curMEDICAOANT = 0
      If Not objRs.EOF Then
        If objRs.Fields("MED").Value <> 0 Then
          curMEDICAOANT = objRs.Fields("MED").Value
          strMEDICAODESCR = "Medição Anterior"
        End If
      End If
      objRs.Close
      Set objRs = Nothing
      'NOVO
      If curMEDICAOANT = 0 Then
        'Caso a medição anterior não seja encontrada, procurar na leitura especial inicial
        lngCorTurnoCorrente = RetornaCodTurnoCorrente(datTurnoCorrente, _
                                                      lngPeriodo)
        strDATAINI = Format(Day(datTurnoCorrente), "00") & "/" & Format(Month(datTurnoCorrente), "00") & "/" & Year(datTurnoCorrente)
        strDATAFIM = Format(Day(DateAdd("D", 1, datTurnoCorrente)), "00") & "/" & Format(Month(DateAdd("D", 1, datTurnoCorrente)), "00") & "/" & Year(DateAdd("D", 1, datTurnoCorrente))
        strSql = "SELECT ISNULL(LEITURAMAQUINAFECHA.MEDICAOENTRADA, 0) AS MED FROM LEITURAFECHA " & _
              " INNER JOIN LEITURAMAQUINAFECHA ON LEITURAFECHA.PKID = LEITURAMAQUINAFECHA.LEITURAFECHAID " & _
              " WHERE LEITURAMAQUINAFECHA.MAQUINAID = " & Formata_Dados(lngMAQUINAID, tpDados_Longo) & _
              " AND LEITURAFECHA.DATA >= " & Formata_Dados(strDATAINI, tpDados_DataHora) & _
              " AND LEITURAFECHA.DATA < " & Formata_Dados(strDATAFIM, tpDados_DataHora) & _
              " AND LEITURAFECHA.STATUS = " & Formata_Dados("I", tpDados_Texto) & _
              " AND LEITURAFECHA.PERIODOID = " & Formata_Dados(lngPeriodo, tpDados_Longo) & _
              " ORDER BY LEITURAMAQUINAFECHA.PKID DESC "
        Set objRs = objGeral.ExecutarSQL(strSql)
        curMEDICAOANT = 0
        If Not objRs.EOF Then
          If objRs.Fields("MED").Value <> 0 Then
            curMEDICAOANT = objRs.Fields("MED").Value
            strMEDICAODESCR = "Leitura Especial"
          End If
        End If
      End If
      '
      Set objGeral = Nothing
      '
      'Verifica se nenhuma medição foi lançada
      If curMEDICAOANT = 0 Then
        strMsg = strMsg & "Leitura Especial Inicial não lançada para a linha " & intLinha + 1 & vbCrLf
      Else
        curMEDICAO = CCur(grdBL.Columns("Medição").Text)
        If (curVALORCALCFIM + curMEDICAOANT) <> curMEDICAO Then
          strMsg = strMsg & "Medição calculada : " & Format(curVALORCALCFIM, "###,##0") & vbCrLf
          strMsg = strMsg & "Mais a Medição anterior : " & Format(curMEDICAOANT, "###,##0") & vbCrLf
          strMsg = strMsg & "Difere da medição lançada : " & Format(curMEDICAO, "###,##0") & vbCrLf & vbCrLf
          strMsg = strMsg & "Favor informar a medição correta na linha " & intLinha + 1 & vbCrLf
        End If
      End If
    End If
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserBoletoCredInc.ValidaCamposItemBLDestino]"
    ValidaCamposItemBLDestino = False
  Else
    ValidaCamposItemBLDestino = True
  End If
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserBoletoCredInc.ValidaCamposItemBLDestino]"
  ValidaCamposItemBLDestino = False
End Function

Private Function ValidaCamposItemBLDestinoAllSel() As Boolean
  On Error GoTo trata
  Dim blnRetorno            As Boolean
  Dim blnSelItem            As Boolean
  Dim blnEncontrouErro      As Boolean
  Dim intRows               As Integer
  'Validar todas as linhas da matriz
  blnSelItem = False
  blnEncontrouErro = False
  blnRetorno = True
  
  For intRows = 0 To ITEMBL_LINHASMATRIZ - 1
    grdBL.Bookmark = CLng(intRows)
    '
    If grdBL.Columns(1).Text & "" <> "" Or grdBL.Columns(2).Text & "" <> "" Or grdBL.Columns(3).Text & "" <> "" Then
      'Selecionou um item
      blnSelItem = True
    End If
    If grdBL.Columns(1).Text & "" <> "" Or _
      grdBL.Columns(2).Text & "" <> "" Or _
      grdBL.Columns(3).Text & "" <> "" Then
      If Not ValidaCamposItemBLDestino(grdBL.Row) Then
        blnEncontrouErro = True
      End If
      If blnEncontrouErro = True Then Exit For
    End If
  Next
  '
  If blnSelItem = True And blnEncontrouErro = False Then
    blnRetorno = False
  End If
  If blnSelItem = False And blnEncontrouErro = False Then
    TratarErroPrevisto "Entre com ao menos 1 ítem do Boleto", "[frmUserBoletoCredInc.ValidaCamposItemBLDestinoAllSel]"
  End If
  grdBL.ReBind
  grdBL.SetFocus
  ValidaCamposItemBLDestinoAllSel = blnRetorno
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserBoletoCredInc.ValidaCamposItemBLDestinoAllSel]"
  ValidaCamposItemBLDestinoAllSel = False
End Function

'''Private Function ValidaCamposItemBLDestinoAll() As Boolean
'''  On Error GoTo trata
'''  Dim blnRetorno            As Boolean
'''  Dim blnLancouItem         As Boolean
'''  Dim intRows               As Integer
'''  'Validar todas as linhas da matriz
'''  blnRetorno = True
'''  blnLancouItem = False
'''  For intRows = 0 To ITEMBL_LINHASMATRIZ - 1
'''    grdBL.Bookmark = CLng(intRows)
'''    blnRetorno = ValidaCamposItemNFGeral
'''    If Not blnRetorno Then Exit For
'''    If blnRetorno Then
'''      If grdBL.Columns("Avaria").Text & "" <> "" Or grdBL.Columns("Recebido").Text & "" <> "" Then
'''        blnLancouItem = True
'''      End If
'''    End If
'''  Next
'''  '
'''  If blnLancouItem = False Then
'''    blnRetorno = False
'''    TratarErroPrevisto "Nenhum item lançado para esta NF.", "[frmUserBoletoCredInc.ValidaCamposItemBLDestinoAll]"
'''  End If
'''  If blnRetorno = True Then
'''    'Nenhum erro encontrado
'''    If Not Valida_String(txtNFCliente, TpObrigatorio, False) Then
'''      TratarErroPrevisto "Número NF Cliente inválido.", "[frmUserBoletoCredInc.ValidaCamposItemBLDestinoAll]"
'''      blnRetorno = False
'''    End If
'''  End If
'''  If blnRetorno = True Then
'''    If Not Valida_Data(mskData, TpObrigatorio, False) Then
'''      TratarErroPrevisto "Data inválida.", "[frmUserBoletoCredInc.ValidaCamposItemBLDestinoAll]"
'''      blnRetorno = False
'''    End If
'''  End If
'''  If blnRetorno = True Then
'''    If Not Valida_Data(mskDataEmissao, TpObrigatorio, False) Then
'''      TratarErroPrevisto "Data de emissão inválida.", "[frmUserBoletoCredInc.ValidaCamposItemBLDestinoAll]"
'''      blnRetorno = False
'''    End If
'''  End If
'''  If blnRetorno = True Then
'''    If Not Valida_Data(mskDataBoletoCred, TpObrigatorio, False) Then
'''      TratarErroPrevisto "Data de devolução inválida.", "[frmUserBoletoCredInc.ValidaCamposItemBLDestinoAll]"
'''      blnRetorno = False
'''    End If
'''  End If
'''  If blnRetorno = False Then
'''    grdBL.ReBind
'''    grdBL.SetFocus
'''  End If
'''  ValidaCamposItemBLDestinoAll = Not blnRetorno
'''  Exit Function
'''trata:
'''  TratarErro Err.Number, _
'''             Err.Description, _
'''             "[frmUserBoletoCredInc.ValidaCamposItemBLDestinoAll]"
'''  ValidaCamposItemBLDestinoAll = False
'''End Function
'''
Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'BoletoCred
  LimparCampoTexto txtBoleto
  LimparCampoTexto txtUsuario
  LimparCampoTexto txtSenha
  LimparCampoMask mskTotal
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserBoletoCredInc.LimparCampos]", _
            Err.Description
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  '
  If Me.ActiveControl.Name <> "grdBL" Then
    If KeyAscii = 13 Then
      SendKeys "{tab}"
    End If
  Else
    If KeyAscii = 13 And grdBL.Row <> -1 Then
      If grdBL.Col = 4 Then
        blnSairRow = True
        blnSairGrid = True
        '
        ITEMBL_Matriz(1, grdBL.Row) = grdBL.Columns(1).Text
        ITEMBL_Matriz(2, grdBL.Row) = grdBL.Columns(2).Text
        ITEMBL_Matriz(3, grdBL.Row) = grdBL.Columns(3).Text
        ITEMBL_Matriz(4, grdBL.Row) = grdBL.Columns(4).Text
        '
        'Para cada linha verifica se está em branco, se sim simula o ENTER
        If (grdBL.Columns("Máquina").Text & "" = "" _
           And grdBL.Columns("Medição").Text & "" = "" _
           And grdBL.Columns("Crédito").Text & "" = "" _
           And grdBL.Columns("Valor").Text & "" = "") Or ((grdBL.Row + 1) = ITEMBL_LINHASMATRIZ) Then
          cmdCadastraItem_Click 1
        Else
          grdBL.Col = 1
          grdBL.Row = grdBL.Row + 1
        End If
        blnSairRow = False
        blnSairGrid = False
        '
      Else
        grdBL.Col = grdBL.Col + 1
      End If
    End If
  End If
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserBoletoCredInc.Form_Activate]"
End Sub



Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Montar RecordSet
    ITEMBL_COLUNASMATRIZ = grdBL.Columns.Count
    ITEMBL_LINHASMATRIZ = 0
    ITEMBL_MontaMatriz
    grdBL.Bookmark = Null
    grdBL.ReBind
    grdBL.ApproxCount = ITEMBL_LINHASMATRIZ
    'Montar RecordSet
    ITEMBLLANC_COLUNASMATRIZ = grdBLOrigem.Columns.Count
    ITEMBLLANC_LINHASMATRIZ = 0
    ITEMBLLANC_MontaMatriz
    grdBLOrigem.Bookmark = Null
    grdBLOrigem.ReBind
    grdBLOrigem.ApproxCount = ITEMBLLANC_LINHASMATRIZ
    '
    SetarFoco txtBoleto
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserBoletoCredInc.Form_Activate]"
End Sub

Private Sub Form_Load()
On Error GoTo trata
  '
  blnFechar = False
  blnRetorno = False
  AmpS
  Me.Height = 7035
  Me.Width = 13200
  CenterForm Me
  blnPrimeiraVez = True
  '
  blnSairRow = False
  blnSairGrid = False
  
  LerFiguras Me, tpBmp_Vazio, , , cmdCancelar
  '
'''  tabDetalhes_Click 0
  'LimparCampos
  LimparCampos
  '
  If Status = tpStatus_Incluir Then
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
'''    Set objBoletoCred = New busSisMaq.clsBoletoCred
'''    Set objRs = objBoletoCred.ListarBoletoCred(lngBOLETOARRECID)
'''    '
'''    If Not objRs.EOF Then
'''      txtSequencial.Text = Format(objRs.Fields("SEQUENCIAL").Value, "000") & ""
'''      INCLUIR_VALOR_NO_MASK mskData, objRs.Fields("DATA").Value & "", TpMaskData
'''      INCLUIR_VALOR_NO_MASK mskDataEmissao, objRs.Fields("DATAEMISSAO").Value & "", TpMaskData
'''      INCLUIR_VALOR_NO_MASK mskDataBoletoCred, objRs.Fields("DATADEVOLUCAO").Value & "", TpMaskData
'''      txtNFCliente.Text = Format(objRs.Fields("NUMERONF").Value, "000") & ""
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
'''    '
'''    Set objBoletoCred = Nothing
  End If
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub



'''Private Sub Form_Unload(Cancel As Integer)
'''  If Not blnFechar Then Cancel = True
'''End Sub
'''
'''Private Sub grdBL_BeforeRowColChange(Cancel As Integer)
'''  On Error GoTo trata
'''  'If Not ValidaCamposItemBLDestino(grdBL.Row, _
'''                                  grdBL.Col) Then Cancel = True


'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, "[frmUserBoletoCredInc.grdBL_BeforeRowColChange]"
'''End Sub
'''
Private Sub grdBL_BeforeUpdate(Cancel As Integer)
  On Error GoTo trata
  'Atualiza Matriz
  If blnSairRow = False Then
    ITEMBL_Matriz(1, grdBL.Row) = grdBL.Columns(1).Text
    ITEMBL_Matriz(2, grdBL.Row) = grdBL.Columns(2).Text
    ITEMBL_Matriz(3, grdBL.Row) = grdBL.Columns(3).Text
  End If
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserBoletoCredInc.grdBL_BeforeRowColChange]"
End Sub


Private Sub grdBL_UnboundReadDataEx( _
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
               Offset + intI, ITEMBL_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, ITEMBL_COLUNASMATRIZ, ITEMBL_LINHASMATRIZ, ITEMBL_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ITEMBL_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserBoletoCredInc.grdGeral_UnboundReadDataEx]"
End Sub

''''''Private Sub grdBLOrigem_BeforeRowColChange(Cancel As Integer)
''''''  On Error GoTo trata
''''''  If Not ValidaCamposGrupoOrigem(grdBLOrigem.Row, _
''''''                                 grdBLOrigem.Col) Then Cancel = True
''''''  Exit Sub
''''''trata:
''''''  TratarErro Err.Number, Err.Description, "[frmUserBoletoCredInc.grdBLOrigem_BeforeRowColChange]"
''''''End Sub
''''''
''''''Private Sub grdBLOrigem_BeforeUpdate(Cancel As Integer)
''''''  On Error GoTo trata
''''''  'Atualiza Matriz
''''''  ITEMBL_Matriz(7, grdBLOrigem.Row) = grdBLOrigem.Columns(7).Text
''''''  Exit Sub
''''''trata:
''''''  TratarErro Err.Number, Err.Description, "[frmUserBoletoCredInc.grdBLOrigem_BeforeRowColChange]"
''''''End Sub


Private Sub grdBLOrigem_UnboundReadDataEx( _
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
               Offset + intI, ITEMBLLANC_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, ITEMBLLANC_COLUNASMATRIZ, ITEMBLLANC_LINHASMATRIZ, ITEMBLLANC_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ITEMBLLANC_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserBoletoCredInc.grdGeral_UnboundReadDataEx]"
End Sub

Private Sub grdBL_Validate(Cancel As Boolean)
  'Fazer validações ao retirar do grid ou clicar em outro controle
  On Error GoTo trata
  'Cancel = ValidaCamposItemBLDestinoAll
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserBoletoCredInc.grdBL_Validate]"
End Sub



Private Sub txtBoleto_GotFocus()
  Seleciona_Conteudo_Controle txtBoleto
End Sub
Private Sub txtBoleto_LostFocus()
  On Error GoTo trata
  Pintar_Controle txtBoleto, tpCorContr_Normal
  If Me.ActiveControl.Name <> "grdBL" Then Exit Sub
  If Not ValidaCampos Then
    Exit Sub
  End If
  'MsgBox "ok"
  'Montar RecordSet
  ITEMBL_COLUNASMATRIZ = grdBL.Columns.Count
  ITEMBL_LINHASMATRIZ = 0
  ITEMBL_MontaMatriz
  grdBL.Bookmark = Null
  grdBL.ReBind
  grdBL.ApproxCount = ITEMBL_LINHASMATRIZ
  'Montar RecordSet
  ITEMBLLANC_COLUNASMATRIZ = grdBLOrigem.Columns.Count
  ITEMBLLANC_LINHASMATRIZ = 0
  ITEMBLLANC_MontaMatriz
  grdBLOrigem.Bookmark = Null
  grdBLOrigem.ReBind
  grdBLOrigem.ApproxCount = ITEMBLLANC_LINHASMATRIZ
  '
  SetarFoco grdBL
  grdBL.Col = 1
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserBoletoCredInc.txtSenha_LostFocus]"
  
End Sub
Private Sub txtSenha_GotFocus()
  Seleciona_Conteudo_Controle txtUsuario
End Sub

Private Sub txtSenha_LostFocus()
  On Error GoTo trata
  Pintar_Controle txtUsuario, tpCorContr_Normal
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserBoletoCredInc.txtSenha_LostFocus]"
End Sub


Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strSql                As String
  Dim objRs                 As ADODB.Recordset
  Dim objGeral              As busSisMaq.clsGeral
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCampos = False
  lngBOLETOARRECID = 0
  lngFUNCIONARIOID = 0
  If Not Valida_String(txtBoleto, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o número do boleto" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(txtUsuario, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o usuário" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(txtSenha, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a senha" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(strMsg) = 0 Then
    'Ok
    'Valida usuário
    Set objGeral = New busSisMaq.clsGeral
    strSql = "Select FUNCIONARIO.USUARIO, FUNCIONARIO.SENHA, FUNCIONARIO.NIVEL, FUNCIONARIO.PESSOAID, PESSOA.NOME "
    strSql = strSql & " FROM FUNCIONARIO INNER JOIN PESSOA ON PESSOA.PKID = FUNCIONARIO.PESSOAID "
    strSql = strSql & " INNER JOIN ARRECADADOR ON PESSOA.PKID = ARRECADADOR.PESSOAID "
    strSql = strSql & " WHERE FUNCIONARIO.SENHA =  " & Formata_Dados(Encripta(UCase$(txtSenha.Text)), tpDados_Texto)
    strSql = strSql & " AND FUNCIONARIO.USUARIO =  " & Formata_Dados(txtUsuario.Text, tpDados_Texto)
    strSql = strSql & " AND FUNCIONARIO.INDEXCLUIDO =  " & Formata_Dados("N", tpDados_Texto)
  
    Set objRs = objGeral.ExecutarSQL(strSql)
    'Verifica se o usuário existe
    If objRs.EOF Then
      strMsg = strMsg & "Senha/usuário não encontrado"
      Pintar_Controle txtSenha, tpCorContr_Erro
      SetarFoco txtSenha
    Else
      lngFUNCIONARIOID = objRs.Fields("PESSOAID").Value & ""
    End If
    '
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
  End If
  If Len(strMsg) = 0 Then
    'Ok
    'Valida Boleto do usuário
    Set objGeral = New busSisMaq.clsGeral
    strSql = "Select BOLETOARREC.PKID "
    strSql = strSql & " FROM CAIXAARREC INNER JOIN BOLETOARREC ON CAIXAARREC.PKID = BOLETOARREC.CAIXAARRECID "
    strSql = strSql & " WHERE BOLETOARREC.NUMERO =  " & Formata_Dados(txtBoleto.Text, tpDados_Texto)
    strSql = strSql & " AND CAIXAARREC.ARRECADADORID =  " & Formata_Dados(lngFUNCIONARIOID, tpDados_Longo)
  
    Set objRs = objGeral.ExecutarSQL(strSql)
    'Verifica se o boleto existe para o usuário
    If objRs.EOF Then
      strMsg = strMsg & "Boleto não encontrado para este arrecadador"
      Pintar_Controle txtBoleto, tpCorContr_Erro
      SetarFoco txtBoleto
    Else
      lngBOLETOARRECID = objRs.Fields("PKID").Value & ""
    End If
    '
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserBoletoCredInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserBoletoCredInc.ValidaCampos]", _
            Err.Description
End Function


Private Sub txtUsuario_GotFocus()
  Seleciona_Conteudo_Controle txtUsuario
End Sub
Private Sub txtUsuario_LostFocus()
  Pintar_Controle txtUsuario, tpCorContr_Normal
End Sub

