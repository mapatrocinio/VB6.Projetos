VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserLeituraInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leitura"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00C0C0C0&
      Height          =   1245
      Left            =   11760
      ScaleHeight     =   1185
      ScaleWidth      =   1605
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7170
      Width           =   1665
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   880
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   150
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   8295
      Left            =   150
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   150
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   14631
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da leitura"
      TabPicture(0)   =   "userLeituraInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Dados da leitura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   120
         TabIndex        =   10
         Top             =   390
         Width           =   10875
         Begin MSMask.MaskEdBox mskData 
            Height          =   255
            Left            =   750
            TabIndex        =   0
            Top             =   300
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label5 
            Caption         =   "Data"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   11
            Top             =   315
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Leitura a serem lançadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7005
         Left            =   120
         TabIndex        =   9
         Top             =   1230
         Width           =   11385
         Begin TrueDBGrid60.TDBGrid grdLeituraOrigem 
            Height          =   3195
            Left            =   90
            OleObjectBlob   =   "userLeituraInc.frx":001C
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   3750
            Width           =   7725
         End
         Begin VB.CommandButton cmdCadastraItem 
            Caption         =   ">"
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   10950
            TabIndex        =   2
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdCadastraItem 
            Caption         =   ">>"
            Height          =   375
            Index           =   1
            Left            =   10950
            TabIndex        =   3
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton cmdCadastraItem 
            Caption         =   "<"
            Height          =   375
            Index           =   2
            Left            =   10950
            TabIndex        =   4
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton cmdCadastraItem 
            Caption         =   "<<"
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   10950
            TabIndex        =   5
            Top             =   1320
            Width           =   375
         End
         Begin TrueDBGrid60.TDBGrid grdLeitura 
            Height          =   3555
            Left            =   90
            OleObjectBlob   =   "userLeituraInc.frx":7A3E
            TabIndex        =   1
            Top             =   240
            Width           =   10845
         End
      End
   End
End
Attribute VB_Name = "frmUserLeituraInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                 As tpStatus
Public lngLEITURAID         As Long
Public blnRetorno               As Boolean
Public blnFechar                As Boolean
Public blnPrimeiraVez          As Boolean
Dim ITEMLEI_COLUNASMATRIZ         As Long
Dim ITEMLEI_LINHASMATRIZ          As Long
Private ITEMLEI_Matriz()          As String

Dim ITEMLEILANC_COLUNASMATRIZ         As Long
Dim ITEMLEILANC_LINHASMATRIZ          As Long
Private ITEMLEILANC_Matriz()          As String

Private blnSairRow                As Boolean
Private blnSairGrid               As Boolean

Public Sub ITEMLEI_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim objRsInt  As ADODB.Recordset
  Dim objRsConf As ADODB.Recordset
  Dim objRsFabricado As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGeral  As busSisMaq.clsGeral
  '
  On Error GoTo trata

  Set objGeral = New busSisMaq.clsGeral
  '
  strSql = "SELECT MAQUINA.PKID, SERIE.NUMERO, TIPO.TIPO, EQUIPAMENTO.NUMERO, " & _
      "RELMOVIMENTOMAQUINA.LEITURADIAANTESARREC, ISNULL(RELMOVIMENTOMAQUINA.LEITURAATUALANTESARREC,0), (RELMOVIMENTOMAQUINA.LEITURADIAANTESARREC + ISNULL(RELMOVIMENTOMAQUINA.LEITURAATUALANTESARREC,0)), " & _
      "RELMOVIMENTOMAQUINA.LEITURADIAANTESATEND, ISNULL(RELMOVIMENTOMAQUINA.LEITURAATUALANTESATEND,0), (RELMOVIMENTOMAQUINA.LEITURADIAANTESATEND + ISNULL(RELMOVIMENTOMAQUINA.LEITURAATUALANTESATEND,0)), " & _
      " '' AS ENTRADA, '' AS SAIDA, x=1 "
  strSql = strSql & " FROM SERIE " & _
          " INNER JOIN EQUIPAMENTO ON SERIE.PKID = EQUIPAMENTO.SERIEID " & _
          " INNER JOIN MAQUINA ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
          "       AND MAQUINA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
          " INNER JOIN RELMOVIMENTOMAQUINA ON RELMOVIMENTOMAQUINA.MAQUINAID =  MAQUINA.PKID " & _
          "       AND RELMOVIMENTOMAQUINA.PESSOAID = " & Formata_Dados(giFuncionarioId, tpDados_Longo) & _
          " INNER JOIN TIPO ON TIPO.PKID = MAQUINA.TIPOID " & _
          " WHERE MAQUINA.PKID NOT IN (SELECT LEITURAMAQUINA.MAQUINAID FROM LEITURAMAQUINA " & _
          "         WHERE LEITURAMAQUINA.LEITURAID  = " & Formata_Dados(lngLEITURAID, tpDados_Longo) & _
          "         ) " & _
          " ORDER BY EQUIPAMENTO.NUMERO, SERIE.NUMERO, TIPO.TIPO"
  '
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If Not objRs.EOF Then
    ITEMLEI_LINHASMATRIZ = objRs.RecordCount
  Else
    ITEMLEI_LINHASMATRIZ = 0
  End If
  If Not objRs.EOF Then
    ReDim ITEMLEI_Matriz(0 To ITEMLEI_COLUNASMATRIZ - 1, 0 To ITEMLEI_LINHASMATRIZ - 1)
  Else
    ReDim ITEMLEI_Matriz(0 To ITEMLEI_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To ITEMLEI_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To ITEMLEI_COLUNASMATRIZ - 1  'varre as colunas
          If intJ = ITEMLEI_COLUNASMATRIZ - 1 Then
            ITEMLEI_Matriz(intJ, intI) = intI & ""
          Else
            ITEMLEI_Matriz(intJ, intI) = objRs(intJ) & ""
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

Public Sub ITEMLEILANC_MontaMatriz()
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim intI          As Integer
  Dim intJ          As Integer
  Dim intRows       As Integer
  Dim clsGer        As busSisMaq.clsGeral
  '
  On Error GoTo trata


  Set clsGer = New busSisMaq.clsGeral
  '
  strSql = "SELECT LEITURAMAQUINA.PKID, MAQUINA.PKID, SERIE.NUMERO, TIPO.TIPO, EQUIPAMENTO.NUMERO, LEITURAMAQUINA.MEDICAOENTRADA, LEITURAMAQUINA.MEDICAOSAIDA "
  strSql = strSql & " FROM SERIE " & _
          " INNER JOIN EQUIPAMENTO ON SERIE.PKID = EQUIPAMENTO.SERIEID " & _
          " INNER JOIN MAQUINA ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
          " INNER JOIN TIPO ON TIPO.PKID = MAQUINA.TIPOID " & _
          " INNER JOIN LEITURAMAQUINA ON MAQUINA.PKID = LEITURAMAQUINA.MAQUINAID " & _
          " WHERE LEITURAMAQUINA.LEITURAID  = " & Formata_Dados(lngLEITURAID, tpDados_Longo) & _
          " ORDER BY EQUIPAMENTO.NUMERO, SERIE.NUMERO, TIPO.TIPO"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  '
  If Not objRs.EOF Then
    objRs.MoveFirst
    ITEMLEILANC_LINHASMATRIZ = objRs.RecordCount
  Else
    ITEMLEILANC_LINHASMATRIZ = 0
  End If
  If Not objRs.EOF Then
    ReDim ITEMLEILANC_Matriz(0 To ITEMLEILANC_COLUNASMATRIZ - 1, 0 To ITEMLEILANC_LINHASMATRIZ - 1)
  Else
    ReDim ITEMLEILANC_Matriz(0 To ITEMLEILANC_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To ITEMLEILANC_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To ITEMLEILANC_COLUNASMATRIZ - 1  'varre as colunas
          ITEMLEILANC_Matriz(intJ, intI) = objRs(intJ) & ""
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


'''Private Sub cboLeitura_Click()
'''  On Error GoTo trata
'''  Dim objLeitura As busSisMaq.clsLeitura
'''  Dim objRs As ADODB.Recordset
'''  'Alterna para status de alteração/inclusão
'''  If cboLeitura.Text = "" Then
'''    Status = tpStatus_Incluir
'''    lngLEITURAID = 0
'''    Form_Load
'''    'Montar RecordSet
'''    ITEMLEILANC_COLUNASMATRIZ = grdLeituraOrigem.Columns.Count
'''    ITEMLEILANC_LINHASMATRIZ = 0
'''    ITEMLEILANC_MontaMatriz
'''    grdLeituraOrigem.Bookmark = Null
'''    grdLeituraOrigem.ReBind
'''    grdLeituraOrigem.ApproxCount = ITEMLEILANC_LINHASMATRIZ
'''    '
'''    SetarFoco txtNFCliente
'''    Exit Sub
'''  End If
'''  Set objLeitura = New busSisMaq.clsLeitura
'''  Set objRs = objLeitura.ListarLeituraPeloSeq(lngCONTRATOID, _
'''                                                  lngOBRAID, _
'''                                                  Left(cboLeitura.Text, 3))
'''  If objRs.EOF Then
'''    TratarErroPrevisto "Devolução " & cboLeitura.Text & " não cadastrada!"
'''    Status = tpStatus_Incluir
'''    lngLEITURAID = 0
'''    Form_Load
'''  Else
'''    Status = tpStatus_Alterar
'''    lngLEITURAID = objRs.Fields("PKID").Value
'''    Form_Load
'''  End If
'''  'Montar RecordSet
'''  ITEMLEI_COLUNASMATRIZ = grdLeitura.Columns.Count
'''  ITEMLEI_LINHASMATRIZ = 0
'''  ITEMLEI_MontaMatriz
'''  grdLeitura.Bookmark = Null
'''  grdLeitura.ReBind
'''  grdLeitura.ApproxCount = ITEMLEI_LINHASMATRIZ
'''  'Montar RecordSet
'''  ITEMLEILANC_COLUNASMATRIZ = grdLeituraOrigem.Columns.Count
'''  ITEMLEILANC_LINHASMATRIZ = 0
'''  ITEMLEILANC_MontaMatriz
'''  grdLeituraOrigem.Bookmark = Null
'''  grdLeituraOrigem.ReBind
'''  grdLeituraOrigem.ApproxCount = ITEMLEILANC_LINHASMATRIZ
'''  '
'''  SetarFoco txtNFCliente
'''  objRs.Close
'''  Set objRs = Nothing
'''  Set objLeitura = Nothing
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''  AmpN
'''End Sub
'''
Private Sub cmdCadastraItem_Click(Index As Integer)
  On Error GoTo trata
  TratarAssociacao Index + 1
  SetarFoco grdLeitura
  grdLeitura.Col = 1
  If grdLeitura.Row > -1 Then
    grdLeitura.Row = 0
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub TratarAssociacao(pIndice As Integer)
  On Error GoTo trata
  Dim objLeitura       As busSisMaq.clsLeitura
  Dim objLeituraMaq    As busSisMaq.clsLeituraMaquina
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
    strData = mskData.Text
    If lngLEITURAID = 0 Then
      'Leitura não cadastrada para o dia
      Set objLeitura = New busSisMaq.clsLeitura
      objLeitura.InserirLeitura lngLEITURAID, _
                                giFuncionarioId, _
                                strData
      Set objLeitura = Nothing
    End If
    '
    Set objLeituraMaq = New busSisMaq.clsLeituraMaquina
    For intI = 0 To ITEMLEI_LINHASMATRIZ - 1
      grdLeitura.Bookmark = CLng(intI)
      If grdLeitura.Columns("Entrada").Text & "" <> "" And _
          grdLeitura.Columns("Saída").Text & "" <> "" Then
        'Propósito: Retornar todos os ítens
        '
        lngMAQUINAID = grdLeitura.Columns("MAQUINAID").Text & ""
        '
        objLeituraMaq.InserirLeituraMaquina lngLEITURAID, _
                                            lngMAQUINAID, _
                                            grdLeitura.Columns("Entrada").Text & "", _
                                            grdLeitura.Columns("Saída").Text & "", _
                                            strData
                                
        blnRet = True
        'Verifica consolidação
        'VerificaStatusConsolicacao lngLEITURAID
        'Indica se quantidade restante fechou
      End If
    Next
    Set objLeitura = Nothing
    Set objGeral = Nothing
    '
    blnFechar = True
    Unload Me
  Case 3 'Retirar Selecionados
    'Devolução
    'Pede liberação do gerente
    'frmUserLoginLibera.lngFUNCIONARIOID = 0
    'frmUserLoginLibera.strNivel = "'GER','ADM'"
    'frmUserLoginLibera.Show vbModal
    'If Len(Trim(gsNomeUsuLib)) = 0 Then
    '  TratarErroPrevisto "É necessário confirmação do gerente para executar esta ação.", "cmdConfirmar_Click"
    '  Exit Sub
    'End If
    '
    Set objLeituraMaq = New busSisMaq.clsLeituraMaquina
    blnSel = False
    For intI = 0 To grdLeituraOrigem.SelBookmarks.Count - 1
      grdLeituraOrigem.Bookmark = CLng(grdLeituraOrigem.SelBookmarks.Item(intI))
      'excluir debito
      objLeituraMaq.ExcluirLeituraMaquina grdLeituraOrigem.Columns("LEITURAMAQUINAID").Text
      'Verifica consolidação
      'VerificaStatusConsolicacaoArrec lngLEITURAID

      blnSel = True
      blnRet = True
    Next
    Set objLeitura = Nothing
    If blnSel = False Then
      TratarErroPrevisto "Nenhum leitura selecionada para exclusão.", "[frmUserLeituraInc.TratarAssociacao]"
    End If
'''  Case 4 'retirar Todos
'''    'Devolução
'''    Set objLeitura = New busSisMaq.clsLeitura
'''    For intI = 0 To ITEMLEILANC_LINHASMATRIZ - 1
'''      grdLeituraOrigem.Bookmark = CLng(intI)
'''      If IsNull(grdLeituraOrigem.Bookmark) Then grdLeituraOrigem.Bookmark = CLng(intI)
'''
'''      'retornar quantidade ao itens no estoque
'''      objLeitura.AlterarEstoquePelaLeitura grdLeituraOrigem.Columns("ESTOQUEID").Text, _
'''                                               grdLeituraOrigem.Columns("Devol.").Text, _
'''                                               "RET"
'''      objLeitura.ExcluirItemDeVolucao grdLeituraOrigem.Columns("ITEMDEVOLUCAOID").Text
'''      'Verifica consolidação
'''      VerificaStatusConsolicacao grdLeituraOrigem.Columns("NFID").Text
'''      blnRet = True
'''    Next
'''    Set objLeitura = Nothing
  End Select
'''  '
'''  Set clsEstInter = Nothing
'''    '
  If blnRet Then 'Houve Auteração, Atualiza grids
    blnRetorno = True
    '
    ITEMLEI_COLUNASMATRIZ = grdLeitura.Columns.Count
    ITEMLEI_LINHASMATRIZ = 0
    ITEMLEI_MontaMatriz
    grdLeitura.Bookmark = Null
    grdLeitura.ReBind
    grdLeitura.ApproxCount = ITEMLEI_LINHASMATRIZ
    '
    'Montar RecordSet
    ITEMLEILANC_COLUNASMATRIZ = grdLeituraOrigem.Columns.Count
    ITEMLEILANC_LINHASMATRIZ = 0
    ITEMLEILANC_MontaMatriz
    grdLeituraOrigem.Bookmark = Null
    grdLeituraOrigem.ReBind
    grdLeituraOrigem.ApproxCount = ITEMLEILANC_LINHASMATRIZ
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
'''  If grdLeitura.Columns("Informado").Text = "" And grdLeitura.Columns("Avaria").Text = "" And grdLeitura.Columns("Recebido").Text = "" Then
'''    'Não lançou item
'''    ValidaCamposItemNFGeral = True
'''    Exit Function
'''  End If
'''  'Validção de quantidade Informada
'''  If Not Valida_Moeda(grdLeitura.Columns("Informado"), TpNaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
'''    strMsg = strMsg & "Quantidade informada inválida" & vbCrLf
'''  End If
'''  'Validção de quantidade avaria
'''  If Not Valida_Moeda(grdLeitura.Columns("Avaria"), TpNaoObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
'''    strMsg = strMsg & "Quantidade avaria inválida" & vbCrLf
'''  End If
'''  'Validção de quantidade avaria
'''  If Not Valida_Moeda(grdLeitura.Columns("Recebido"), TpObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
'''    strMsg = strMsg & "Quantidade recebido inválida" & vbCrLf
'''  End If
'''  If strMsg = "" Then
'''    'Avaria e Recebido não Recebido
'''    If grdLeitura.Columns("Avaria").Text = "" And grdLeitura.Columns("Recebido").Text = "" Then
'''      strMsg = strMsg & "Informar a quantidade de avaria ou recebido na NFSF." & vbCrLf
'''      SetarFoco grdLeitura
'''    End If
'''  End If
'''  If strMsg = "" Then
'''    'Quantidade informada > quantidade restante
'''    If (CLng(IIf(grdLeitura.Columns("Recebido").Text & "" = "", "0", grdLeitura.Columns("Recebido").Text))) > CLng(grdLeitura.Columns("Restante").Text) Then
'''      strMsg = strMsg & "Quantidade informada não pode ser maior que a quantidade restante da peça na NFSF." & vbCrLf
'''      SetarFoco grdLeitura
'''    End If
'''  End If
'''  If strMsg = "" Then
'''    'Quantidade informada > quantidade restante
'''    If (CLng(IIf(grdLeitura.Columns("Avaria").Text & "" = "", "0", grdLeitura.Columns("Avaria").Text))) > CLng(IIf(grdLeitura.Columns("Recebido").Text & "" = "", "0", grdLeitura.Columns("Recebido").Text)) Then
'''      strMsg = strMsg & "Quantidade de avaria não pode ser maior que a quantidade recebida da peça na NFSF." & vbCrLf
'''      SetarFoco grdLeitura
'''    End If
'''  End If
'''  '
'''  If Len(strMsg) <> 0 Then
'''    TratarErroPrevisto strMsg, "[frmUserLeituraInc.ValidaCamposItemNFGeral]"
'''    ValidaCamposItemNFGeral = False
'''  Else
'''    ValidaCamposItemNFGeral = True
'''  End If
'''  Exit Function
'''trata:
'''  TratarErro Err.Number, _
'''             Err.Description, _
'''             "[frmUserLeituraInc.ValidaCamposItemNFGeral]"
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
  '
  blnSetarFocoControle = True
  '
  'Validção da Medição de entrada
  If Not Valida_Moeda(grdLeitura.Columns("Entrada"), TpObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
    strMsg = strMsg & "Medição de Entrada informada inválida na linha " & intLinha + 1 & vbCrLf
  End If
  'Validção da Medição
  If Not Valida_Moeda(grdLeitura.Columns("Saída"), TpObrigatorio, blnSetarFocoControle, blnPintarControle:=False, blnValidarPeloClip:=False) Then
    strMsg = strMsg & "Medição de saída informada inválida na linha " & intLinha + 1 & vbCrLf
  End If
  '
  If Len(strMsg) = 0 Then
    'Não encontrou erro, válida entrada e saída
    If CCur(grdLeitura.Columns("Entrada").Value) < CCur(grdLeitura.Columns("Prev. Ent.").Value) Then
        strMsg = strMsg & vbCrLf & "Medição de Entrada"
        strMsg = strMsg & vbCrLf & "Informado : " & grdLeitura.Columns("Entrada").Value
        strMsg = strMsg & vbCrLf & "Previsto : " & grdLeitura.Columns("Prev. Ent.").Value
        strMsg = strMsg & vbCrLf
    End If
    If CCur(grdLeitura.Columns("Saída").Value) < CCur(grdLeitura.Columns("Prev. Sai.").Value) Then
        strMsg = strMsg & vbCrLf & "Medição de Saída"
        strMsg = strMsg & vbCrLf & "Informado : " & grdLeitura.Columns("Saída").Value
        strMsg = strMsg & vbCrLf & "Previsto : " & grdLeitura.Columns("Prev. Sai.").Value
        strMsg = strMsg & vbCrLf
    End If
    If strMsg <> "" Then
      strMsg = "ATENÇÃO: Medição de Entrada/Saída menor que o previsto para máquina nro. " & grdLeitura.Columns("Número").Value & vbCrLf & strMsg
      If MsgBox(strMsg & vbCrLf & "Deseja continuar e lançar esta medição?", vbYesNo, TITULOSISTEMA) = vbYes Then
        strMsg = ""
      End If
    End If
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserLeituraInc.ValidaCamposItemBLDestino]"
    ValidaCamposItemBLDestino = False
  Else
    ValidaCamposItemBLDestino = True
  End If
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserLeituraInc.ValidaCamposItemBLDestino]"
  ValidaCamposItemBLDestino = False
End Function

Private Function ValidaCamposItemBLDestinoAllSel() As Boolean
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
  
  
  For intRows = 0 To ITEMLEI_LINHASMATRIZ - 1
    grdLeitura.Bookmark = CLng(intRows)
    '
    If grdLeitura.Columns("Entrada").Text & "" <> "" And _
      grdLeitura.Columns("Saída").Text & "" <> "" Then
      If Not ValidaCamposItemBLDestino(grdLeitura.Row) Then
        blnEncontrouErro = True
        blnEncontrouErroLinha = True
      Else
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
    TratarErroPrevisto "Entre com ao menos 1 medição de entrada e saída", "[frmUserLeituraInc.ValidaCamposItemBLDestinoAllSel]"
  End If
  grdLeitura.ReBind
  grdLeitura.SetFocus
  ValidaCamposItemBLDestinoAllSel = blnRetorno
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserLeituraInc.ValidaCamposItemBLDestinoAllSel]"
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
'''  For intRows = 0 To ITEMLEI_LINHASMATRIZ - 1
'''    grdLeitura.Bookmark = CLng(intRows)
'''    blnRetorno = ValidaCamposItemNFGeral
'''    If Not blnRetorno Then Exit For
'''    If blnRetorno Then
'''      If grdLeitura.Columns("Avaria").Text & "" <> "" Or grdLeitura.Columns("Recebido").Text & "" <> "" Then
'''        blnLancouItem = True
'''      End If
'''    End If
'''  Next
'''  '
'''  If blnLancouItem = False Then
'''    blnRetorno = False
'''    TratarErroPrevisto "Nenhum item lançado para esta NF.", "[frmUserLeituraInc.ValidaCamposItemBLDestinoAll]"
'''  End If
'''  If blnRetorno = True Then
'''    'Nenhum erro encontrado
'''    If Not Valida_String(txtNFCliente, TpObrigatorio, False) Then
'''      TratarErroPrevisto "Número NF Cliente inválido.", "[frmUserLeituraInc.ValidaCamposItemBLDestinoAll]"
'''      blnRetorno = False
'''    End If
'''  End If
'''  If blnRetorno = True Then
'''    If Not Valida_Data(mskData, TpObrigatorio, False) Then
'''      TratarErroPrevisto "Data inválida.", "[frmUserLeituraInc.ValidaCamposItemBLDestinoAll]"
'''      blnRetorno = False
'''    End If
'''  End If
'''  If blnRetorno = True Then
'''    If Not Valida_Data(mskDataEmissao, TpObrigatorio, False) Then
'''      TratarErroPrevisto "Data de emissão inválida.", "[frmUserLeituraInc.ValidaCamposItemBLDestinoAll]"
'''      blnRetorno = False
'''    End If
'''  End If
'''  If blnRetorno = True Then
'''    If Not Valida_Data(mskDataLeitura, TpObrigatorio, False) Then
'''      TratarErroPrevisto "Data de devolução inválida.", "[frmUserLeituraInc.ValidaCamposItemBLDestinoAll]"
'''      blnRetorno = False
'''    End If
'''  End If
'''  If blnRetorno = False Then
'''    grdLeitura.ReBind
'''    grdLeitura.SetFocus
'''  End If
'''  ValidaCamposItemBLDestinoAll = Not blnRetorno
'''  Exit Function
'''trata:
'''  TratarErro Err.Number, _
'''             Err.Description, _
'''             "[frmUserLeituraInc.ValidaCamposItemBLDestinoAll]"
'''  ValidaCamposItemBLDestinoAll = False
'''End Function
'''
Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Leitura
  LimparCampoMask mskData
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserLeituraInc.LimparCampos]", _
            Err.Description
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  '
  If Me.ActiveControl.Name <> "grdLeitura" Then
    If KeyAscii = 13 Then
      SendKeys "{tab}"
    End If
  Else
    If KeyAscii = 13 And grdLeitura.Row <> -1 Then
      If grdLeitura.Col = 11 Then
        blnSairRow = True
        blnSairGrid = True
        '
        'ITEMLEI_Matriz(4, grdLeitura.Row) = grdLeitura.Columns(4).Text
        'ITEMLEI_Matriz(5, grdLeitura.Row) = grdLeitura.Columns(5).Text
        ITEMLEI_Matriz(10, grdLeitura.Columns("ROWNUM").Value) = grdLeitura.Columns(10).Text
        ITEMLEI_Matriz(11, grdLeitura.Columns("ROWNUM").Value) = grdLeitura.Columns(11).Text
        
        '
        'Para cada linha verifica se está em branco, se sim simula o ENTER
        'If (grdLeitura.Columns("Máquina").Text & "" = "" _
        '   And grdLeitura.Columns("Medição").Text & "" = "" _
        '   And grdLeitura.Columns("Valor").Text & "" = "") Or ((grdLeitura.Row + 1) = ITEMLEI_LINHASMATRIZ) Then
        '  cmdCadastraItem_Click 1
        If grdLeitura.Columns("ROWNUM").Value + 1 = ITEMLEI_LINHASMATRIZ Then
          cmdCadastraItem_Click 1
        Else
          grdLeitura.Col = 10
          'grdLeitura.Row = grdLeitura.Row + 1
          grdLeitura.MoveNext
        End If
        blnSairRow = False
        blnSairGrid = False
        '
      Else
        grdLeitura.Col = grdLeitura.Col + 1
      End If
    End If
  End If
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserLeituraInc.Form_Activate]"
End Sub



Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Montar RecordSet
    ITEMLEI_COLUNASMATRIZ = grdLeitura.Columns.Count
    ITEMLEI_LINHASMATRIZ = 0
    ITEMLEI_MontaMatriz
    grdLeitura.Bookmark = Null
    grdLeitura.ReBind
    grdLeitura.ApproxCount = ITEMLEI_LINHASMATRIZ
    'Montar RecordSet
    ITEMLEILANC_COLUNASMATRIZ = grdLeituraOrigem.Columns.Count
    ITEMLEILANC_LINHASMATRIZ = 0
    ITEMLEILANC_MontaMatriz
    grdLeituraOrigem.Bookmark = Null
    grdLeituraOrigem.ReBind
    grdLeituraOrigem.ApproxCount = ITEMLEILANC_LINHASMATRIZ
    '
    SetarFoco mskData
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserLeituraInc.Form_Activate]"
End Sub

Private Sub Form_Load()
On Error GoTo trata
  '
  blnFechar = False
  blnRetorno = False
  AmpS
  Me.Height = 8970
  Me.Width = 13590
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
'''    Set objLeitura = New busSisMaq.clsLeitura
'''    Set objRs = objLeitura.ListarLeitura(lngLEITURAID)
'''    '
'''    If Not objRs.EOF Then
'''      txtSequencial.Text = Format(objRs.Fields("SEQUENCIAL").Value, "000") & ""
'''      INCLUIR_VALOR_NO_MASK mskData, objRs.Fields("DATA").Value & "", TpMaskData
'''      INCLUIR_VALOR_NO_MASK mskDataEmissao, objRs.Fields("DATAEMISSAO").Value & "", TpMaskData
'''      INCLUIR_VALOR_NO_MASK mskDataLeitura, objRs.Fields("DATADEVOLUCAO").Value & "", TpMaskData
'''      txtNFCliente.Text = Format(objRs.Fields("NUMERONF").Value, "000") & ""
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
'''    '
'''    Set objLeitura = Nothing
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
'''Private Sub grdLeitura_BeforeRowColChange(Cancel As Integer)
'''  On Error GoTo trata
'''  'If Not ValidaCamposItemBLDestino(grdLeitura.Row, _
'''                                  grdLeitura.Col) Then Cancel = True



'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, "[frmUserLeituraInc.grdLeitura_BeforeRowColChange]"
'''End Sub
'''
Private Sub grdLeitura_BeforeUpdate(Cancel As Integer)
  On Error GoTo trata
  'Atualiza Matriz
  If blnSairRow = False Then
    ITEMLEI_Matriz(10, grdLeitura.Row) = grdLeitura.Columns(10).Text
    ITEMLEI_Matriz(11, grdLeitura.Row) = grdLeitura.Columns(11).Text
  End If
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserLeituraInc.grdLeitura_BeforeRowColChange]"
End Sub


Private Sub grdLeitura_UnboundReadDataEx( _
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
               Offset + intI, ITEMLEI_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, ITEMLEI_COLUNASMATRIZ, ITEMLEI_LINHASMATRIZ, ITEMLEI_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ITEMLEI_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserLeituraInc.grdGeral_UnboundReadDataEx]"
End Sub

''''''Private Sub grdLeituraOrigem_BeforeRowColChange(Cancel As Integer)
''''''  On Error GoTo trata
''''''  If Not ValidaCamposGrupoOrigem(grdLeituraOrigem.Row, _
''''''                                 grdLeituraOrigem.Col) Then Cancel = True
''''''  Exit Sub
''''''trata:
''''''  TratarErro Err.Number, Err.Description, "[frmUserLeituraInc.grdLeituraOrigem_BeforeRowColChange]"
''''''End Sub
''''''
''''''Private Sub grdLeituraOrigem_BeforeUpdate(Cancel As Integer)
''''''  On Error GoTo trata
''''''  'Atualiza Matriz
''''''  ITEMLEI_Matriz(7, grdLeituraOrigem.Row) = grdLeituraOrigem.Columns(7).Text
''''''  Exit Sub
''''''trata:
''''''  TratarErro Err.Number, Err.Description, "[frmUserLeituraInc.grdLeituraOrigem_BeforeRowColChange]"
''''''End Sub


Private Sub grdLeituraOrigem_UnboundReadDataEx( _
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
               Offset + intI, ITEMLEILANC_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, ITEMLEILANC_COLUNASMATRIZ, ITEMLEILANC_LINHASMATRIZ, ITEMLEILANC_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ITEMLEILANC_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserLeituraInc.grdGeral_UnboundReadDataEx]"
End Sub

Private Sub grdLeitura_Validate(Cancel As Boolean)
  'Fazer validações ao retirar do grid ou clicar em outro controle
  On Error GoTo trata
  'Cancel = ValidaCamposItemBLDestinoAll
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserLeituraInc.grdLeitura_Validate]"
End Sub


Private Sub mskData_GotFocus()
  Seleciona_Conteudo_Controle mskData
End Sub
Private Sub mskData_LostFocus()
  On Error GoTo trata
  Dim objGeral As busSisMaq.clsGeral
  Dim lngSeq  As Long
  '
  Pintar_Controle mskData, tpCorContr_Normal
  If Me.ActiveControl.Name <> "grdLeitura" Then Exit Sub
  If Not ValidaCampos Then
    Exit Sub
  End If
  'MsgBox "ok"
  'Tratar inserção em tabela temporária
  Set objGeral = New busSisMaq.clsGeral
  '
  lngSeq = objGeral.ExecutarSQLRetInteger("SP_REL_MOVIMENTO_MAQUINA", Array( _
                                            mp("@PESSOAID", adInteger, 4, giFuncionarioId), _
                                            mp("@PERIODOID", adInteger, 4, 0), _
                                            mp("@DATAINICHR", adVarChar, 30, mskData.Text), _
                                            mp("@DATAFIMCHR", adVarChar, 30, mskData.Text)))
  Set objGeral = Nothing
  'Montar RecordSet
  ITEMLEI_COLUNASMATRIZ = grdLeitura.Columns.Count
  ITEMLEI_LINHASMATRIZ = 0
  ITEMLEI_MontaMatriz
  grdLeitura.Bookmark = Null
  grdLeitura.ReBind
  grdLeitura.ApproxCount = ITEMLEI_LINHASMATRIZ
  'Montar RecordSet
  ITEMLEILANC_COLUNASMATRIZ = grdLeituraOrigem.Columns.Count
  ITEMLEILANC_LINHASMATRIZ = 0
  ITEMLEILANC_MontaMatriz
  grdLeituraOrigem.Bookmark = Null
  grdLeituraOrigem.ReBind
  grdLeituraOrigem.ApproxCount = ITEMLEILANC_LINHASMATRIZ
  '
  SetarFoco grdLeitura
  grdLeitura.Col = 10
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserLeituraInc.txtSenha_LostFocus]"
  
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
  lngLEITURAID = 0
  If Not Valida_Data(mskData, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a da válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(strMsg) = 0 Then
    'Ok
    Set objGeral = New busSisMaq.clsGeral
    strSql = "Select LEITURA.PKID "
    strSql = strSql & " FROM LEITURA "
    strSql = strSql & " WHERE LEITURA.DATA =  " & Formata_Dados(mskData.Text, tpDados_DataHora)
    strSql = strSql & " AND LEITURA.LEITURISTAID =  " & Formata_Dados(giFuncionarioId, tpDados_Longo)
  
    Set objRs = objGeral.ExecutarSQL(strSql)
    'Verifica se o boleto existe para o usuário
    If objRs.EOF Then
      lngLEITURAID = 0
    Else
      lngLEITURAID = objRs.Fields("PKID").Value & ""
    End If
    '
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserLeituraInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserLeituraInc.ValidaCampos]", _
            Err.Description
End Function
