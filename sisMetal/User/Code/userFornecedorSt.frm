VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmFornecedorSt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encaminhar pedido ao fornecedor"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   6525
      Left            =   10050
      ScaleHeight     =   6525
      ScaleWidth      =   1860
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2925
         Left            =   120
         ScaleHeight     =   2865
         ScaleWidth      =   1605
         TabIndex        =   4
         Top             =   3510
         Width           =   1665
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   990
            Width           =   1335
         End
      End
      Begin Crystal.CrystalReport Report1 
         Left            =   630
         Top             =   810
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   6315
      Left            =   90
      TabIndex        =   5
      Top             =   120
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   11139
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados gerais"
      TabPicture(0)   =   "userFornecedorSt.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
End
Attribute VB_Name = "frmFornecedorSt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''Option Explicit
'''
'''Public Status                         As tpStatus
'''Public lngPEDIDOID                    As Long
'''Public strAnoOS                       As String
'''Public blnRetorno                     As Boolean
'''Public blnPrimeiraVez                 As Boolean
'''Public blnFechar                      As Boolean
'''
'''Private Sub cmdCancelar_Click()
'''  blnFechar = True
'''  '
'''  Unload Me
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''  Unload Me
'''End Sub
'''
'''Private Sub cmdImprimir_Click()
'''  On Error GoTo TratErro
'''  AmpS
'''  '
'''  Report1.Connect = ConnectRpt
'''  Report1.ReportFileName = gsReportPath & "Pedido.rpt"
'''  '
'''  'If optSai1.Value Then
'''    Report1.Destination = 0 'Video
'''  'ElseIf optSai2.Value Then
'''  '  Report1.Destination = 1   'Impressora
'''  'End If
'''  Report1.CopiesToPrinter = 1
'''  Report1.WindowState = crptMaximized
'''  '
'''  Report1.Formulas(0) = "PEDIDOID = " & Formata_Dados(lngPEDIDOID, tpDados_Longo)
'''  '
'''  Report1.Action = 1
'''  '
'''  AmpN
'''  Exit Sub
'''
'''TratErro:
'''  TratarErro Err.Number, _
'''             Err.Description, _
'''             Err.Source
'''  AmpN
'''End Sub
'''
'''Private Sub cmdOK_Click()
'''  On Error GoTo trata
'''  Dim objPedido               As busSisMetal.clsPedido
'''  '
'''  Select Case tabDetalhes.Tab
'''  Case 0 'Alteração do status para encaminhado para fornecedor
'''    'Valida se cor já cadastrada
'''    If MsgBox("Confirma envio do pedido " & strAnoOS & " para o fornecedor " & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco txtItemPedido
'''      Exit Sub
'''    End If
'''    '
'''    Set objPedido = New busSisMetal.clsPedido
'''    objPedido.AlterarStatusFornecedor lngPEDIDOID
'''    Set objPedido = Nothing
'''    blnFechar = True
'''    Unload Me
'''  End Select
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''End Sub
'''
'''
'''Private Sub Form_Activate()
'''  On Error GoTo trata
'''  If blnPrimeiraVez Then
'''    'Setar foco
'''    SetarFoco txtItemPedido
'''    blnPrimeiraVez = False
'''  End If
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, "[frmFornecedorSt.Form_Activate]"
'''End Sub
'''
'''Private Sub Form_Load()
'''  On Error GoTo trata
'''  Dim objRs           As ADODB.Recordset
'''  Dim strSql          As String
'''  Dim objGeral        As busSisMetal.clsGeral
'''  Dim strItem         As String
'''  '
'''  Dim curQtdAnodTot As Currency
'''  Dim curQtdEmpTot  As Currency
'''  Dim curPesoAnodTot As Currency
'''  Dim curPesoEmpTot  As Currency
'''  Dim curValorAnodTot As Currency
'''  Dim curValorEmpTot  As Currency
'''  '
'''  blnFechar = False
'''  blnRetorno = False
'''  AmpS
'''  Me.Height = 7005
'''  Me.Width = 12000
'''  CenterForm Me
'''  blnPrimeiraVez = True
'''  '
'''  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, , , , , , cmdImprimir
'''  '
'''  If Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
'''    Set objGeral = New busSisMetal.clsGeral
'''    '
'''    strSql = "SELECT ITEM_PEDIDO.PKID, LF.CODIGO_LINHA_FORNECEDOR, " & _
'''            "TIPO_LINHA.NOME + ' - ' + LINHA.CODIGO AS LINHA_CODIGO, ITEM_PEDIDO.QTD_ANODIZADORA, ITEM_PEDIDO.PESO_ANODIZADORA, ISNULL(VALOR_ALUMINIO, 0) * ISNULL(ITEM_PEDIDO.PESO_ANODIZADORA, 0) AS VALOR_ANODIZADORA, " & _
'''            "ITEM_PEDIDO.QTD_EMPRESA, ITEM_PEDIDO.PESO_EMPRESA, ISNULL(VALOR_ALUMINIO, 0) * ISNULL(ITEM_PEDIDO.PESO_EMPRESA, 0) AS VALOR_EMPRESA, ITEM_PEDIDO.COMPRIMENTO_VARA " & _
'''            "FROM ITEM_PEDIDO " & _
'''            " INNER JOIN PEDIDO ON PEDIDO.PKID = ITEM_PEDIDO.PEDIDOID " & _
'''            " LEFT JOIN LINHA ON LINHA.PKID = ITEM_PEDIDO.LINHAID " & _
'''            " LEFT JOIN TIPO_LINHA ON TIPO_LINHA.PKID = LINHA.TIPO_LINHAID " & _
'''            " LEFT JOIN LOJA AS FORNECEDOR ON FORNECEDOR.PKID = PEDIDO.FORNECEDORID " & _
'''            " LEFT JOIN " & _
'''            "   (SELECT LINHA_FORNECEDOR.CODIGO AS CODIGO_LINHA_FORNECEDOR, LINHA_FORNECEDOR.LINHAID, LINHA_FORNECEDOR.FORNECEDORID " & _
'''            "     FROM LINHA_FORNECEDOR) LF ON LF.LINHAID = LINHA.PKID AND LF.FORNECEDORID = PEDIDO.FORNECEDORID " & _
'''            " WHERE ITEM_PEDIDO.PEDIDOID = " & Formata_Dados(lngPEDIDOID, tpDados_Longo) & _
'''            " ORDER BY ITEM_PEDIDO.PKID"
'''    '
'''    Set objRs = objGeral.ExecutarSQL(strSql)
'''    strItem = ""
'''    If objRs.EOF Then
'''      strItem = "Não há itens cadastrados para este pedido"
'''    Else
'''      '
'''      curQtdAnodTot = 0
'''      curQtdEmpTot = 0
'''      curPesoAnodTot = 0
'''      curPesoEmpTot = 0
'''      curValorAnodTot = 0
'''      curValorEmpTot = 0
'''      '
'''      'cabeçalho
'''      strItem = "Linha-perfil" & vbTab & "Cod. Forn." & vbTab & "Qtd. Anod." & vbTab & "Peso Anod." & vbTab & "Vr. Anod." & vbTab & "Qtd. Emp." & vbTab & "Peso Emp." & vbTab & "Vr. Emp." & vbCrLf
'''      Do While Not objRs.EOF
'''        'Para cada linha montar item
'''        strItem = strItem & _
'''            objRs.Fields("LINHA_CODIGO").Value & vbTab & _
'''            objRs.Fields("CODIGO_LINHA_FORNECEDOR").Value & vbTab & _
'''            Format(IIf(IsNull(objRs.Fields("QTD_ANODIZADORA").Value), 0, objRs.Fields("QTD_ANODIZADORA").Value), "###,##0") & vbTab & _
'''            Format(IIf(IsNull(objRs.Fields("PESO_ANODIZADORA").Value), 0, objRs.Fields("PESO_ANODIZADORA").Value), "###,##0.000") & vbTab & _
'''            Format(IIf(IsNull(objRs.Fields("VALOR_ANODIZADORA").Value), 0, objRs.Fields("VALOR_ANODIZADORA").Value), "###,##0.00") & vbTab & _
'''            Format(IIf(IsNull(objRs.Fields("QTD_EMPRESA").Value), 0, objRs.Fields("QTD_EMPRESA").Value), "###,##0") & vbTab & _
'''            Format(IIf(IsNull(objRs.Fields("PESO_EMPRESA").Value), 0, objRs.Fields("PESO_EMPRESA").Value), "###,##0.000") & vbTab & _
'''            Format(IIf(IsNull(objRs.Fields("VALOR_EMPRESA").Value), 0, objRs.Fields("VALOR_EMPRESA").Value), "###,##0.00") & vbCrLf
'''        curQtdAnodTot = curQtdAnodTot + IIf(IsNull(objRs.Fields("QTD_ANODIZADORA").Value), 0, objRs.Fields("QTD_ANODIZADORA").Value)
'''        curPesoAnodTot = curPesoAnodTot + IIf(IsNull(objRs.Fields("PESO_ANODIZADORA").Value), 0, objRs.Fields("PESO_ANODIZADORA").Value)
'''        curValorAnodTot = curValorAnodTot + IIf(IsNull(objRs.Fields("VALOR_ANODIZADORA").Value), 0, objRs.Fields("VALOR_ANODIZADORA").Value)
'''        curQtdEmpTot = curQtdEmpTot + IIf(IsNull(objRs.Fields("QTD_EMPRESA").Value), 0, objRs.Fields("QTD_EMPRESA").Value)
'''        curPesoEmpTot = curPesoEmpTot + IIf(IsNull(objRs.Fields("PESO_EMPRESA").Value), 0, objRs.Fields("PESO_EMPRESA").Value)
'''        curValorEmpTot = curValorEmpTot + IIf(IsNull(objRs.Fields("VALOR_EMPRESA").Value), 0, objRs.Fields("VALOR_EMPRESA").Value)
'''        '
'''        objRs.MoveNext
'''      Loop
'''
'''    End If
'''    strItem = strItem & _
'''        "TOTAL" & vbTab & _
'''        Format(IIf(IsNull(curQtdAnodTot), 0, curQtdAnodTot), "###,##0") & vbTab & _
'''        Format(IIf(IsNull(curPesoAnodTot), 0, curPesoAnodTot), "###,##0.000") & vbTab & _
'''        Format(IIf(IsNull(curValorAnodTot), 0, curValorAnodTot), "###,##0.00") & vbTab & _
'''        Format(IIf(IsNull(curQtdEmpTot), 0, curQtdEmpTot), "###,##0") & vbTab & _
'''        Format(IIf(IsNull(curPesoEmpTot), 0, curPesoEmpTot), "###,##0.000") & vbTab & _
'''        Format(IIf(IsNull(curValorEmpTot), 0, curValorEmpTot), "###,##0.00") & vbCrLf
'''    '
'''    txtItemPedido.Text = strItem
''''''    'Pega Dados do Banco de dados
''''''    Set objVara = New busSisMetal.clsVara
''''''    Set objRs = objVara.SelecionarVara(lngPEDIDOID)
''''''    '
''''''    If Not objRs.EOF Then
''''''      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VALOR").Value, TpMaskMoeda
''''''      txtItemPedido.Text = objRs.Fields("NOME").Value & ""
''''''    End If
''''''    Set objVara = Nothing
'''  End If
'''  '
'''  AmpN
'''  Exit Sub
'''trata:
'''  AmpN
'''  TratarErro Err.Number, Err.Description, Err.Source
'''  Unload Me
'''End Sub
'''
'''Private Sub Form_Unload(Cancel As Integer)
'''  If Not blnFechar Then Cancel = True
'''End Sub
'''
'''Private Sub txtItemPedido_GotFocus()
'''  Selecionar_Conteudo txtItemPedido
'''End Sub
'''
'''Private Sub txtItemPedido_LostFocus()
'''  Pintar_Controle txtItemPedido, tpCorContr_Normal
'''End Sub
'''
