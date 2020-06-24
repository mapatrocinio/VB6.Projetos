VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserDetalhamentoArrec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalhamento"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   8430
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5715
      Left            =   6570
      ScaleHeight     =   5715
      ScaleWidth      =   1860
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   90
         ScaleHeight     =   2025
         ScaleWidth      =   1575
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1635
         Begin VB.CommandButton cmdFiltrar 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   150
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1020
            Width           =   1335
         End
      End
      Begin MSComctlLib.ImageList imlSmallIcons 
         Left            =   120
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   13
         ImageHeight     =   13
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userDetalhamentoArrec.frx":0000
               Key             =   "closed"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userDetalhamentoArrec.frx":059C
               Key             =   "closedgreen"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userDetalhamentoArrec.frx":0B38
               Key             =   "open"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userDetalhamentoArrec.frx":10D4
               Key             =   "opengreen"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userDetalhamentoArrec.frx":1670
               Key             =   "detalhenaook"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userDetalhamentoArrec.frx":198C
               Key             =   "detalheok"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userDetalhamentoArrec.frx":1CA8
               Key             =   "nocheck"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userDetalhamentoArrec.frx":2584
               Key             =   "vertexto"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userDetalhamentoArrec.frx":28A0
               Key             =   "check"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userDetalhamentoArrec.frx":317C
               Key             =   "leaf"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userDetalhamentoArrec.frx":32EE
               Key             =   "open_"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userDetalhamentoArrec.frx":3460
               Key             =   "smlBook"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.TreeView tvwDB 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   9975
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlSmallIcons"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmUserDetalhamentoArrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mNode                 As MSComctlLib.node ' Module-level Node variable.
Private mItem                 As MSComctlLib.ListItem ' Module-level ListItem variable.
Private lngCurrentIndex       As Long ' Flag to assure this node wasn't already clicked.
Private lngTURNOCORRENTEID    As Long
Public strSqlWhereTurno       As String
Public strSqlWhereLocacao     As String


Public Function RetornaChaveNo(strChave As String) As Long
  On Error GoTo trata:
  Dim strRetorno As String
  strRetorno = Replace(strChave, "TURNOID_", "")
  strRetorno = Replace(strRetorno, "CAIXAARRECID_", "")
  strRetorno = Replace(strRetorno, "A_", "")
  RetornaChaveNo = CLng(strRetorno)
  Exit Function
trata:
  Err.Raise Err.Number, "[frmUserDetalhamentoArrec.RetornaChaveNo]", Err.Description
End Function

Private Sub CarregarTurnoArrec(pTurno As String, _
                               lngTURNOID As Long, _
                               Optional ByRef ParentNode As MSComctlLib.node)
  On Error GoTo trata:
  ' While the record is not the last record, add a ListItem object.
  ' Use the Name field for the ListItem object's text.
  Dim objGeral As busSisMaq.clsGeral
  Dim objRS As ADODB.Recordset
  Dim strSql As String
  '
  ' Check that the node isn't already populated. If it is, then
  ' add only the ListItem objects to the ListView and exit.
  On Error GoTo Sai
  If ParentNode.Children Then
    'AddListItemsOnly pTurno
    Exit Sub
  End If
Sai:
  On Error GoTo trata:
  Set objGeral = New busSisMaq.clsGeral
  '
  strSql = "SELECT CAIXAARREC.PKID, PESSOA.NOME, ISNULL(CAIXAARREC.TURNOFECHAID,0) AS TURNOFECHAID FROM CAIXAARREC " & _
      " INNER JOIN ARRECADADOR ON ARRECADADOR.PESSOAID = CAIXAARREC.ARRECADADORID " & _
      " INNER JOIN PESSOA ON PESSOA.PKID = ARRECADADOR.PESSOAID " & _
      " WHERE TURNOENTRADAID = " & Formata_Dados(lngTURNOID, tpDados_Longo) & _
      " ORDER BY CAIXAARREC.PKID DESC"
      
  Set objRS = objGeral.ExecutarSQL(strSql)
  Do While Not objRS.EOF
    Set mNode = tvwDB.Nodes.Add(pTurno, tvwChild, "CAIXAARRECID_" & objRS.Fields("PKID").Value, objRS.Fields("NOME").Value, IIf(objRS.Fields("TURNOFECHAID").Value = 0, "closedgreen", "closed"))
    mNode.Tag = "CAIXAARREC" ' Identifies the table.
    objRS.MoveNext
  Loop
  ' Sort the Turnos nodes.
  tvwDB.Nodes(pTurno).Sorted = False
  ' Expand top node.
  tvwDB.Nodes(pTurno).Expanded = True
  lngCurrentIndex = RetornaChaveNo(pTurno)
  '
'  Set objLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, "[frmUserDetalhamentoArrec.CarregarTurnoArrec]", Err.Description
End Sub

Private Sub CarregarUnidades(pTurno As String, _
                             Optional ByRef ParentNode As MSComctlLib.node)
  On Error GoTo trata:
  ' While the record is not the last record, add a ListItem object.
  ' Use the Name field for the ListItem object's text.
  'Dim objLoc As busSisMaq.clsLocacao
  Dim objRS As ADODB.Recordset
  '
  ' Check that the node isn't already populated. If it is, then
  ' add only the ListItem objects to the ListView and exit.
  On Error GoTo Sai
  If ParentNode.Children Then
    'AddListItemsOnly pTurno
    Exit Sub
  End If
Sai:
  On Error GoTo trata:
'''  Set objLoc = New busSisMaq.clsLocacao
'''  '
'''  Set objRs = objLoc.ListarLocacaoPorUnidade(RetornaChaveNo(pTurno), strSqlWhereLocacao)
'''  Do While Not objRs.EOF
'''    Set mNode = tvwDB.Nodes.Add(pTurno, tvwChild, "LOCACAOID_" & objRs.Fields("PKID").Value, objRs.Fields("NUMERO").Value, IIf(objRs.Fields("OCUPADO").Value = True, "closedgreen", "closed"))
'''    mNode.Tag = "LOCACAO" ' Identifies the table.
'''
'''    objRs.MoveNext
'''  Loop
  ' Sort the Turnos nodes.
  tvwDB.Nodes(pTurno).Sorted = False
  ' Expand top node.
  tvwDB.Nodes(pTurno).Expanded = True
  lngCurrentIndex = RetornaChaveNo(pTurno)
  '
'  Set objLoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, "[frmUserDetalhamentoArrec.CarregarUnidades]", Err.Description
End Sub

Private Sub DetalharTurnoArrec(pLocacao As String, _
                             pNode As MSComctlLib.node)
  On Error GoTo trata:
  Dim objForm   As SisMaq.frmUserOperArrCons
  Set objForm = New SisMaq.frmUserOperArrCons
  '
  objForm.lngTURNOARRECEPESQ = RetornaChaveNo(pLocacao)
  objForm.Status = tpStatus_Consultar
  objForm.Show vbModal
  Set objForm = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, "[frmUserDetalhamentoArrec.CarregarUnidades]", Err.Description
End Sub


Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdFiltrar_Click()
  Dim objForm As SisMaq.frmUserFiltroOper
  On Error GoTo trata
  Set objForm = New SisMaq.frmUserFiltroOper
  objForm.Show vbModal
  'If objForm.strSqlWhereTurno <> "" Then
    strSqlWhereTurno = objForm.strSqlWhereTurno
    strSqlWhereLocacao = objForm.strSqlWhereLocacao
    Form_Load
  'End If
  Set objForm = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Load()
  Dim objTurno As busSisMaq.clsTurno
  Dim objRS As ADODB.Recordset
  Dim lngTotal As Long
  On Error GoTo trata
  AmpS
  Me.Height = 6090
  Me.Width = 8520
  CenterForm Me
  lngTURNOCORRENTEID = RetornaCodTurnoCorrente

  Set objTurno = New busSisMaq.clsTurno
  ' Configure TreeView
  tvwDB.Nodes.Clear
  With tvwDB
    .Sorted = False
    .LabelEdit = False
    .LineStyle = tvwRootLines
  End With
  Set objRS = objTurno.ListarTurnoPorUnidade(strSqlWhereTurno)
  lngTotal = 1
  Do While Not objRS.EOF
    Set mNode = tvwDB.Nodes.Add()
    With mNode ' Add node turno.
      .Text = objRS.Fields("DESCTURNO").Value & ""
      .Key = "TURNOID_" & objRS.Fields("PKID").Value
      .Tag = "TURNO"
      .Image = IIf(lngTURNOCORRENTEID = objRS.Fields("PKID").Value, "closedgreen", "closed")
    End With
    If lngTotal = 1 Then
      'CarregarUnidades "TURNOID_" & objRs.Fields("PKID").Value
      CarregarTurnoArrec "TURNOID_" & objRS.Fields("PKID").Value, objRS.Fields("PKID").Value
    End If
    lngTotal = lngTotal + 1
    objRS.MoveNext
  Loop
  objRS.Close
  Set objRS = Nothing
  
  'CarregarUnidades "TURNOID_" & RetornaCodTurnoCorrente
  Set objTurno = Nothing
  LerFiguras Me, tpBmp_Vazio, pbtnFechar:=cmdFechar, pbtnFiltrar:=cmdFiltrar
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub



Private Sub tvwDB_Collapse(ByVal node As MSComctlLib.node)
    ' Only nodes that are folders can be collapsed.
    If node.Tag = "CAIXAARREC" Then node.Image = IIf(node.Image = "opengreen", "closedgreen", "closed")
    If node.Tag = "TURNO" Then node.Image = IIf(RetornaChaveNo(node.Key) = lngTURNOCORRENTEID, "closedgreen", "closed")
End Sub

Private Sub tvwDB_Expand(ByVal node As MSComctlLib.node)
    ' Only the top node, and publisher nodes can be expanded.

    node.Sorted = False
    If node.Tag = "CAIXAARREC" Then node.Image = IIf(node.Image = "closedgreen", "opengreen", "open")
    If node.Tag = "TURNO" Then node.Image = IIf(RetornaChaveNo(node.Key) = lngTURNOCORRENTEID, "opengreen", "open")
    '
    ' If the Tag is "Publisher" and the mItelngCurrentIndex
    ' index isn't the same as the Node.key, then
    ' invoke the CarregaInfTurnoArrec function.
    'If node.Tag = "CAIXAARREC" And lngCurrentIndex <> RetornaChaveNo(node.Key) _
    'Then CarregaInfTurnoArrec node, Val(node.Key)
    
    'If node.Tag = "Publisher" Then PopStatus node

    

End Sub
'''
'''Private Function CarregaInfTurnoArrec(ByRef ParentNode As MSComctlLib.node, _
'''                                    pCAIXAARRECID As String) As Boolean
'''  On Error GoTo trata
'''  Dim newNode As MSComctlLib.node ' For new Node.
'''  Dim objRS As ADODB.Recordset
'''  Dim objGer As busSisMaq.clsGeral
'''  Dim strSql As String
'''
'''  ' Check that the node isn't already populated. If it is, then
'''  ' add only the ListItem objects to the ListView and exit.
'''  If ParentNode.Children Then
'''    'AddListItemsOnly LOCACAOID
'''    Exit Function
'''  End If
'''  '
'''  Set objGer = New busSisMaq.clsGeral
'''  strSql = "SELECT BOLETOARREC.PKID, BOLETOARREC.NUMERO, BOLETOARREC.DATAENTRADA " & _
'''    " FROM BOLETOARREC " & _
'''    " WHERE BOLETOARREC.CAIXAARRECID=" & RetornaChaveNo(pCAIXAARRECID)
'''  '
'''  Set objRS = objGer.ExecutarSQL(strSql)
'''  '
'''  If objRS.EOF Then
'''    ' If no results, return a false and exit
'''    CarregaInfTurnoArrec = False
'''    Exit Function
'''  Else
'''    Do While Not objRS.EOF
'''      AddNode newNode, _
'''              ParentNode, _
'''              "BOLETOARRECID_" & objRS.Fields("PKID").Value, _
'''              "Boleto Nro " & objRS.Fields("NUMERO").Value, _
'''              "vertexto", _
'''              "DETALHAR"
'''
'''      objRS.MoveNext
'''    Loop
''''''    If objRS.Fields("OCUPADO").Value Then
''''''      AddNode newNode, _
''''''              ParentNode, _
''''''              "A_" & RetornaChaveNo(pLocacao), _
''''''              "Ocupada", _
''''''              "check", _
''''''              "DETALHAR"
''''''    End If
''''''    AddNode newNode, _
''''''            ParentNode, _
''''''            "B_" & RetornaChaveNo(pLocacao), _
''''''            "Entrada/Depósito", _
''''''            "vertexto", _
''''''            "DETALHAR"
''''''    AddNode newNode, _
''''''            ParentNode, _
''''''            "C_" & RetornaChaveNo(pLocacao), _
''''''            "Fechamento", _
''''''            IIf(objRS.Fields("FECHAMENTO").Value = True, "detalheok", "detalhenaook"), _
''''''            "DETALHAR"
''''''    AddNode newNode, _
''''''            ParentNode, _
''''''            "D_" & RetornaChaveNo(pLocacao), _
''''''            "Recebimento" & IIf(objRS.Fields("NOME_GARCOM").Value & "" = "<Não informado>", "", " - " & objRS.Fields("NOME_GARCOM").Value), _
''''''            IIf(objRS.Fields("RECEBIMENTO").Value = True, "detalheok", "detalhenaook"), _
''''''            "DETALHAR"
''''''    AddNode newNode, _
''''''            ParentNode, _
''''''            "N_" & RetornaChaveNo(pLocacao), _
''''''            "Recebimento Empresa", _
''''''            "detalheok", _
''''''            "DETALHAR"
''''''
''''''    AddNode newNode, _
''''''            ParentNode, _
''''''            "E_" & RetornaChaveNo(pLocacao), _
''''''            "Liberação" & IIf(objRS.Fields("NOME_CAMAREIRALIBERACAO").Value & "" = "", "", " - " & objRS.Fields("NOME_CAMAREIRALIBERACAO").Value), _
''''''            IIf(objRS.Fields("LIBERADO").Value = True, "check", "nocheck"), _
''''''            "DETALHAR"
''''''    AddNode newNode, _
''''''            ParentNode, _
''''''            "F_" & RetornaChaveNo(pLocacao), _
''''''            "Saída" & IIf(objRS.Fields("SAIU").Value = True, "-" & Format(objRS.Fields("DTHORASAIDA").Value, "DD/MM hh:mm"), ""), _
''''''            IIf(objRS.Fields("SAIU").Value = True, "check", "nocheck"), _
''''''            "DETALHAR"
''''''    AddNode newNode, _
''''''            ParentNode, _
''''''            "G_" & RetornaChaveNo(pLocacao), _
''''''            "Suite/Apto. Limpo" & IIf(objRS.Fields("NOME_CAMAREIRALIMPEZA").Value & "" = "", "", " - " & objRS.Fields("NOME_CAMAREIRALIMPEZA").Value), _
''''''            IIf(objRS.Fields("LIMPO").Value = True, "check", "nocheck"), _
''''''            "DETALHAR"
''''''    AddNode newNode, _
''''''            ParentNode, _
''''''            "H_" & RetornaChaveNo(pLocacao), _
''''''            "Pedidos", _
''''''            IIf(objRS.Fields("QTDPEDIDO").Value = 0, "detalhenaook", "detalheok"), _
''''''            "DETALHAR"
''''''    AddNode newNode, _
''''''            ParentNode, _
''''''            "I_" & RetornaChaveNo(pLocacao), _
''''''            "Extras", _
''''''            IIf(objRS.Fields("QTDEXTRA").Value = 0, "detalhenaook", "detalheok"), _
''''''            "DETALHAR"
''''''    AddNode newNode, _
''''''            ParentNode, _
''''''            "J_" & RetornaChaveNo(pLocacao), _
''''''            "Lembrete", _
''''''            "vertexto", _
''''''            "DETALHAR"
''''''    AddNode newNode, _
''''''            ParentNode, _
''''''            "K_" & RetornaChaveNo(pLocacao), _
''''''            "Telefonema", _
''''''            "vertexto", _
''''''            "DETALHAR"
''''''    AddNode newNode, _
''''''            ParentNode, _
''''''            "L_" & RetornaChaveNo(pLocacao), _
''''''            "Serviço Despertador", _
''''''            "vertexto", _
''''''            "DETALHAR"
'''    'AddListItem mItem, _
'''                rsTitles
'''  End If
'''
''''''  ' Add a first node.
''''''  AddNode newNode, ParentNode, rsTitles
''''''  ' Add a corresponding ListItem.
''''''  AddListItem mItem, rsTitles
'''
'''
'''  ' Go through the rest of the recordset. If the next record
'''  ' is a duplicate, then just add the author's name.
'''  ' Otherwise, add a new Node and ListItem.
''''''  Do While Not objRS.EOF
'''''''''      ' Check the Key against the current ISDN. If they are the same
'''''''''      ' then the record only differs by containing a different
'''''''''      ' author. So add the author to the current list.
'''''''''      If newNode.Key = rsTitles!ISBN Then
'''''''''          ' Add the author to the list of authors.
'''''''''          mItem.ListSubItems("author").Text = _
'''''''''          mItem.ListSubItems("author").Text & _
'''''''''          " & " & rsTitles!author
'''''''''      Else ' Add a new Node and ListItem
'''''''''          AddNode newNode, ParentNode, rsTitles
'''''''''          AddListItem mItem, rsTitles
'''''''''      End If
''''''      objRS.MoveNext
''''''  Loop
'''  CarregaInfTurnoArrec = True ' return true for success
'''
'''  ' Sort the Turnos nodes.
'''  tvwDB.Nodes(pCAIXAARRECID).Sorted = False
'''  ' Expand top node.
'''  tvwDB.Nodes(pCAIXAARRECID).Expanded = True
'''
'''  lngCurrentIndex = RetornaChaveNo(pCAIXAARRECID)
'''
'''  Set objGer = Nothing
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, "[frmUserDetalhamentoArrec.CarregarInfUnidades]", Err.Description
'''End Function
Private Sub AddNode(ByRef newNode As MSComctlLib.node, _
                    ByRef ParentNode As MSComctlLib.node, _
                    Key As String, _
                    Texto As String, _
                    Imagem As String, _
                    Tag As String)
    ' Add a new node. The newNode and ParentNode are both needed.
    If Imagem <> "" Then
      Set newNode = tvwDB.Nodes.Add(ParentNode, _
        tvwChild, Key, Texto, Imagem)
    Else
      Set newNode = tvwDB.Nodes.Add(ParentNode, _
        tvwChild, Key, Texto)
    End If
    newNode.Tag = Tag
End Sub

Private Sub tvwDB_NodeClick(ByVal node As MSComctlLib.node)
    ' If the Tag is "Publisher" and the mItelngCurrentIndex
    ' index isn't the same as the Node.key, then
    ' invoke the CarregaInfTurnoArrec function, which populates the Node.
    If node.Tag = "TURNO" And lngCurrentIndex <> RetornaChaveNo(node.Key) _
    Then CarregarTurnoArrec node.Key, RetornaChaveNo(node.Key), node
    
    If node.Tag = "CAIXAARREC" And lngCurrentIndex <> RetornaChaveNo(node.Key) _
    Then DetalharTurnoArrec RetornaChaveNo(node.Key), node
    
    If node.Tag = "DETALHAR" Then DetalharTurnoArrec node.Key, node
    'If node.Tag = "Publisher" Then PopStatus node
    node.Sorted = False
        
    
End Sub
 

