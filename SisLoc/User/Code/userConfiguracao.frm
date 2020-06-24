VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserConfiguracao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuração do Sistema"
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2985
         Left            =   90
         ScaleHeight     =   2925
         ScaleWidth      =   1575
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1635
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   150
            Width           =   1335
         End
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1020
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1890
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
               Picture         =   "userConfiguracao.frx":0000
               Key             =   "closed"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userConfiguracao.frx":059C
               Key             =   "closedgreen"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userConfiguracao.frx":0B38
               Key             =   "open"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userConfiguracao.frx":10D4
               Key             =   "opengreen"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userConfiguracao.frx":1670
               Key             =   "detalhenaook"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userConfiguracao.frx":198C
               Key             =   "detalheok"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userConfiguracao.frx":1CA8
               Key             =   "nocheck"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userConfiguracao.frx":2584
               Key             =   "vertexto"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userConfiguracao.frx":28A0
               Key             =   "check"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userConfiguracao.frx":317C
               Key             =   "leaf"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userConfiguracao.frx":32EE
               Key             =   "open_"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "userConfiguracao.frx":3460
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
Attribute VB_Name = "frmUserConfiguracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mNode                 As MSComctlLib.node ' Module-level Node variable.
Private mItem                 As MSComctlLib.ListItem ' Module-level ListItem variable.
Private lngCurrentIndex       As Long ' Flag to assure this node wasn't already clicked.
Private lngCONFIGID     As Long
'Public strSqlWhereTurno       As String
'Public strSqlWhereLocacao     As String


Public Function RetornaChaveNo(strChave As String) As Long
  On Error GoTo trata:
  Dim strRetorno As String
  strRetorno = Replace(strChave, "CONFIGURACAOID_", "")
  strRetorno = Replace(strRetorno, "ITEMCONFIGID_", "")
  strRetorno = Replace(strRetorno, "A_", "")
  strRetorno = Replace(strRetorno, "B_", "")
  strRetorno = Replace(strRetorno, "C_", "")
  strRetorno = Replace(strRetorno, "D_", "")
  strRetorno = Replace(strRetorno, "E_", "")
  strRetorno = Replace(strRetorno, "F_", "")
  strRetorno = Replace(strRetorno, "G_", "")
  strRetorno = Replace(strRetorno, "H_", "")
  strRetorno = Replace(strRetorno, "I_", "")
  strRetorno = Replace(strRetorno, "J_", "")
  strRetorno = Replace(strRetorno, "K_", "")
  strRetorno = Replace(strRetorno, "L_", "")
  strRetorno = Replace(strRetorno, "M_", "")
  If Not IsNumeric(strRetorno) Then
    'Novo Tratamento
    strRetorno = Right(strRetorno, 1)
  End If
  RetornaChaveNo = CLng(strRetorno)
  Exit Function
trata:
  Err.Raise Err.Number, "[frmUserConfiguracao.RetornaChaveNo]", Err.Description
End Function

Private Sub CarregarConfiguracoes(pConfiguracao As String, _
                                  Optional ByRef ParentNode As MSComctlLib.node)
  On Error GoTo trata:
  ' While the record is not the last record, add a ListItem object.
  ' Use the Name field for the ListItem object's text.
  Dim intX As Integer
  Dim strTexto As String
  '
  ' Check that the node isn't already populated. If it is, then
  ' add only the ListItem objects to the ListView and exit.
  On Error GoTo Sai
  If ParentNode.Children Then
    'AddListItemsOnly pConfiguracao
    Exit Sub
  End If
Sai:
  On Error GoTo trata:
  '
  'Set objRs = objLoc.ListarLocacaoPorUnidade(RetornaChaveNo(pConfiguracao), strSqlWhereLocacao)
  For intX = 1 To 1
    'Set mNode = tvwDB.Nodes.Add(pConfiguracao, tvwChild, "LOCACAOID_" & objRs.Fields("PKID").Value, objRs.Fields("NUMERO").Value, IIf(objRs.Fields("OCUPADO").Value = True, "closedgreen", "closed"))
    Select Case intX
    'Case 1: strTexto = "Impressão"
    'Case 2: strTexto = "Fechamento"
    'Case 3: strTexto = "Promoção/Cortesia"
    'Case 4: strTexto = "Locação"
    Case 1: strTexto = "Dados Cadastrais"
    'Case 6: strTexto = "Telefonema, Entrada e Pedido"
    'Case 7: strTexto = "Serv. Despertador, Diária e Mov. do Caixa"
    'Case 8: strTexto = "Tabelas Diversas"
    End Select
    Set mNode = tvwDB.Nodes.Add(pConfiguracao, tvwChild, pConfiguracao & "_ITEMCONFIGID_" & intX, strTexto, "closed")
    mNode.Tag = "ITEMCONFIG" ' Identifies the table.
    
  Next
  ' Sort the Turnos nodes.
  tvwDB.Nodes(pConfiguracao).Sorted = False
  ' Expand top node.
  tvwDB.Nodes(pConfiguracao).Expanded = True
  lngCurrentIndex = RetornaChaveNo(pConfiguracao)
  '
  Exit Sub
trata:
  Err.Raise Err.Number, "[frmUserConfiguracao.CarregarConfiguracoes]", Err.Description
End Sub

Private Sub DetalharItem(pConfigItem As String, _
                         pNode As MSComctlLib.node)
  On Error GoTo trata:
  Dim objForm   As Form
  Select Case CInt(RetornaChaveNo(pConfigItem)) - 1
'''  Case 0
'''    'impressão
'''    Set objForm = New SisLoc.frmUserConfigImpressao
'''    objForm.lngCONFIGID = RetornaChaveNo(pNode.Parent.Key)
'''    objForm.Show vbModal
'''  Case 1
'''    'Fechamento
'''    Set objForm = New SisLoc.frmUserConfigFechamento
'''    objForm.lngCONFIGID = RetornaChaveNo(pNode.Parent.Key)
'''    objForm.Show vbModal
'''  Case 2
'''    'Promoção/Cortesia
'''    Set objForm = New SisLoc.frmUserConfigCortesia
'''    objForm.lngCONFIGID = RetornaChaveNo(pNode.Parent.Key)
'''    objForm.Show vbModal
'''  Case 3
'''    'Locação
'''    Set objForm = New SisLoc.frmUserConfigLocacao
'''    objForm.lngCONFIGID = RetornaChaveNo(pNode.Parent.Key)
'''    objForm.Show vbModal
  Case 0
    'Dados Cadastrais
    Set objForm = New SisLoc.frmUserConfigCadastro
    objForm.lngCONFIGID = RetornaChaveNo(pNode.Parent.Key)
    objForm.Show vbModal
'''  Case 5
'''    'Telefonema, Entrada e Pedido
'''    Set objForm = New SisLoc.frmUserConfigTelEntrPed
'''    objForm.lngCONFIGID = RetornaChaveNo(pNode.Parent.Key)
'''    objForm.Show vbModal
'''  Case 6
'''    'Serv Despertador, Diária e Mov do Caixa
'''    Set objForm = New SisLoc.frmUserConfigDespDiaCaixa
'''    objForm.lngCONFIGID = RetornaChaveNo(pNode.Parent.Key)
'''    objForm.Show vbModal
'''  Case 7
'''    'Tabelas Diversas
'''    Set objForm = New SisLoc.frmUserConfigDiversos
'''    objForm.lngCONFIGID = RetornaChaveNo(pNode.Parent.Key)
'''    objForm.Show vbModal
  End Select
  If objForm.bRetorno = True Then
    Form_Load
  End If
  Set objForm = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, "[frmUserConfiguracao.DetalharItem]", Err.Description
End Sub

Private Sub cmdExcluir_Click()
  Dim strMsg                    As String
  Dim strMsgErro                As String
  Dim objConfiguracao           As busSisLoc.clsConfiguracao
  Dim strRetornoErro            As String
  '
  On Error GoTo trata
  '
  If lngCONFIGID = 0 Then
    MsgBox "Selecione uma configuração para exclui-la.", vbExclamation, TITULOSISTEMA
    SetarFoco tvwDB
    Exit Sub
  End If
  Set objConfiguracao = New busSisLoc.clsConfiguracao
  '
  If MsgBox("Deseja excluir a configuração selecionada?", vbYesNo, TITULOSISTEMA) = vbYes Then
    
    If Not objConfiguracao.VerificaExclusaoConfiguracao(lngCONFIGID, _
                                                        strRetornoErro) Then
      Set objConfiguracao = Nothing
      TratarErroPrevisto "Não é possível excluir a configuração, pois há referências na(s) tabela(s):" & vbCrLf & vbCrLf & strRetornoErro, "frmUserConfiguracao.cmdExcluir_Click"
      SetarFoco tvwDB
      Exit Sub
    End If
    'ok
    objConfiguracao.ExcluirConfiguracao lngCONFIGID
    'Reload
    Form_Load
  End If
  SetarFoco tvwDB
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserConfiguracao.cmdExcluir_Click]"
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub



Private Sub cmdIncluir_Click()
  Dim objConfig As busSisLoc.clsConfiguracao
  If MsgBox("Deseja Incluir uma nova configuração?", vbOKCancel, TITULOSISTEMA) = vbCancel Then Exit Sub
  Set objConfig = New busSisLoc.clsConfiguracao
  objConfig.InserirConfiguracao
  Set objConfig = Nothing
  Form_Load
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub
Private Sub Form_Load()
  Dim objConfig As busSisLoc.clsConfiguracao
  Dim objRs As ADODB.Recordset
  Dim lngTotal As Long
  On Error GoTo trata
  AmpS
  Me.Height = 6090
  Me.Width = 8520
  
  CenterForm Me
  'lngCONFIGID = RetornaCodTurnoCorrente
  lngCONFIGID = 0
  If gsNivel <> "ADM" Then
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
  End If
  Set objConfig = New busSisLoc.clsConfiguracao
  ' Configure TreeView
  tvwDB.Nodes.Clear
  With tvwDB
    .Sorted = False
    .LabelEdit = False
    .LineStyle = tvwRootLines
  End With
  Set objRs = objConfig.ListarConfiguracao
  lngTotal = 1
  Do While Not objRs.EOF
    Set mNode = tvwDB.Nodes.Add()
    With mNode ' Add node turno.
      .Text = objRs.Fields("EMPRESA").Value & ""
      .Key = "CONFIGURACAOID_" & objRs.Fields("PKID").Value
      .Tag = "CONFIGURACAO"
      .Image = IIf(lngCONFIGID = objRs.Fields("PKID").Value, "closedgreen", "closed")
    End With
    If lngTotal = 1 Then
      CarregarConfiguracoes "CONFIGURACAOID_" & objRs.Fields("PKID").Value
    End If
    lngTotal = lngTotal + 1
    objRs.MoveNext
  Loop
  objRs.Close
  Set objRs = Nothing
  
  'CarregarConfiguracoes "TURNOID_" & RetornaCodTurnoCorrente
  Set objConfig = Nothing
  LerFiguras Me, tpBmp_Vazio, pbtnFechar:=cmdFechar, pbtnIncluir:=cmdIncluir, pbtnExcluir:=cmdExcluir
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub



Private Sub tvwDB_Collapse(ByVal node As MSComctlLib.node)
    ' Only nodes that are folders can be collapsed.
    If node.Tag = "LOCACAO" Then node.Image = IIf(node.Image = "opengreen", "closedgreen", "closed")
    If node.Tag = "TURNO" Then node.Image = IIf(RetornaChaveNo(node.Key) = lngCONFIGID, "closedgreen", "closed")
End Sub

Private Sub tvwDB_Expand(ByVal node As MSComctlLib.node)
    ' Only the top node, and publisher nodes can be expanded.

    node.Sorted = False
    If node.Tag = "LOCACAO" Then node.Image = IIf(node.Image = "closedgreen", "opengreen", "open")
    If node.Tag = "TURNO" Then node.Image = IIf(RetornaChaveNo(node.Key) = lngCONFIGID, "opengreen", "open")
    '
    ' If the Tag is "Publisher" and the mItelngCurrentIndex
    ' index isn't the same as the Node.key, then
    ' invoke the CarregaInfItem function.
    If node.Tag = "LOCACAO" And lngCurrentIndex <> RetornaChaveNo(node.Key) _
    Then CarregaInfItem node, Val(node.Key)
    
    'If node.Tag = "Publisher" Then PopStatus node

    

End Sub

Private Function CarregaInfItem(ByRef ParentNode As MSComctlLib.node, _
                                pItemConfig As String) As Boolean
  On Error GoTo trata
  Dim newNode As MSComctlLib.node ' For new Node.
  Dim objForm As Form
  Dim objControl As Control
  
  ' Check that the node isn't already populated. If it is, then
  ' add only the ListItem objects to the ListView and exit.
  If ParentNode.Children Then
    'AddListItemsOnly LOCACAOID
    Exit Function
  End If
  '
  '
  '
  'If objRs.EOF Then
  '  ' If no results, return a false and exit
  '  CarregaInfItem = False
  '  Exit Function
  
  Select Case CInt(RetornaChaveNo(pItemConfig))
  Case 1: 'Set objForm = New SisLoc.frmUserConfigImpressao
  Case 2: 'Set objForm = New SisLoc.frmUserConfigFechamento
  Case 3:
  Case 4:
  Case 5:
  Case 6:
  Case 7:
  Case 8:
  End Select
  
  For Each objControl In objForm.Controls
    'objControl.Typeof
    
    MsgBox objControl.Name
  Next
  
  
'''    If objRs.Fields("OCUPADO").Value Then
'''      AddNode newNode, _
'''              ParentNode, _
'''              "A_" & RetornaChaveNo(pItemConfig), _
'''              "Ocupada", _
'''              "check", _
'''              "DETALHAR"
'''    End If
'''    AddNode newNode, _
'''            ParentNode, _
'''            "B_" & RetornaChaveNo(pItemConfig), _
'''            "Entrada/Depósito", _
'''            "vertexto", _
'''            "DETALHAR"
'''    AddNode newNode, _
'''            ParentNode, _
'''            "C_" & RetornaChaveNo(pItemConfig), _
'''            "Fechamento", _
'''            IIf(objRs.Fields("FECHAMENTO").Value = True, "detalheok", "detalhenaook"), _
'''            "DETALHAR"
'''    AddNode newNode, _
'''            ParentNode, _
'''            "D_" & RetornaChaveNo(pItemConfig), _
'''            "Recebimento", _
'''            IIf(objRs.Fields("RECEBIMENTO").Value = True, "detalheok", "detalhenaook"), _
'''            "DETALHAR"
'''    AddNode newNode, _
'''            ParentNode, _
'''            "E_" & RetornaChaveNo(pItemConfig), _
'''            "Liberação" & IIf(objRs.Fields("NOME_CAMAREIRALIBERACAO").Value & "" = "", "", " - " & objRs.Fields("NOME_CAMAREIRALIBERACAO").Value), _
'''            IIf(objRs.Fields("LIBERADO").Value = True, "check", "nocheck"), _
'''            "DETALHAR"
'''    AddNode newNode, _
'''            ParentNode, _
'''            "F_" & RetornaChaveNo(pItemConfig), _
'''            "Saída" & IIf(objRs.Fields("SAIU").Value = True, "-" & Format(objRs.Fields("DTHORASAIDA").Value, "DD/MM hh:mm"), ""), _
'''            IIf(objRs.Fields("SAIU").Value = True, "check", "nocheck"), _
'''            "DETALHAR"
'''    AddNode newNode, _
'''            ParentNode, _
'''            "G_" & RetornaChaveNo(pItemConfig), _
'''            "Suite/Apto. Limpo" & IIf(objRs.Fields("NOME_CAMAREIRALIMPEZA").Value & "" = "", "", " - " & objRs.Fields("NOME_CAMAREIRALIMPEZA").Value), _
'''            IIf(objRs.Fields("LIMPO").Value = True, "check", "nocheck"), _
'''            "DETALHAR"
'''    AddNode newNode, _
'''            ParentNode, _
'''            "H_" & RetornaChaveNo(pItemConfig), _
'''            "Pedidos", _
'''            IIf(objRs.Fields("QTDPEDIDO").Value = 0, "detalhenaook", "detalheok"), _
'''            "DETALHAR"
'''    AddNode newNode, _
'''            ParentNode, _
'''            "I_" & RetornaChaveNo(pItemConfig), _
'''            "Extras", _
'''            IIf(objRs.Fields("QTDEXTRA").Value = 0, "detalhenaook", "detalheok"), _
'''            "DETALHAR"
'''    AddNode newNode, _
'''            ParentNode, _
'''            "J_" & RetornaChaveNo(pItemConfig), _
'''            "Lembrete", _
'''            "vertexto", _
'''            "DETALHAR"
'''    AddNode newNode, _
'''            ParentNode, _
'''            "K_" & RetornaChaveNo(pItemConfig), _
'''            "Telefonema", _
'''            "vertexto", _
'''            "DETALHAR"
'''    AddNode newNode, _
'''            ParentNode, _
'''            "L_" & RetornaChaveNo(pItemConfig), _
'''            "Serviço Despertador", _
'''            "vertexto", _
'''            "DETALHAR"
'''    'AddListItem mItem, _
'''                rsTitles
  Set objForm = Nothing
  CarregaInfItem = True ' return true for success
  
  ' Sort the Turnos nodes.
  tvwDB.Nodes(pItemConfig).Sorted = False
  ' Expand top node.
  tvwDB.Nodes(pItemConfig).Expanded = True
  
  lngCurrentIndex = RetornaChaveNo(pItemConfig)
    
  Exit Function
trata:
  Err.Raise Err.Number, "[frmUserConfiguracao.CarregarInfUnidades]", Err.Description
End Function
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
  On Error Resume Next
    ' If the Tag is "Publisher" and the mItelngCurrentIndex
    ' index isn't the same as the Node.key, then
    ' invoke the CarregaInfItem function, which populates the Node.
    If node.Tag = "CONFIGURACAO" Then
      lngCONFIGID = RetornaChaveNo(node.Key)
    End If
    If node.Tag = "CONFIGURACAO" And lngCurrentIndex <> RetornaChaveNo(node.Key) _
    Then CarregarConfiguracoes node.Key, node
    
    'If node.Tag = "ITEMCONFIG" And lngCurrentIndex <> RetornaChaveNo(node.Key) _
    'Then CarregaInfItem node, node.Key
    If node.Tag = "ITEMCONFIG" Then DetalharItem node.Key, node
    
    
    'If node.Tag = "DETALHAR" Then DetalharItem node.Key, node
    'If node.Tag = "Publisher" Then PopStatus node
    node.Sorted = False
        
    
End Sub
 

