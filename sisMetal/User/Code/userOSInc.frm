VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOSInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de OS"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   KeyPreview      =   -1  'True
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4695
         Left            =   90
         ScaleHeight     =   4635
         ScaleWidth      =   1605
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   720
         Width           =   1665
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&X"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   3600
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   2730
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5295
      Left            =   120
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da OS"
      TabPicture(0)   =   "userOSInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraProf"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Itens da OS"
      TabPicture(1)   =   "userOSInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdOS"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraProf 
         Height          =   3972
         Left            =   120
         TabIndex        =   21
         Top             =   330
         Width           =   7935
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   3612
            Index           =   0
            Left            =   120
            ScaleHeight     =   3615
            ScaleWidth      =   7695
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   240
            Width           =   7695
            Begin VB.Frame Frame1 
               Height          =   1545
               Left            =   0
               TabIndex        =   31
               Top             =   2010
               Width           =   7695
               Begin VB.TextBox txtConferente 
                  Height          =   285
                  Left            =   1230
                  MaxLength       =   50
                  TabIndex        =   11
                  Text            =   "txtConferente"
                  Top             =   1110
                  Width           =   5745
               End
               Begin VB.TextBox txtOperador 
                  Height          =   285
                  Left            =   1230
                  MaxLength       =   50
                  TabIndex        =   10
                  Text            =   "txtOperador"
                  Top             =   780
                  Width           =   5745
               End
               Begin MSMask.MaskEdBox mskPesoBruto 
                  Height          =   255
                  Left            =   1230
                  TabIndex        =   6
                  Top             =   180
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  _Version        =   393216
                  Format          =   "#,##0.000;($#,##0.000)"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskPesoLiquido 
                  Height          =   255
                  Left            =   5280
                  TabIndex        =   7
                  Top             =   180
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  _Version        =   393216
                  Format          =   "#,##0.000;(#,##0.000)"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskVrMetal 
                  Height          =   255
                  Left            =   1230
                  TabIndex        =   8
                  Top             =   480
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  _Version        =   393216
                  Format          =   "#,##0.00;(#,##0.00)"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskValor 
                  Height          =   255
                  Left            =   5280
                  TabIndex        =   9
                  Top             =   480
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  _Version        =   393216
                  Format          =   "#,##0.00;($#,##0.00)"
                  PromptChar      =   "_"
               End
               Begin VB.Label Label2 
                  Caption         =   "Conferente"
                  Height          =   255
                  Index           =   9
                  Left            =   90
                  TabIndex        =   37
                  Top             =   1110
                  Width           =   1155
               End
               Begin VB.Label Label2 
                  Caption         =   "Operador"
                  Height          =   255
                  Index           =   8
                  Left            =   90
                  TabIndex        =   36
                  Top             =   780
                  Width           =   1155
               End
               Begin VB.Label Label2 
                  Caption         =   "Vr. Total"
                  Height          =   255
                  Index           =   7
                  Left            =   4140
                  TabIndex        =   35
                  Top             =   480
                  Width           =   1155
               End
               Begin VB.Label Label2 
                  Caption         =   "Vr. Kilo"
                  Height          =   255
                  Index           =   6
                  Left            =   90
                  TabIndex        =   34
                  Top             =   480
                  Width           =   1155
               End
               Begin VB.Label Label2 
                  Caption         =   "Peso Líq."
                  Height          =   255
                  Index           =   5
                  Left            =   4140
                  TabIndex        =   33
                  Top             =   180
                  Width           =   1155
               End
               Begin VB.Label Label2 
                  Caption         =   "Peso Bruto"
                  Height          =   255
                  Index           =   4
                  Left            =   90
                  TabIndex        =   32
                  Top             =   180
                  Width           =   1155
               End
            End
            Begin VB.TextBox txtNF 
               Height          =   285
               Left            =   1230
               MaxLength       =   50
               TabIndex        =   2
               Text            =   "txtNF"
               Top             =   630
               Width           =   2505
            End
            Begin VB.ComboBox cboFabrica 
               Height          =   315
               Left            =   1230
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   960
               Width           =   4515
            End
            Begin VB.ComboBox cboCor 
               Height          =   315
               Left            =   1230
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   1680
               Visible         =   0   'False
               Width           =   4515
            End
            Begin VB.ComboBox cboFornecedor 
               Height          =   315
               Left            =   1230
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   1320
               Width           =   4515
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   255
               Left            =   30
               ScaleHeight     =   255
               ScaleWidth      =   3855
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   30
               Width           =   3855
               Begin MSMask.MaskEdBox mskData 
                  Height          =   255
                  Index           =   0
                  Left            =   1200
                  TabIndex        =   0
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
                  Index           =   2
                  Left            =   30
                  TabIndex        =   29
                  Top             =   0
                  Width           =   1155
               End
               Begin VB.Label Label2 
                  Caption         =   "Data"
                  Height          =   255
                  Index           =   1
                  Left            =   0
                  TabIndex        =   24
                  Top             =   -360
                  Width           =   615
               End
            End
            Begin MSMask.MaskEdBox mskNumero 
               Height          =   255
               Left            =   1230
               TabIndex        =   1
               Top             =   330
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label2 
               Caption         =   "NF"
               Height          =   255
               Index           =   3
               Left            =   30
               TabIndex        =   30
               Top             =   630
               Width           =   1155
            End
            Begin VB.Label Label5 
               Caption         =   "Fábrica"
               Height          =   195
               Index           =   2
               Left            =   30
               TabIndex        =   28
               Top             =   990
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Anodização"
               Height          =   195
               Index           =   1
               Left            =   30
               TabIndex        =   27
               Top             =   1710
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Fornecedor"
               Height          =   195
               Index           =   0
               Left            =   30
               TabIndex        =   26
               Top             =   1350
               Width           =   1215
            End
            Begin VB.Label Label2 
               Caption         =   "Número"
               Height          =   255
               Index           =   0
               Left            =   30
               TabIndex        =   25
               Top             =   330
               Width           =   1155
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdOS 
         Height          =   4455
         Left            =   -74910
         OleObjectBlob   =   "userOSInc.frx":0038
         TabIndex        =   17
         Top             =   420
         Width           =   7995
      End
   End
End
Attribute VB_Name = "frmOSInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                 As tpStatus
Public lngOSID                As Long
Public blnRetorno             As Boolean
Public blnFechar              As Boolean
Private blnPrimeiraVez        As Boolean
Private lngFORNECEDORID       As Long
Dim OS_COLUNASMATRIZ          As Long
Dim OS_LINHASMATRIZ           As Long
Private OS_Matriz()           As String



Private Sub cboCor_LostFocus()
  Pintar_Controle cboCor, tpCorContr_Normal
End Sub

Private Sub cboFabrica_LostFocus()
  Pintar_Controle cboFabrica, tpCorContr_Normal
End Sub

Private Sub cboFornecedor_LostFocus()
  Pintar_Controle cboFornecedor, tpCorContr_Normal
End Sub

Private Sub cmdAlterar_Click()
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 1 'Itens da OS
    If Len(Trim(grdOS.Columns("PKID").Value & "")) = 0 Then
      MsgBox "Selecione um item da OS!", vbExclamation, TITULOSISTEMA
      Exit Sub
    End If
    frmItemOSInc.Status = tpStatus_Alterar
    frmItemOSInc.lngPKID = grdOS.Columns("PKID").Value
    frmItemOSInc.lngOSID = lngOSID
    frmItemOSInc.lngFORNECEDORID = lngFORNECEDORID
    frmItemOSInc.strFornecedor = cboFornecedor.Text
    frmItemOSInc.strNumero = mskNumero.Text
    frmItemOSInc.strData = mskData(0).Text
    frmItemOSInc.Show vbModal

    If frmItemOSInc.blnRetorno Then
      OS_MontaMatriz
      grdOS.Bookmark = Null
      grdOS.ReBind
      grdOS.ApproxCount = OS_LINHASMATRIZ
    End If
    SetarFoco grdOS
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub cmdExcluir_Click()
  Dim objItemOS     As busSisMetal.clsItemOS
  Dim objGer        As busSisMetal.clsGeral
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  '
  On Error GoTo trata
  If Len(Trim(grdOS.Columns("PKID").Value & "")) = 0 Then
    MsgBox "Selecione um item da OS para exclusão.", vbExclamation, TITULOSISTEMA
    Exit Sub
  End If
  '
'''  Set objGer = New busSisMetal.clsGeral
'''  'ITEM_PEDIDO
'''  strSql = "Select * from ITEM_PEDIDO WHERE PEDIDOID = " & grdOS.Columns("PKID").Value
'''  Set objRs = objGer.ExecutarSQL(strSql)
'''  If Not objRs.EOF Then
'''    objRs.Close
'''    Set objRs = Nothing
'''    Set objGer = Nothing
'''    TratarErroPrevisto "OS não pode ser excluido pois já possui itens lançados.", "frmOSLis.cmdExcluir_Click"
'''    SetarFoco grdGeral
'''    Exit Sub
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
'''  '
'''  Set objGer = Nothing
'''  '
  '
  If MsgBox("Confirma exclusão do item da OS " & grdOS.Columns("Linha-perfil").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then Exit Sub
'''  'OK
'''  Set objItemOS = New busSisMetal.clsItemOS
'''
'''  objItemOS.ExcluirItemOS CLng(grdOS.Columns("PKID").Value)
'''  '
'''  OS_MontaMatriz
'''  grdOS.Bookmark = Null
'''  grdOS.ReBind
'''
'''  Set objItemOS = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdIncluir_Click()
  On Error GoTo trata
  Dim objForm As Form
  '
  Select Case tabDetalhes.Tab
  Case 1 'Itens da OS
    frmItemOSInc.Status = tpStatus_Incluir
    frmItemOSInc.lngPKID = 0
    frmItemOSInc.lngOSID = lngOSID
    frmItemOSInc.lngFORNECEDORID = lngFORNECEDORID
    frmItemOSInc.strFornecedor = cboFornecedor.Text
    frmItemOSInc.strNumero = mskNumero.Text
    frmItemOSInc.strData = mskData(0).Text
    frmItemOSInc.Show vbModal

    If frmItemOSInc.blnRetorno Then
      OS_MontaMatriz
      grdOS.Bookmark = Null
      grdOS.ReBind
      grdOS.ApproxCount = OS_LINHASMATRIZ
    End If
    SetarFoco grdOS
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdOK_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim objOS                   As busSisMetal.clsOS
  Dim objRs                   As ADODB.Recordset
  Dim objGer                  As busSisMetal.clsGeral
  Dim lngFORNECEDORID         As Long
  Dim lngCORID                As Long
  Dim lngFABRICAID            As Long
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração da OS
    If Not ValidaCampos Then Exit Sub
'''    'Valida se pedido já cadastrada
'''    Set objGer = New busSisMetal.clsGeral
'''    strSql = "Select * From LINHA WHERE (NOME = " & Formata_Dados(txtNome.Text, tpDados_Texto, tpNulo_Aceita) & _
'''      " OR CODIGO = " & Formata_Dados(txtCodigo.Text, tpDados_Texto, tpNulo_Aceita) & ") " & _
'''      " AND PKID <> " & Formata_Dados(lngLINHAID, tpDados_Longo, tpNulo_NaoAceita)
'''    Set objRs = objGer.ExecutarSQL(strSql)
'''    If Not objRs.EOF Then
'''      objRs.Close
'''      Set objRs = Nothing
'''      Set objGer = Nothing
'''      TratarErroPrevisto "Nome ou Código da linha já cadastrada", "cmdOK_Click"
'''      Pintar_Controle txtNome, tpCorContr_Erro
'''      Pintar_Controle txtCodigo, tpCorContr_Erro
'''      SetarFoco txtNome
'''      Exit Sub
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
'''    '
    Set objGer = New busSisMetal.clsGeral
    'FORNECEDOR
    lngFORNECEDORID = 0
    strSql = "SELECT LOJA.PKID FROM LOJA " & _
      " INNER JOIN FORNECEDOR ON FORNECEDOR.LOJAID = LOJA.PKID " & _
      " WHERE LOJA.NOME = " & Formata_Dados(cboFornecedor.Text, tpDados_Texto)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngFORNECEDORID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    '
    'ANODIZADORA
    lngCORID = 0
    strSql = "SELECT COR.PKID FROM COR " & _
      " WHERE COR.NOME = " & Formata_Dados(cboCor.Text, tpDados_Texto)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngCORID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    '
    'FABRICA
    lngFABRICAID = 0
    strSql = "SELECT LOJA.PKID FROM LOJA " & _
      " INNER JOIN FABRICA ON FABRICA.LOJAID = LOJA.PKID " & _
      " WHERE LOJA.NOME = " & Formata_Dados(cboFabrica.Text, tpDados_Texto)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngFABRICAID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objGer = Nothing
    '
    Set objOS = New busSisMetal.clsOS
    'Altera ou incluiu pedido
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objOS.AlterarOS lngOSID, _
                      lngFORNECEDORID, _
                      lngCORID, _
                      lngFABRICAID, _
                      IIf(Len(mskNumero.ClipText) = 0, "", mskNumero.Text), _
                      txtNF.Text, _
                      IIf(Len(mskPesoBruto.ClipText) = 0, "", mskPesoBruto.Text), _
                      IIf(Len(mskPesoLiquido.ClipText) = 0, "", mskPesoLiquido.Text), _
                      IIf(Len(mskVrMetal.ClipText) = 0, "", mskVrMetal.Text), _
                      IIf(Len(mskValor.ClipText) = 0, "", mskValor.Text), _
                      txtOperador.Text, _
                      txtConferente.Text
      'Set objOS = Nothing
      '
      blnRetorno = True
      Status = tpStatus_Alterar
      'Reload na tela
      Form_Load
      'Acerta tabs
      'tabDetalhes.TabVisible(2) = True
      tabDetalhes.Tab = 1
      cmdIncluir_Click
      
    ElseIf Status = tpStatus_Incluir Then
      'Código para inclusão
      '
      objOS.InserirOS lngOSID, _
                      lngFORNECEDORID, _
                      lngCORID, _
                      lngFABRICAID, _
                      IIf(Len(mskNumero.ClipText) = 0, "", mskNumero.Text), _
                      txtNF.Text, _
                      IIf(Len(mskPesoBruto.ClipText) = 0, "", mskPesoBruto.Text), _
                      IIf(Len(mskPesoLiquido.ClipText) = 0, "", mskPesoLiquido.Text), _
                      IIf(Len(mskVrMetal.ClipText) = 0, "", mskVrMetal.Text), _
                      IIf(Len(mskValor.ClipText) = 0, "", mskValor.Text), _
                      txtOperador.Text, _
                      txtConferente.Text
      'Set objOS = Nothing
      '
      blnRetorno = True
      Status = tpStatus_Alterar
      'Reload na tela
      Form_Load
      'Acerta tabs
      'tabDetalhes.TabVisible(2) = True
      tabDetalhes.Tab = 1
      cmdIncluir_Click
    End If
    Set objOS = Nothing
    'blnFechar = True
    'Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  Dim objOS     As busSisMetal.clsOS
  '
  blnFechar = False
  blnRetorno = False
  AmpS
  Me.Height = 5985
  Me.Width = 10470
  CenterForm Me
  blnPrimeiraVez = True
  lngFORNECEDORID = 0
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, cmdExcluir, , cmdIncluir, cmdAlterar
  '
  LimparCampos
  tabDetalhes_Click 0
  '
  'Fornecedor
  strSql = "SELECT LOJA.NOME FROM LOJA " & _
      " INNER JOIN FORNECEDOR ON LOJA.PKID = FORNECEDOR.LOJAID " & _
      " WHERE LOJA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
      "ORDER BY LOJA.NOME"
  PreencheCombo cboFornecedor, strSql, False, True
  'Cor
  strSql = "SELECT COR.NOME FROM COR " & _
      "ORDER BY COR.NOME"
  PreencheCombo cboCor, strSql, False, True
  'Fabrica
  strSql = "SELECT LOJA.NOME FROM LOJA " & _
      " INNER JOIN FABRICA ON LOJA.PKID = FABRICA.LOJAID " & _
      " WHERE LOJA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
      "ORDER BY LOJA.NOME"
  PreencheCombo cboFabrica, strSql, False, True
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o OS
    'NÃO PERMITE ALTERAR O FORNECEDOR DEVIDO AOS ÍTENS LANÇADOS
    '------------------------------------
    cboFornecedor.Enabled = True
    Label5(0).Enabled = True
    '------------------------------------
    '
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = False
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objOS = New busSisMetal.clsOS
    Set objRs = objOS.ListarOS(lngOSID)
    '
    If Not objRs.EOF Then
      'Campos fixos
      INCLUIR_VALOR_NO_MASK mskData(0), objRs.Fields("DATA").Value, TpMaskData
      'Campos inserts
      INCLUIR_VALOR_NO_MASK mskNumero, objRs.Fields("NUMERO").Value, TpMaskLongo
      txtNF.Text = objRs.Fields("NF").Value & ""
      INCLUIR_VALOR_NO_COMBO objRs.Fields("NOME_FORNECEDOR").Value & "", cboFornecedor
      INCLUIR_VALOR_NO_COMBO objRs.Fields("NOME_COR").Value & "", cboCor
      INCLUIR_VALOR_NO_COMBO objRs.Fields("NOME_FABRICA").Value & "", cboFabrica
      INCLUIR_VALOR_NO_MASK mskPesoBruto, objRs.Fields("PESOBRUTO").Value, TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskPesoLiquido, objRs.Fields("PESOLIQUIDO").Value, TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskVrMetal, objRs.Fields("VALORMETAL").Value, TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VALOR").Value, TpMaskMoeda
      txtOperador.Text = objRs.Fields("OPERADOR").Value & ""
      txtConferente.Text = objRs.Fields("CONFERENTE").Value & ""
      '
      lngFORNECEDORID = objRs.Fields("FORNECEDORID").Value
    End If
    'NÃO PERMITE ALTERAR O FORNECEDOR DEVIDO AOS ÍTENS LANÇADOS
    '------------------------------------
    cboFornecedor.Enabled = False
    Label5(0).Enabled = False
    '------------------------------------
    Set objOS = Nothing
    '
    If Status = tpStatus_Alterar Then
      tabDetalhes.TabEnabled(0) = True
      tabDetalhes.TabEnabled(1) = True
    ElseIf Status = tpStatus_Consultar Then
      tabDetalhes.TabEnabled(0) = False
      tabDetalhes.TabEnabled(1) = True
      tabDetalhes.Tab = 1
      tabDetalhes_Click 1
    End If
    '
  End If
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
End Sub

Private Sub cmdCancelar_Click()
  blnFechar = True
  blnRetorno = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub


Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'OS
  LimparCampoMask mskData(0)
  LimparCampoMask mskNumero
  LimparCampoTexto txtNF
  LimparCampoCombo cboFornecedor
  LimparCampoCombo cboCor
  LimparCampoCombo cboFabrica
  LimparCampoMask mskPesoBruto
  LimparCampoMask mskPesoLiquido
  LimparCampoMask mskVrMetal
  LimparCampoMask mskValor
  LimparCampoTexto txtOperador
  LimparCampoTexto txtConferente
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmOSInc.LimparCampos]", _
            Err.Description
End Sub

Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  '
  If Not Valida_Moeda(mskNumero, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Número da OS inválido" & vbCrLf
  End If
  'If Not Valida_String(txtNF, TpObrigatorio, blnSetarFocoControle) Then
  '  strMsg = strMsg & "Informar o número da NF" & vbCrLf
  'End If
  If Not Valida_String(cboFabrica, TpnaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar a fábrica" & vbCrLf
  End If
  If Not Valida_String(cboFornecedor, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o fornecedor" & vbCrLf
  End If
  If Not Valida_String(cboCor, TpnaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar a anodização" & vbCrLf
  End If
  '
  If Not Valida_Moeda(mskPesoBruto, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Peso bruto inválido" & vbCrLf
  End If
  If Not Valida_Moeda(mskPesoLiquido, TpnaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Peso líquido inválido" & vbCrLf
  End If
  If Not Valida_Moeda(mskVrMetal, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preço do kilo inválido" & vbCrLf
  End If
  If Not Valida_Moeda(mskValor, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Valor total inválido" & vbCrLf
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmOSInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[frmOSInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Setar foco
    SetarFoco mskNumero
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmOSInc.Form_Activate]"
End Sub

Public Sub OS_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGeral  As busSisMetal.clsGeral
  '
  On Error GoTo trata
  
  Set objGeral = New busSisMetal.clsGeral
  '
  strSql = "SELECT ITEM_OS.PKID, " & _
          "TIPO_LINHA.NOME + ' - ' + LINHA.CODIGO, ITEM_OS.QUANTIDADE, BAIXA_OS.PESO_OS, BAIXA_OS.PESO_OS * IsNull(OS.VALORMETAL, 0) " & _
          "FROM ITEM_OS " & _
          " LEFT JOIN (SELECT BAIXA_PEDIDO_OS.ITEM_OSID, " & _
          "   SUM(BAIXA_PEDIDO_OS.QUANTIDADE) AS QUANTIDADE_OS, " & _
          "   SUM(BAIXA_PEDIDO_OS.PESO) AS PESO_OS " & _
          "   From BAIXA_PEDIDO_OS " & _
          "   GROUP BY BAIXA_PEDIDO_OS.ITEM_OSID " & _
          "   ) AS BAIXA_OS ON BAIXA_OS.ITEM_OSID = ITEM_OS.PKID " & _
          " INNER JOIN OS ON OS.PKID = ITEM_OS.OSID " & _
          " LEFT JOIN LINHA ON LINHA.PKID = ITEM_OS.LINHAID " & _
          " LEFT JOIN TIPO_LINHA ON TIPO_LINHA.PKID = LINHA.TIPO_LINHAID " & _
          "WHERE ITEM_OS.OSID = " & Formata_Dados(lngOSID, tpDados_Longo) & _
          " ORDER BY ITEM_OS.PKID"
  '
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    OS_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim OS_Matriz(0 To OS_COLUNASMATRIZ - 1, 0 To OS_LINHASMATRIZ - 1)
  Else
    ReDim OS_Matriz(0 To OS_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To OS_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To OS_COLUNASMATRIZ - 1  'varre as colunas
          OS_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  '
  Set objGeral = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then Cancel = True
End Sub

Private Sub grdOS_UnboundReadDataEx( _
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
               Offset + intI, OS_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, OS_COLUNASMATRIZ, OS_LINHASMATRIZ, OS_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, OS_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmOSInc.grdGeral_UnboundReadDataEx]"
End Sub



Private Sub mskNumero_GotFocus()
  Seleciona_Conteudo_Controle mskNumero
End Sub
Private Sub mskNumero_LostFocus()
  Pintar_Controle mskNumero, tpCorContr_Normal
End Sub

Private Sub mskPesoBruto_GotFocus()
  Seleciona_Conteudo_Controle mskPesoBruto
End Sub
Private Sub mskPesoBruto_LostFocus()
  Pintar_Controle mskPesoBruto, tpCorContr_Normal
End Sub

Private Sub mskPesoLiquido_GotFocus()
  Seleciona_Conteudo_Controle mskPesoLiquido
End Sub

Private Sub Carrega_total()
  On Error GoTo trata
  Dim curVrTotal As Currency
  '
  If Not Valida_Moeda(mskPesoLiquido, TpObrigatorio, False, False, False) Then
    Exit Sub
  End If
  If Not Valida_Moeda(mskVrMetal, TpObrigatorio, False, False, False) Then
    Exit Sub
  End If
  'Campos carregados
  curVrTotal = CCur(mskPesoLiquido.Text) * CCur(mskVrMetal.Text)
  curVrTotal = Format(curVrTotal, "###,##0.00")
  INCLUIR_VALOR_NO_MASK mskValor, curVrTotal, TpMaskMoeda
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub
Private Sub mskPesoLiquido_LostFocus()
  On Error GoTo trata
  If Me.ActiveControl.Name = "cmdCancelar" Then Exit Sub
  Pintar_Controle mskPesoLiquido, tpCorContr_Normal
  '
  Carrega_total
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub mskValor_GotFocus()
  Seleciona_Conteudo_Controle mskValor
End Sub
Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub

Private Sub mskVrMetal_GotFocus()
  Seleciona_Conteudo_Controle mskVrMetal
End Sub
Private Sub mskVrMetal_LostFocus()
  On Error GoTo trata
  If Me.ActiveControl.Name = "cmdCancelar" Then Exit Sub
  Pintar_Controle mskVrMetal, tpCorContr_Normal
  '
  Carrega_total
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  
End Sub

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    grdOS.Enabled = False
    pictrava(0).Enabled = True
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    SetarFoco cboFornecedor
  Case 1
    'Itens pedido
    grdOS.Enabled = True
    pictrava(0).Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = True
    'Montar RecordSet
    OS_COLUNASMATRIZ = grdOS.Columns.Count
    OS_LINHASMATRIZ = 0
    OS_MontaMatriz
    grdOS.Bookmark = Null
    grdOS.ReBind
    grdOS.ApproxCount = OS_LINHASMATRIZ
    '
    SetarFoco grdOS
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "SisMetal.frmOSInc.tabDetalhes"
  AmpN
End Sub

Private Sub txtConferente_GotFocus()
  Seleciona_Conteudo_Controle txtConferente
End Sub
Private Sub txtConferente_LostFocus()
  Pintar_Controle txtConferente, tpCorContr_Normal
End Sub

Private Sub txtNF_GotFocus()
  Seleciona_Conteudo_Controle txtNF
End Sub
Private Sub txtNF_LostFocus()
  Pintar_Controle txtNF, tpCorContr_Normal
End Sub

Private Sub txtOperador_GotFocus()
  Seleciona_Conteudo_Controle txtOperador
End Sub
Private Sub txtOperador_LostFocus()
  Pintar_Controle txtOperador, tpCorContr_Normal
End Sub

