VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFamiliaProdutoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de família de produtos"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   8520
      ScaleHeight     =   2565
      ScaleWidth      =   1860
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   30
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   330
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   4048
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da família de produtos"
      TabPicture(0)   =   "userFamiliaProdutoInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picTrava(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.PictureBox picTrava 
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   0
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   7695
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   480
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
            Height          =   1215
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   7695
            Begin VB.ComboBox cboIPI 
               Height          =   315
               Left            =   5880
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   570
               Width           =   1695
            End
            Begin VB.ComboBox cboICMS 
               Height          =   315
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   570
               Width           =   1695
            End
            Begin VB.TextBox txtDescricao 
               Height          =   285
               Left            =   1200
               MaxLength       =   50
               TabIndex        =   0
               Text            =   "txtDescricao"
               Top             =   240
               Width           =   6375
            End
            Begin VB.Label Label11 
               Caption         =   "Perc. ICMS"
               Height          =   255
               Left            =   240
               TabIndex        =   12
               Top             =   630
               Width           =   1095
            End
            Begin VB.Label Label16 
               Caption         =   "Perc. IPI"
               Height          =   255
               Left            =   4650
               TabIndex        =   11
               Top             =   630
               Width           =   1095
            End
            Begin VB.Label Label6 
               Caption         =   "Descrição"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   10
               Top             =   240
               Width           =   855
            End
         End
      End
   End
End
Attribute VB_Name = "frmFamiliaProdutoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                     As tpStatus
Public lngFAMILIAPRODUTOSID       As Long
Public blnRetorno                 As Boolean
Public blnFechar                  As Boolean
Private blnPrimeiraVez            As Boolean



Private Sub cboICMS_LostFocus()
  Pintar_Controle cboICMS, tpCorContr_Normal
End Sub

Private Sub cboIPI_LostFocus()
  Pintar_Controle cboIPI, tpCorContr_Normal
End Sub

Private Sub cmdCancelar_Click()
  blnFechar = True
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim strMsgErro              As String
  Dim objRs                   As ADODB.Recordset
  Dim objFamiliaProduto       As busSisMetal.clsFamiliaProduto
  Dim objGeral                As busSisMetal.clsGeral
  Dim lngIPIID                As Long
  Dim lngICMSID               As Long
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração da Unidade de estoque
    If Not ValidaCampos Then Exit Sub
    'Valida se unidade de estoque já cadastrada
    Set objGeral = New busSisMetal.clsGeral
    strSql = "Select * From FAMILIAPRODUTOS WHERE DESCRICAO = " & Formata_Dados(txtDescricao.Text, tpDados_Texto, tpNulo_NaoAceita) & _
      " AND PKID <> " & Formata_Dados(lngFAMILIAPRODUTOSID, tpDados_Longo, tpNulo_NaoAceita)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      TratarErroPrevisto "Família de produtos já cadastrada", "cmdOK_Click"
      Pintar_Controle txtDescricao, tpCorContr_Erro
      SetarFoco txtDescricao
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    'IPIID
    lngIPIID = 0
    strSql = "SELECT PKID FROM IPI WHERE IPI = " & Formata_Dados(cboIPI.Text, tpDados_Moeda)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngIPIID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    'ICMSID
    lngICMSID = 0
    strSql = "SELECT PKID FROM ICMS WHERE ICMS = " & Formata_Dados(cboICMS.Text, tpDados_Moeda)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngICMSID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    '
    Set objFamiliaProduto = New busSisMetal.clsFamiliaProduto
    '
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      '
      objFamiliaProduto.AlterarFamiliaProduto lngFAMILIAPRODUTOSID, _
                                              txtDescricao.Text, _
                                              lngIPIID, _
                                              lngICMSID

      blnRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Pega Informações para inserir
      '
      objFamiliaProduto.InserirFamiliaProduto txtDescricao.Text, _
                                              lngIPIID, _
                                              lngICMSID
      '
      blnRetorno = True
    End If
    Set objFamiliaProduto = Nothing
    blnFechar = True
    Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  Dim strMsg     As String
  '
  If Len(txtDescricao.Text) = 0 Then
    strMsg = strMsg & "Informar a família de produtos" & vbCrLf
    Pintar_Controle txtDescricao, tpCorContr_Erro
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserFamiliaProdutoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    tabDetalhes.Tab = 0
    txtDescricao.SetFocus
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserFamiliaProdutoInc.Form_Activate]"
End Sub

Private Sub txtDescricao_GotFocus()
  Selecionar_Conteudo txtDescricao
End Sub

Private Sub txtDescricao_LostFocus()
  Pintar_Controle txtDescricao, tpCorContr_Normal
End Sub



Private Sub Form_Load()
On Error GoTo trata
  Dim objRs     As ADODB.Recordset
  Dim strSql    As String
  Dim objFamiliaProduto   As busSisMetal.clsFamiliaProduto
  '
  blnFechar = False
  blnRetorno = False
  AmpS
  Me.Height = 2940
  Me.Width = 10470
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  txtDescricao.Text = ""
  LimparCampoCombo cboIPI
  LimparCampoCombo cboICMS
  '
  'IPI
  strSql = "Select IPI from IPI ORDER BY IPI"
  PreencheCombo cboIPI, strSql, False, True
  'ICMS
  strSql = "Select ICMS from ICMS ORDER BY ICMS"
  PreencheCombo cboICMS, strSql, False, True
  '
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    '
  ElseIf Status = tpStatus_Alterar Then
    'Pega Dados do Banco de dados
    Set objFamiliaProduto = New busSisMetal.clsFamiliaProduto
    Set objRs = objFamiliaProduto.ListarFamiliaProduto(lngFAMILIAPRODUTOSID)
    '
    If Not objRs.EOF Then
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
      If objRs.Fields("IPI").Value & "" <> "" Then
        cboIPI.Text = objRs.Fields("IPI").Value & ""
      End If
      If objRs.Fields("ICMS").Value & "" <> "" Then
        cboICMS.Text = objRs.Fields("ICMS").Value & ""
      End If
      
      '
    End If
    Set objFamiliaProduto = Nothing
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
