VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmModeloInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de modelo"
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   30
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   330
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   4048
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do modelo"
      TabPicture(0)   =   "userModeloInc.frx":0000
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
         TabIndex        =   7
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
            TabIndex        =   8
            Top             =   0
            Width           =   7695
            Begin VB.ComboBox cboMarca 
               Height          =   315
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   0
               Top             =   240
               Width           =   3855
            End
            Begin VB.TextBox txtNome 
               Height          =   285
               Left            =   1200
               MaxLength       =   50
               TabIndex        =   1
               Text            =   "txtNome"
               Top             =   600
               Width           =   6375
            End
            Begin VB.Label Label11 
               Caption         =   "Marca"
               Height          =   255
               Left            =   240
               TabIndex        =   10
               Top             =   300
               Width           =   1095
            End
            Begin VB.Label Label6 
               Caption         =   "Nome"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   9
               Top             =   600
               Width           =   855
            End
         End
      End
   End
End
Attribute VB_Name = "frmModeloInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                     As tpStatus
Public lngMODELOID                As Long
Public blnRetorno                 As Boolean
Public blnFechar                  As Boolean
Private blnPrimeiraVez            As Boolean



Private Sub cboMarca_LostFocus()
  Pintar_Controle cboMarca, tpCorContr_Normal
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
  Dim objModelo               As busElite.clsModelo
  Dim objGeral                As busElite.clsGeral
  Dim lngMARCAID              As Long
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração da Unidade de estoque
    If Not ValidaCampos Then Exit Sub
    '
    Set objGeral = New busElite.clsGeral
    'MARCAID
    lngMARCAID = 0
    strSql = "SELECT PKID FROM MARCA WHERE NOME = " & Formata_Dados(cboMarca.Text, tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngMARCAID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    If lngMARCAID = 0 Then
      Set objGeral = Nothing
      TratarErroPrevisto "Selecionar uma marca", "cmdOK_Click"
      Pintar_Controle cboMarca, tpCorContr_Erro
      SetarFoco cboMarca
      Exit Sub
    End If
    'Valida se modelo já cadastrado
    strSql = "Select * From MODELO WHERE NOME = " & Formata_Dados(txtNome.Text, tpDados_Texto, tpNulo_NaoAceita) & _
      " AND MARCAID = " & Formata_Dados(lngMARCAID, tpDados_Longo) & _
      " AND PKID <> " & Formata_Dados(lngMODELOID, tpDados_Longo, tpNulo_NaoAceita)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      TratarErroPrevisto "Modelo já cadastrado", "cmdOK_Click"
      Pintar_Controle txtNome, tpCorContr_Erro
      SetarFoco txtNome
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objModelo = New busElite.clsModelo
    '
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      '
      objModelo.AlterarModelo lngMODELOID, _
                              txtNome.Text, _
                              lngMARCAID

      blnRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Pega Informações para inserir
      '
      objModelo.InserirModelo txtNome.Text, _
                              lngMARCAID
      '
      blnRetorno = True
    End If
    Set objModelo = Nothing
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
  If Len(txtNome.Text) = 0 Then
    strMsg = strMsg & "Informar o nome do modelo" & vbCrLf
    Pintar_Controle txtNome, tpCorContr_Erro
    SetarFoco txtNome
  End If
  If Len(cboMarca.Text) = 0 Then
    strMsg = strMsg & "Selecionar a marca" & vbCrLf
    Pintar_Controle cboMarca, tpCorContr_Erro
    SetarFoco cboMarca
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserModeloInc.ValidaCampos]"
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
    cboMarca.SetFocus
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserModeloInc.Form_Activate]"
End Sub

Private Sub txtNome_GotFocus()
  Selecionar_Conteudo txtNome
End Sub

Private Sub txtNome_LostFocus()
  Pintar_Controle txtNome, tpCorContr_Normal
End Sub



Private Sub Form_Load()
On Error GoTo trata
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  Dim objModelo     As busElite.clsModelo
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
  txtNome.Text = ""
  LimparCampoCombo cboMarca
  '
  'MARCA
  strSql = "Select NOME from MARCA ORDER BY NOME"
  PreencheCombo cboMarca, strSql, False, True
  '
  If Status = tpStatus_Incluir Then
    '
  ElseIf Status = tpStatus_Alterar Then
    'Pega Dados do Banco de dados
    Set objModelo = New busElite.clsModelo
    Set objRs = objModelo.ListarModelo(lngMODELOID)
    '
    If Not objRs.EOF Then
      txtNome.Text = objRs.Fields("NOME").Value & ""
      If objRs.Fields("NOME_MARCA").Value & "" <> "" Then
        cboMarca.Text = objRs.Fields("NOME_MARCA").Value & ""
      End If
      
      '
    End If
    Set objModelo = Nothing
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
