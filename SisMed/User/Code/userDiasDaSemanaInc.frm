VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUserDiasDaSemanaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de dia da semana"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   3390
      Left            =   8520
      ScaleHeight     =   3390
      ScaleWidth      =   1860
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2055
         Left            =   60
         ScaleHeight     =   1995
         ScaleWidth      =   1605
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1140
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
      Height          =   3105
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   5477
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do dia da semana"
      TabPicture(0)   =   "userDiasDaSemanaInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picTrava(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.PictureBox picTrava 
         BorderStyle     =   0  'None
         Height          =   2445
         Index           =   0
         Left            =   120
         ScaleHeight     =   2445
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
            Height          =   2385
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   7695
            Begin VB.TextBox txtDiaDaSemana 
               Height          =   285
               Left            =   1560
               MaxLength       =   50
               TabIndex        =   1
               Top             =   630
               Width           =   5655
            End
            Begin VB.TextBox txtCodigo 
               Height          =   285
               Left            =   1560
               MaxLength       =   2
               TabIndex        =   0
               Top             =   270
               Width           =   495
            End
            Begin VB.Label Label6 
               Caption         =   "Dia da Semana"
               Height          =   255
               Index           =   2
               Left            =   270
               TabIndex        =   10
               Top             =   630
               Width           =   1455
            End
            Begin VB.Label Label6 
               Caption         =   "Código"
               Height          =   255
               Index           =   0
               Left            =   270
               TabIndex        =   9
               Top             =   270
               Width           =   1455
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserDiasDaSemanaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                     As tpStatus
Public lngDIADASEMANAID           As Long
Public bRetorno                   As Boolean
Public bFechar                    As Boolean
Private blnPrimeiraVez            As Boolean



Private Sub cmdCancelar_Click()
  bFechar = True
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
  Dim objDiasDaSemana         As busSisMed.clsDiasDaSemana
  Dim objGer                  As busSisMed.clsGeral
  
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração de Grupo cardápio
    If Not ValidaCampos Then Exit Sub
    'Valida se dia da semana já cadastrado
    Set objGer = New busSisMed.clsGeral
    strSql = "Select * From DIASDASEMANA WHERE (DIADASEMANA = " & Formata_Dados(txtDiaDaSemana.Text, tpDados_Texto, tpNulo_Aceita) & _
      " OR CODIGO = " & Formata_Dados(txtCodigo.Text, tpDados_Longo, tpNulo_Aceita) & ") " & _
      " AND PKID <> " & Formata_Dados(lngDIADASEMANAID, tpDados_Longo, tpNulo_NaoAceita)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGer = Nothing
      TratarErroPrevisto "Dia da Semana já cadastrado", "cmdOK_Click"
      Pintar_Controle txtCodigo, tpCorContr_Erro
      SetarFoco txtCodigo
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGer = Nothing
    '
    Set objDiasDaSemana = New busSisMed.clsDiasDaSemana
    '
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objDiasDaSemana.AlterarDiaDaSemana lngDIADASEMANAID, _
                                         txtCodigo.Text, _
                                         txtDiaDaSemana.Text
                            
      bRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Pega Informações para inserir
      '
      objDiasDaSemana.InserirDiaDaSemana txtCodigo.Text, _
                                         txtDiaDaSemana.Text
      '
      bRetorno = True
    End If
    Set objDiasDaSemana = Nothing
    bFechar = True
    Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg        As String
  Dim blnSetarFoco  As Boolean
  '
  blnSetarFoco = True
  If Not IsNumeric(txtCodigo.Text) Then
    strMsg = strMsg & "Informar o código válido" & vbCrLf
    Pintar_Controle txtCodigo, tpCorContr_Erro
    blnSetarFoco = False
    SetarFoco txtCodigo
  End If
  If Not Valida_String(txtDiaDaSemana, TpObrigatorio, blnSetarFoco) Then
    strMsg = strMsg & "Preencher o Dia da Semana" & vbCrLf
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserDiasDaSemanaInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  ValidaCampos = False
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    tabDetalhes.Tab = 0
    SetarFoco txtCodigo
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserDiasDaSemanaInc.Form_Activate]"
End Sub

Private Sub txtCodigo_GotFocus()
  Selecionar_Conteudo txtCodigo
End Sub

Private Sub txtCodigo_LostFocus()
  Pintar_Controle txtCodigo, tpCorContr_Normal
End Sub

Private Sub txtDiaDaSemana_GotFocus()
  Selecionar_Conteudo txtDiaDaSemana
End Sub

Private Sub txtDiaDaSemana_LostFocus()
  Pintar_Controle txtDiaDaSemana, tpCorContr_Normal
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs             As ADODB.Recordset
  Dim strSql            As String
  Dim objDiasDaSemana   As busSisMed.clsDiasDaSemana
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 3765
  Me.Width = 10470
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    txtCodigo.Text = ""
    txtDiaDaSemana.Text = ""
    'INCLUIR_VALOR_NO_MASK mskHora(0), "", TpMaskOutros
    'INCLUIR_VALOR_NO_MASK mskHora(1), "", TpMaskOutros
    '
    txtCodigo.Enabled = True
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objDiasDaSemana = New busSisMed.clsDiasDaSemana
    Set objRs = objDiasDaSemana.ListarDiaDaSemana(lngDIADASEMANAID)
    '
    If Not objRs.EOF Then
      txtCodigo.Text = objRs.Fields("CODIGO").Value & ""
      txtDiaDaSemana.Text = objRs.Fields("DIADASEMANA").Value & ""
      '
    End If
    Set objDiasDaSemana = Nothing
    txtCodigo.Enabled = False
  End If
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub



Private Sub Form_Unload(Cancel As Integer)
  If Not bFechar Then Cancel = True
End Sub
