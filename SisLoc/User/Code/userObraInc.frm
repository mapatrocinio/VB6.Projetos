VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUserObraInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de obra da empresa/contrato"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4125
      Left            =   8430
      ScaleHeight     =   4125
      ScaleWidth      =   1860
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   90
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1830
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   150
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   3795
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   6694
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userObraInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2925
            Index           =   0
            Left            =   120
            ScaleHeight     =   2925
            ScaleWidth      =   7575
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtContrato 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   12
               TabStop         =   0   'False
               Text            =   "txtContrato"
               Top             =   420
               Width           =   6075
            End
            Begin VB.TextBox txtDescricao 
               BackColor       =   &H00FFFFFF&
               Height          =   585
               Left            =   1320
               MaxLength       =   255
               MultiLine       =   -1  'True
               TabIndex        =   1
               Text            =   "userObraInc.frx":001C
               Top             =   780
               Width           =   6075
            End
            Begin VB.TextBox txtEmpresa 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtEmpresa"
               Top             =   90
               Width           =   6075
            End
            Begin VB.Label Label5 
               Caption         =   "Descrição"
               Height          =   195
               Index           =   5
               Left            =   60
               TabIndex        =   11
               Top             =   765
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Contrato"
               Height          =   195
               Index           =   24
               Left            =   60
               TabIndex        =   10
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Empresa"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   9
               Top             =   105
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserObraInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long
Public lngCONTRATOID            As Long
Public strDescrEmpresa          As String
Public strDescrContrato         As String

Private blnPrimeiraVez          As Boolean

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Contrato
  LimparCampoTexto txtEmpresa
  LimparCampoTexto txtContrato
  LimparCampoTexto txtDescricao
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserObraInc.LimparCampos]", _
            Err.Description
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

Private Sub cmdOk_Click()
  Dim objObra                   As busSisLoc.clsObra
  Dim objGeral                  As busSisLoc.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busSisLoc.clsGeral
  Set objObra = New busSisLoc.clsObra
  'Validar se Obra já cadastrada
  strSql = "SELECT * FROM OBRA " & _
    " WHERE OBRA.DESCRICAO = " & Formata_Dados(txtDescricao.Text, tpDados_Texto) & _
    " AND OBRA.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle txtDescricao, tpCorContr_Erro
    TratarErroPrevisto "Obra já cadastrada"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objObra = Nothing
    cmdOk.Enabled = True
    SetarFoco txtDescricao
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar Obra
    objObra.AlterarObra lngPKID, _
                        txtDescricao.Text
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir Obra
    objObra.InserirObra lngCONTRATOID & "", _
                        txtDescricao.Text
  End If
  Set objObra = Nothing
  blnRetorno = True
  blnFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  cmdOk.Enabled = True
End Sub


Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCampos = False
  If Not Valida_String(txtDescricao, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a descrição da obra válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserObraInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserObraInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco txtDescricao
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserObraInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objObra           As busSisLoc.clsObra
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 4605
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  'Limpar Campos
  LimparCampos
  '
  txtEmpresa.Text = strDescrEmpresa
  txtContrato.Text = strDescrContrato
  '
  If Status = tpStatus_Incluir Then
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objObra = New busSisLoc.clsObra
    Set objRs = objObra.SelecionarObraPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      '
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objObra = Nothing
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

Private Sub txtDescricao_GotFocus()
  Seleciona_Conteudo_Controle txtDescricao
End Sub
Private Sub txtDescricao_LostFocus()
  Pintar_Controle txtDescricao, tpCorContr_Normal
End Sub
