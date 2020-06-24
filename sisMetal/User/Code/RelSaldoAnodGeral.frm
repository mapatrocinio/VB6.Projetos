VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRelSaldoAnodGeral 
   Caption         =   "Saldo de Perfis para Anodização"
   ClientHeight    =   3825
   ClientLeft      =   450
   ClientTop       =   1920
   ClientWidth     =   7605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3825
   ScaleWidth      =   7605
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   9
      Top             =   120
      Width           =   7395
      Begin VB.OptionButton optSai1 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optSai2 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   7605
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2850
      Width           =   7605
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   880
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdRelatorio 
         Caption         =   "ENTER"
         Default         =   -1  'True
         Height          =   880
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   7335
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "txtCodigo"
         Top             =   210
         Width           =   5865
      End
      Begin VB.TextBox txtLinhaFim 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3660
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "txtLinhaFim"
         Top             =   570
         Width           =   3495
      End
      Begin VB.TextBox txtCodigoFim 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "txtCodigoFim"
         Top             =   570
         Width           =   2355
      End
      Begin VB.Label Label1 
         Caption         =   "Nome da Linha/Código Perfil"
         Height          =   615
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3360
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmRelSaldoAnodGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lngLINHAID As Long


Private Sub cboEmpresa_LostFocus()
  Pintar_Controle cboEmpresa, tpCorContr_Normal
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdRelatorio_Click()
  On Error GoTo TratErro
  AmpS
  '
  If optSai1.Value Then
    Report1.Destination = 0 'Video
  ElseIf optSai2.Value Then
    Report1.Destination = 1   'Impressora
  End If
  Report1.CopiesToPrinter = 1
  Report1.WindowState = crptMaximized
  '
  Report1.Formulas(0) = "LINHAID = " & IIf(lngLINHAID = 0, "true = true", "{VW_CONS_ESTOQUE_PERFIL.LINHAID} = " & lngLINHAID) & ""
  'Report1.Formulas(3) = "MOTEL = '" & gsNomeEmpresa & "'"
  '
  Report1.Action = 1
  '
  AmpN
  Exit Sub
  
TratErro:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

Private Sub Form_Activate()
  SetarFoco txtCodigo
End Sub

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Ítens do pedido
  LimparCampoTexto txtCodigo
  LimparCampoTexto txtCodigoFim
  LimparCampoTexto txtLinhaFim
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmRelSaldoPerfil.LimparCampos]", _
            Err.Description
End Sub

Private Sub Form_Load()
  On Error GoTo RotErro
  AmpS
  Me.Height = 4335
  Me.Width = 7725
  CenterForm Me
  '
  LimparCampos
  lngLINHAID = 0
  LerFiguras Me, tpBmp_Vazio, , , cmdCancelar, , , , , , cmdRelatorio
  '
  Report1.Connect = ConnectRpt
  Report1.ReportFileName = gsReportPath & "EstoqueTotal.rpt"
  '
  AmpN
  Exit Sub
RotErro:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

Private Sub txtCodigo_GotFocus()
  Seleciona_Conteudo_Controle txtCodigo
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  KeyAscii = TRANSFORMA_MAIUSCULA(KeyAscii)
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtCodigo_LostFocus()
  On Error GoTo trata
  Dim objLinhaCons    As Form
  Dim objLinhaPerfil  As busSisMetal.clsLinhaPerfil
  Dim objRs           As ADODB.Recordset
  If Me.ActiveControl.Name = "cmdCancelar" Or Me.ActiveControl.Name = "cmdRelatorio" Then Exit Sub

  Pintar_Controle txtCodigo, tpCorContr_Normal
  If Len(txtCodigo.Text) = 0 Then
    If Len(txtCodigoFim.Text) <> 0 And Len(txtLinhaFim.Text) <> 0 Then
      txtCodigoFim.Text = ""
      txtLinhaFim.Text = ""
      lngLINHAID = 0
      SetarFoco cmdRelatorio
      Exit Sub
    Else
      'TratarErroPrevisto "Entre com o código ou descrição da linha."
      'Pintar_Controle txtCodigo, tpCorContr_Erro
      'SetarFoco txtCodigo
      txtCodigoFim.Text = ""
      txtLinhaFim.Text = ""
      lngLINHAID = 0
      SetarFoco cmdRelatorio
      Exit Sub
    End If
  End If
  Set objLinhaPerfil = New busSisMetal.clsLinhaPerfil
  '
  Set objRs = objLinhaPerfil.CapturaItemLinha(txtCodigo.Text)
  If objRs.EOF Then
    LimparCampoTexto txtCodigoFim
    LimparCampoTexto txtLinhaFim
    TratarErroPrevisto "Descrição/Código da linha não cadastrado"
    Pintar_Controle txtCodigo, tpCorContr_Erro
    SetarFoco txtCodigo
    Exit Sub
  Else
    If objRs.RecordCount = 1 Then
      txtCodigoFim.Text = objRs.Fields("CODIGO").Value & ""
      txtLinhaFim.Text = objRs.Fields("NOME").Value & ""
      lngLINHAID = objRs.Fields("PKID").Value
      SetarFoco cmdRelatorio
    Else
      'Novo : apresentar tela para seleção da linha
      Set objLinhaCons = New frmLinhaCons
      objLinhaCons.intIcOrigemLn = 5
      objLinhaCons.strCodigoDescricao = txtCodigo.Text
      objLinhaCons.Show vbModal
      SetarFoco cmdRelatorio
    End If
    '
    objRs.Close
    Set objRs = Nothing
    Set objLinhaPerfil = Nothing
'''    cmdOk.Default = True
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

