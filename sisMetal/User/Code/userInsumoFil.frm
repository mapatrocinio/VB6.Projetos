VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInsumoFil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtrar"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   8250
      ScaleHeight     =   2985
      ScaleWidth      =   1860
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   0
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   5
         Top             =   690
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
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   4683
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do filtro"
      TabPicture(0)   =   "userInsumoFil.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
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
         Height          =   1935
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   7575
         Begin VB.ComboBox cboCor 
            Height          =   315
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   930
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Left            =   1380
            MaxLength       =   50
            TabIndex        =   0
            Text            =   "txtCodigo"
            Top             =   420
            Width           =   5865
         End
         Begin VB.Label Label1 
            Caption         =   "Nome da Linha/Código Perfil"
            Height          =   615
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Cor"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmInsumoFil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnRetorno                     As Boolean
Public blnPrimeiraVez                 As Boolean
Public blnFechar                      As Boolean
Public intTipoInsumo                  As tpInsumo
Public lngCORID                       As Long
Public strNOME                        As String

Private Sub cmdCancelar_Click()
  On Error GoTo trata
  '
  strNOME = txtCodigo.Text
  lngCORID = 0
  'frmInsumoLis.strNOMEFIL = ""
  'frmInsumoLis.lngCORIDFIL = 0
  'frmInsumoLis.MontaMatriz frmInsumoLis.strNOMEFIL, frmInsumoLis.lngCORIDFIL
  'frmInsumoLis.grdGeral.Bookmark = Null
  'frmInsumoLis.grdGeral.ReBind
  '
  blnFechar = True
  blnRetorno = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub


Private Sub cmdOK_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim objRs                   As ADODB.Recordset
  Dim objGer                  As busSisMetal.clsGeral
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração da Despesa
    If Not ValidaCampos Then Exit Sub
    '
    Select Case intTipoInsumo
    Case tpInsumo_Perfil
      '
      'COR
      lngCORID = 0
      If cboCor.Text <> "" Then
        Set objGer = New busSisMetal.clsGeral
        strSql = "SELECT PKID FROM COR WHERE NOME = " & Formata_Dados(cboCor.Text, tpDados_Texto)
        Set objRs = objGer.ExecutarSQL(strSql)
        If Not objRs.EOF Then
          lngCORID = objRs.Fields("PKID").Value
        End If
        objRs.Close
        Set objRs = Nothing
        '
        Set objGer = Nothing
      End If
      strNOME = txtCodigo.Text
      'frmInsumoLis.strNOMEFIL = txtCodigo.Text
      'frmInsumoLis.lngCORIDFIL = lngCORID
      'frmInsumoLis.MontaMatriz frmInsumoLis.strNOMEFIL, frmInsumoLis.lngCORIDFIL
      'frmInsumoLis.grdGeral.Bookmark = Null
      'frmInsumoLis.grdGeral.ReBind
    Case tpInsumo_Acessorio
      strNOME = txtCodigo.Text
      'frmInsumoLis.strNOMEFIL = txtCodigo.Text
      lngCORID = 0
      'frmInsumoLis.MontaMatriz frmInsumoLis.strNOMEFIL, frmInsumoLis.lngCORIDFIL
      'frmInsumoLis.grdGeral.Bookmark = Null
      'frmInsumoLis.grdGeral.ReBind
    End Select
    '
    blnFechar = True
    blnRetorno = True
    Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  Dim strMsg              As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  '
  Select Case intTipoInsumo
  Case tpInsumo_Perfil
    If txtCodigo.Text = "" And cboCor.Text = "" Then
      strMsg = strMsg & "Informar o perfil ou selecionar a cor" & vbCrLf
      SetarFoco txtCodigo
      'Pintar_Controle txtNome, tpCorContr_Erro
    End If
  Case tpInsumo_Acessorio
    If txtCodigo.Text = "" Then
      strMsg = strMsg & "Informar o nome do acessório" & vbCrLf
      SetarFoco txtCodigo
      'Pintar_Controle txtNome, tpCorContr_Erro
    End If
  End Select
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmInsumoFil.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Setar foco
    SetarFoco txtCodigo
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmInsumoFil.Form_Activate]"
End Sub

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Insumo
  LimparCampoTexto txtCodigo
  LimparCampoCombo cboCor
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmInsumoFil.LimparCampos]", _
            Err.Description
End Sub


Private Sub Form_Load()
On Error GoTo trata
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  '
  blnFechar = False
  blnRetorno = False
  AmpS
  Me.Height = 3360
  Me.Width = 10200
  CenterForm Me
  blnPrimeiraVez = True
  '
  Select Case intTipoInsumo
  Case tpInsumo_Perfil
    Me.Caption = Me.Caption & " Perfil"
    Label1(0).Caption = "Nome da Linha/Código Perfil"
    Label1(1).Visible = True
    cboCor.Visible = True
  Case tpInsumo_Acessorio
    Me.Caption = Me.Caption & " Acessório"
    Label1(0).Caption = "Nome do Acessório"
    Label1(1).Visible = False
    cboCor.Visible = False
  End Select
  'Limpar Campos
  LimparCampos
  'Cor
  strSql = "Select NOME from COR ORDER BY NOME"
  PreencheCombo cboCor, strSql, False, True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  AmpN
  Exit Sub
trata:
  AmpN
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then Cancel = True
End Sub

Private Sub txtCodigo_GotFocus()
  Selecionar_Conteudo txtCodigo
End Sub

Private Sub txtCodigo_LostFocus()
  Pintar_Controle txtCodigo, tpCorContr_Normal
End Sub

Private Sub cboCor_LostFocus()
  Pintar_Controle cboCor, tpCorContr_Normal
End Sub

