VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAgenciaCNPJInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de CNPJ para Agência"
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
      TabPicture(0)   =   "userAgenciaCNPJInc.frx":0000
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
         Height          =   2655
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2385
            Index           =   0
            Left            =   120
            ScaleHeight     =   2385
            ScaleWidth      =   7575
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtAgenciaCNPJ 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtAgenciaCNPJ"
               Top             =   90
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskCnpj 
               Height          =   255
               Left            =   1320
               TabIndex        =   1
               Top             =   450
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   18
               Mask            =   "##.###.###/####-##"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Cnpj"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   10
               Top             =   450
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Agência"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   9
               Top             =   135
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmAgenciaCNPJInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long
Public lngAGENCIAID             As Long
Public strNomeAgencia           As String

Private blnPrimeiraVez          As Boolean

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Valor Plano Convênio
  LimparCampoTexto txtAgenciaCNPJ
  LimparCampoMask mskCnpj
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserAgenciaCNPJInc.LimparCampos]", _
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

Private Sub cmdOK_Click()
  Dim objAgenciaCNPJ            As busElite.clsAgenciaCNPJ
  Dim objGeral                  As busElite.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busElite.clsGeral
  Set objAgenciaCNPJ = New busElite.clsAgenciaCNPJ
  'Validar se plano convênio já cadastrado
  strSql = "SELECT * FROM AGENCIACNPJ " & _
    " WHERE AGENCIACNPJ.CNPJ = " & Formata_Dados(mskCnpj.ClipText, tpDados_Texto) & _
    " AND AGENCIACNPJ.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle mskCnpj, tpCorContr_Erro
    TratarErroPrevisto "CNPJ já cadastrado"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objAgenciaCNPJ = Nothing
    cmdOk.Enabled = True
    SetarFoco mskCnpj
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar AgenciaCNPJ
    objAgenciaCNPJ.AlterarAgenciaCNPJ lngPKID, _
                                      mskCnpj.ClipText
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir AgenciaCNPJ
    objAgenciaCNPJ.InserirAgenciaCNPJ lngAGENCIAID, _
                                      mskCnpj.ClipText
  End If
  Set objAgenciaCNPJ = Nothing
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
  If Not Valida_String(mskCnpj, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o cnpj" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(Trim(mskCnpj.ClipText)) > 0 Then
    If Len(Trim(mskCnpj.ClipText)) <> 14 Then
      strMsg = strMsg & "Informar o CNPJ válido" & vbCrLf
      Pintar_Controle mskCnpj, tpCorContr_Erro
      SetarFoco mskCnpj
      tabDetalhes.Tab = 0
      blnSetarFocoControle = False
    End If
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserAgenciaCNPJInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserAgenciaCNPJInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco mskCnpj
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserAgenciaCNPJInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objAgenciaCNPJ           As busElite.clsAgenciaCNPJ
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
  txtAgenciaCNPJ.Text = strNomeAgencia
  If Status = tpStatus_Incluir Then
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objAgenciaCNPJ = New busElite.clsAgenciaCNPJ
    Set objRs = objAgenciaCNPJ.SelecionarAgenciaCNPJPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      INCLUIR_VALOR_NO_MASK mskCnpj, objRs.Fields("CNPJ").Value, TpMaskSemMascara
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objAgenciaCNPJ = Nothing
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
Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then Cancel = True
End Sub

Private Sub mskCnpj_GotFocus()
  Seleciona_Conteudo_Controle mskCnpj
End Sub
Private Sub mskCnpj_LostFocus()
  Pintar_Controle mskCnpj, tpCorContr_Normal
End Sub

