VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUserFormaPgtoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forma de pagamento"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   3045
      Left            =   8250
      ScaleHeight     =   3045
      ScaleWidth      =   1860
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2115
         Left            =   120
         ScaleHeight     =   2055
         ScaleWidth      =   1605
         TabIndex        =   4
         Top             =   780
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1020
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   4895
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da forma de pagamento"
      TabPicture(0)   =   "userFormaPgtoInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Informações cadastrais"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   7335
         Begin VB.TextBox txtFormaPgto 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   0
            Top             =   360
            Width           =   5055
         End
         Begin VB.Frame Frame5 
            Caption         =   "Sexo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   840
            TabIndex        =   7
            Top             =   3480
            Width           =   2295
         End
         Begin VB.Label Label9 
            Caption         =   "Forma de pagamento"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "frmUserFormaPgtoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                         As tpStatus
Public lngFORMAPGTOID                 As Long
Public bRetorno                       As Boolean
Public blnPrimeiraVez                 As Boolean
Public bFechar                        As Boolean


Private Sub cmdCancelar_Click()
  bFechar = True
  '
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
  Dim objFormaPgto            As busApler.clsFormaPgto
  Dim clsGer                  As busApler.clsGeral
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração da Despesa
    If Not ValidaCampos Then Exit Sub
    'Valida se unidade de estoque já cadastrada
    Set clsGer = New busApler.clsGeral
    strSql = "Select * From FORMAPGTO WHERE FORMAPGTO = " & Formata_Dados(txtFormaPgto.Text, tpDados_Texto, tpNulo_NaoAceita) & _
      " AND PKID <> " & Formata_Dados(lngFORMAPGTOID, tpDados_Longo, tpNulo_Aceita)
    Set objRs = clsGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set clsGer = Nothing
      TratarErroPrevisto "Forma de Pagamento já cadastrada", "cmdOK_Click"
      SetarFoco txtFormaPgto
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing

    Set objFormaPgto = New busApler.clsFormaPgto

    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objFormaPgto.AlterarFormaPgto lngFORMAPGTOID, _
                                    txtFormaPgto.Text
      bRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Código para inclusão
      '
      objFormaPgto.IncluirFormaPgto txtFormaPgto
      bRetorno = True
    End If
    Set objFormaPgto = Nothing
    bFechar = True
    Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  Dim strMsg     As String
  '
  If Len(txtFormaPgto.Text) = 0 Then
    strMsg = strMsg & "Informar a forma de pagamento válida" & vbCrLf
    Pintar_Controle txtFormaPgto, tpCorContr_Erro
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserFormaPgtoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    SetarFoco txtFormaPgto
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserFormaPgtoInc.Form_Activate]"
End Sub



Private Sub Form_Load()
On Error GoTo trata
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim objFormaPgto    As busApler.clsFormaPgto
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 3420
  Me.Width = 10200
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão
    LimparCampoTexto txtFormaPgto
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objFormaPgto = New busApler.clsFormaPgto
    Set objRs = objFormaPgto.SelecionarFormaPgto(lngFORMAPGTOID)
    '
    If Not objRs.EOF Then
      txtFormaPgto.Text = objRs.Fields("FORMAPGTO").Value & ""
    End If
    Set objFormaPgto = Nothing
  End If
  '
  AmpN
  Exit Sub
trata:
  AmpN
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not bFechar Then Cancel = True
End Sub

Private Sub txtFormaPgto_GotFocus()
  Selecionar_Conteudo txtFormaPgto
End Sub

Private Sub txtFormaPgto_LostFocus()
  Pintar_Controle txtFormaPgto, tpCorContr_Normal
End Sub


