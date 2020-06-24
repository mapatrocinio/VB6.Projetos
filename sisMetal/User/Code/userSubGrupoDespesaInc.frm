VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSubGrupoDespesaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sub Grupo de Despesa"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4185
      Left            =   8250
      ScaleHeight     =   4185
      ScaleWidth      =   1860
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   1875
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   1605
         TabIndex        =   5
         Top             =   2160
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Default         =   -1  'True
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   3855
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do Sub Grupo"
      TabPicture(0)   =   "userSubGrupoDespesaInc.frx":0000
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
         Height          =   3135
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   7335
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
            TabIndex        =   8
            Top             =   3480
            Width           =   2295
         End
         Begin VB.TextBox txtDescricao 
            Height          =   285
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   1
            Top             =   720
            Width           =   5175
         End
         Begin MSMask.MaskEdBox mskSubGrupo 
            Height          =   255
            Left            =   1560
            TabIndex        =   0
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblCheque 
            Caption         =   "Descrição"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Código"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmSubGrupoDespesaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                         As tpStatus
Public lngGRUPODESPESAID              As Long
Public lngSUBGRUPODESPESAID           As Long
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
  Dim strMsgErro              As String
  Dim objRs                   As ADODB.Recordset
  Dim objSubGrupoDespesa      As busSisMetal.clsSubGrupoDespesa
  Dim objGeral                As busSisMetal.clsGeral
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração
    If Not ValidaCampos Then Exit Sub
    'Valida se grupo da despesa já cadastrada
    Set objGeral = New busSisMetal.clsGeral
    strSql = "Select PKID From SUBGRUPODESPESA WHERE CODIGO = " & Formata_Dados(mskSubGrupo.Text, tpDados_Texto, tpNulo_NaoAceita) & _
      " AND PKID <> " & lngSUBGRUPODESPESAID & _
      " AND GRUPODESPESAID = " & lngGRUPODESPESAID
      
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      TratarErroPrevisto "Código do sub grupo de despesa já cadastrado", "cmdOK_Click"
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    
    strSql = "Select PKID From SUBGRUPODESPESA WHERE DESCRICAO = " & Formata_Dados(txtDescricao.Text, tpDados_Texto, tpNulo_NaoAceita) & _
      " AND PKID <> " & lngSUBGRUPODESPESAID & _
      " AND GRUPODESPESAID = " & lngGRUPODESPESAID
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      TratarErroPrevisto "Descrição do sub grupo de despesa já cadastrada", "cmdOK_Click"
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing

    Set objSubGrupoDespesa = New busSisMetal.clsSubGrupoDespesa
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objSubGrupoDespesa.AlterarSubGrupoDespesa lngSUBGRUPODESPESAID, _
                                                mskSubGrupo.Text, _
                                                txtDescricao.Text
      Set objSubGrupoDespesa = Nothing
      bRetorno = True
      bFechar = True
      Unload Me
    ElseIf Status = tpStatus_Incluir Then
      'Código para inclusão
      '
      objSubGrupoDespesa.IncluirSubGrupoDespesa lngGRUPODESPESAID, _
                                                mskSubGrupo.Text, _
                                                txtDescricao.Text
      INCLUIR_VALOR_NO_MASK mskSubGrupo, "__", TpMaskOutros
      txtDescricao.Text = ""
      bRetorno = True
      SetarFoco mskSubGrupo
    End If
    Set objSubGrupoDespesa = Nothing
    'bFechar = True
    'Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  Dim strMsg     As String
  '
  If Not IsNumeric(mskSubGrupo.Text) Or Len(mskSubGrupo.ClipText) <> 2 Then
    strMsg = strMsg & "Informar o código do sub grupo da despesa válido" & vbCrLf
    Pintar_Controle mskSubGrupo, tpCorContr_Erro
  End If
  If Len(txtDescricao.Text) = 0 Then
    strMsg = strMsg & "Informar a descrição válida" & vbCrLf
    Pintar_Controle txtDescricao, tpCorContr_Erro
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserSubGrupoDespesaInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco mskSubGrupo
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserSubGrupoDespesaInc.Form_Activate]"
End Sub



Private Sub Form_Load()
On Error GoTo trata
  Dim objRs               As ADODB.Recordset
  Dim strSql              As String
  Dim objSubGrupoDespesa  As busSisMetal.clsSubGrupoDespesa
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 4560
  Me.Width = 10200
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    LimparCampoMask mskSubGrupo
    LimparCampoTexto txtDescricao
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objSubGrupoDespesa = New busSisMetal.clsSubGrupoDespesa
    Set objRs = objSubGrupoDespesa.SelecionarSubGrupoDespesa(lngSUBGRUPODESPESAID)
    '
    If Not objRs.EOF Then
      INCLUIR_VALOR_NO_MASK mskSubGrupo, objRs.Fields("CODIGO").Value & "", TpMaskOutros
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
    End If
    Set objSubGrupoDespesa = Nothing
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

Private Sub mskSubGrupo_GotFocus()
  Selecionar_Conteudo mskSubGrupo
End Sub

Private Sub mskSubGrupo_LostFocus()
  Pintar_Controle mskSubGrupo, tpCorContr_Normal
End Sub

Private Sub txtDescricao_GotFocus()
  Selecionar_Conteudo txtDescricao
End Sub

Private Sub txtDescricao_LostFocus()
  Pintar_Controle txtDescricao, tpCorContr_Normal
End Sub

