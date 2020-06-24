VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserParcelaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Parcela"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   3420
      Left            =   6510
      ScaleHeight     =   3420
      ScaleWidth      =   1860
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   30
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   150
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   3135
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da Parcela"
      TabPicture(0)   =   "userParcelaInc.frx":0000
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
         Height          =   2625
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   6015
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2295
            Index           =   2
            Left            =   120
            ScaleHeight     =   2295
            ScaleWidth      =   5775
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   120
            Width           =   5775
            Begin VB.TextBox txtParcela 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   1
               TabStop         =   0   'False
               Text            =   "txtParcela"
               Top             =   420
               Width           =   525
            End
            Begin VB.TextBox txtUnidade 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtUnidade"
               Top             =   90
               Width           =   4275
            End
            Begin MSMask.MaskEdBox mskDtVencimento 
               Height          =   255
               Left            =   1320
               TabIndex        =   2
               Top             =   780
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   16777215
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskVrParcela 
               Height          =   255
               Left            =   1320
               TabIndex        =   3
               Top             =   1080
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Valor Parcela"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   14
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Parcela"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   13
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Dt. Vencimento"
               Height          =   195
               Index           =   3
               Left            =   60
               TabIndex        =   12
               Top             =   735
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Unidade"
               Height          =   195
               Index           =   24
               Left            =   60
               TabIndex        =   11
               Top             =   60
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserParcelaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean
Public strNumeroAptoPrinc       As String

Public lngParcelaId             As Long
Private blnPrimeiraVez          As Boolean 'Propósito: Preencher lista no combo



Private Sub TratarCampos()
  On Error GoTo trata
  'Configurações iniciais
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserParcelaInc.TratarCampos]", _
            Err.Description
End Sub

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Cliente
  LimparCampoTexto txtUnidade
  LimparCampoTexto txtParcela
  LimparCampoMask mskDtVencimento
  LimparCampoMask mskVrParcela
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserParcelaInc.LimparCampos]", _
            Err.Description
End Sub

Private Sub cmdCancelar_Click()
  Dim objCartPromo As busApler.clsCartaoPromocional
  blnFechar = True
  blnRetorno = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Dim objParcela                As busApler.clsParcela
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaParcela Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objParcela = New busApler.clsParcela
  '
  If Status = tpStatus_Alterar Then
    'Alterar Parcela
    objParcela.AlterarParcela lngParcelaId, _
                              mskDtVencimento.Text, _
                              "", _
                              mskVrParcela.Text, _
                              ""
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir Parcela
  End If
  Set objParcela = Nothing
  blnRetorno = True
  blnFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  cmdOk.Enabled = True
End Sub

Private Function ValidaParcela() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaParcela = False
  If Not Valida_Data(mskDtVencimento, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de vencimento válida" & vbCrLf
  End If
  If Not Valida_Moeda(mskVrParcela, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o valor da parcela válido" & vbCrLf
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserParcelaInc.ValidaParcela]"
    ValidaParcela = False
  Else
    ValidaParcela = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserParcelaInc.ValidaParcela]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  
  If blnPrimeiraVez Then
    SetarFoco mskDtVencimento
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserParcelaInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objParcela              As busApler.clsParcela
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 3795
  Me.Width = 8460
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  TratarCampos
  'Limpar Campos
  LimparCampos
  'tabDetalhes_Click 0
  txtUnidade.Text = strNumeroAptoPrinc
  If Status = tpStatus_Incluir Then
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objParcela = New busApler.clsParcela
    Set objRs = objParcela.SelecionarParcela(lngParcelaId)
    '
    If Not objRs.EOF Then
      txtParcela.Text = objRs.Fields("PARCELA").Value & ""
      INCLUIR_VALOR_NO_MASK mskDtVencimento, objRs.Fields("DTVENCIMENTO").Value & "", TpMaskData
      INCLUIR_VALOR_NO_MASK mskVrParcela, objRs.Fields("VRPARCELA").Value & "", TpMaskMoeda
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objParcela = Nothing
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

Private Sub mskDtVencimento_GotFocus()
  Seleciona_Conteudo_Controle mskDtVencimento
End Sub
Private Sub mskDtVencimento_LostFocus()
  Pintar_Controle mskDtVencimento, tpCorContr_Normal
End Sub

Private Sub mskVrParcela_GotFocus()
  Seleciona_Conteudo_Controle mskVrParcela
End Sub
Private Sub mskVrParcela_LostFocus()
  Pintar_Controle mskVrParcela, tpCorContr_Normal
End Sub
