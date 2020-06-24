VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserDebitoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajuste de lançamento do Atendente"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   2880
      Left            =   8430
      ScaleHeight     =   2880
      ScaleWidth      =   1860
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2025
         Left            =   90
         ScaleHeight     =   1965
         ScaleWidth      =   1605
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   690
         Width           =   1665
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   960
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   2595
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   4577
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userDebitoInc.frx":0000
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
         Height          =   2175
         Left            =   120
         TabIndex        =   10
         Top             =   330
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   1935
            Index           =   0
            Left            =   120
            ScaleHeight     =   1935
            ScaleWidth      =   7575
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtNumero 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   5670
               TabIndex        =   2
               Text            =   "txtNumero"
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtBoleto 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1320
               TabIndex        =   1
               Text            =   "txtBoleto"
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtData 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1320
               TabIndex        =   3
               Text            =   "txtData"
               Top             =   810
               Width           =   1815
            End
            Begin VB.TextBox txtMaquina 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1320
               TabIndex        =   0
               Text            =   "txtMaquina"
               Top             =   150
               Width           =   6165
            End
            Begin MSMask.MaskEdBox mskValor 
               Height          =   255
               Left            =   1320
               TabIndex        =   4
               Top             =   1140
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskMedicao 
               Height          =   255
               Left            =   1320
               TabIndex        =   5
               Top             =   1440
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Medição"
               Height          =   285
               Index           =   5
               Left            =   60
               TabIndex        =   18
               Top             =   1410
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Número"
               Height          =   285
               Index           =   4
               Left            =   4410
               TabIndex        =   17
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Boleto"
               Height          =   285
               Index           =   3
               Left            =   60
               TabIndex        =   16
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Data"
               Height          =   285
               Index           =   2
               Left            =   60
               TabIndex        =   15
               Top             =   810
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Máquina"
               Height          =   285
               Index           =   0
               Left            =   60
               TabIndex        =   14
               Top             =   150
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Valor"
               Height          =   285
               Index           =   1
               Left            =   60
               TabIndex        =   13
               Top             =   1110
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserDebitoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long

Private blnPrimeiraVez          As Boolean

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Debito
  LimparCampoTexto txtMaquina
  LimparCampoTexto txtData
  LimparCampoTexto txtBoleto
  LimparCampoTexto txtNumero
  LimparCampoMask mskValor
  LimparCampoMask mskMedicao
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserDebitoInc.LimparCampos]", _
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
  Dim objDebito                As busSisMaq.clsDebito
  Dim objGeral                  As busSisMaq.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngTURNOID                As Long
  Dim lngCAIXAID                As Long
  Dim strStatus                 As String
  Dim strData                   As String
  Dim strMsg                    As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objDebito = New busSisMaq.clsDebito
  'DONO
  lngTURNOID = RetornaCodTurnoCorrente
  'Status Cadastroou T Turno
  strStatus = "C"
  '
  'Pedir senha de caixa
  lngCAIXAID = RetornaCaixaTurnoCorrente
  'Pede liberação do caixa
  frmUserLoginLibera.lngFUNCIONARIOID = lngCAIXAID
  frmUserLoginLibera.Show vbModal
  If Len(Trim(gsNomeUsuLib)) = 0 Then
    strMsg = "É necessário confirmação do caixa para executar esta ação."
    TratarErroPrevisto strMsg, "cmdConfirmar_Click"
    cmdOk.Enabled = True
    SetarFoco mskValor
    Exit Sub
  End If
  '
  If Status = tpStatus_Alterar Then
    'Alterar Debito
    objDebito.AlterarDebito lngPKID, _
                            mskMedicao.ClipText, _
                            mskValor.ClipText
    blnRetorno = True
    blnFechar = True
    Unload Me
    '
'''  ElseIf Status = tpStatus_Incluir Then
'''    'Inserir Debito
'''    strData = Format(Now, "DD/MM/YYYY hh:mm")
'''
'''    objDebito.InserirDebito lngTURNOID, _
'''                              mskValor.ClipText, _
'''                              strStatus, _
'''                              strData, _
'''                              giFuncionarioId
'''    blnRetorno = True
'''    blnFechar = True
'''    Unload Me
'''    '
  End If
  Set objDebito = Nothing
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
  If Not Valida_Moeda(mskValor, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o valor válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserDebitoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserDebitoInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco mskValor
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserDebitoInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objDebito             As busSisMaq.clsDebito
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 3360
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  'Limpar Campos
  LimparCampos
  If Status = tpStatus_Incluir Then
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objDebito = New busSisMaq.clsDebito
    Set objRs = objDebito.SelecionarDebitoPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      txtMaquina.Text = objRs.Fields("NUMERO_EQUIPAMENTO").Value & ""
      txtData.Text = Format(objRs.Fields("DATA").Value, "DD/MM/YYYY hh:mm")
      txtBoleto.Text = objRs.Fields("NUMERO_BOLETO").Value & ""
      txtNumero.Text = objRs.Fields("NUMERO").Value & ""
      '
      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VALORPAGO").Value, TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskMedicao, objRs.Fields("MEDICAO").Value, TpMaskMoeda
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objDebito = Nothing
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

Private Sub mskMedicao_GotFocus()
  Seleciona_Conteudo_Controle mskMedicao
End Sub
Private Sub mskMedicao_LostFocus()
  Pintar_Controle mskMedicao, tpCorContr_Normal
End Sub

Private Sub mskValor_GotFocus()
  Seleciona_Conteudo_Controle mskValor
End Sub
Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub
