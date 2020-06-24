VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserProcPrestadorInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de procedimento para prestador"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   3435
      Left            =   8430
      ScaleHeight     =   3435
      ScaleWidth      =   1860
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2055
         Left            =   90
         ScaleHeight     =   1995
         ScaleWidth      =   1605
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1230
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   3165
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   5583
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userProcPrestadorInc.frx":0000
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
         TabIndex        =   12
         Top             =   390
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2385
            Index           =   0
            Left            =   120
            ScaleHeight     =   2385
            ScaleWidth      =   7575
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   180
            Width           =   7575
            Begin VB.ComboBox cboProcedimento 
               Height          =   315
               Left            =   1290
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   420
               Width           =   6105
            End
            Begin VB.TextBox txtPrestador 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtPrestador"
               Top             =   90
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskPercCasa 
               Height          =   255
               Left            =   1290
               TabIndex        =   2
               Top             =   780
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskPercPrest 
               Height          =   255
               Left            =   1290
               TabIndex        =   3
               Top             =   1080
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskPercRX 
               Height          =   255
               Left            =   1290
               TabIndex        =   4
               Top             =   1380
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskPercTecRX 
               Height          =   255
               Left            =   1290
               TabIndex        =   5
               Top             =   1680
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskPercDonoUltra 
               Height          =   255
               Left            =   1290
               TabIndex        =   6
               Top             =   1980
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "% Dono Ultra"
               Height          =   195
               Index           =   5
               Left            =   60
               TabIndex        =   20
               Top             =   1980
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "% Tec. RX"
               Height          =   195
               Index           =   4
               Left            =   60
               TabIndex        =   19
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "% Dono RX"
               Height          =   195
               Index           =   3
               Left            =   60
               TabIndex        =   18
               Top             =   1380
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "% Prest."
               Height          =   195
               Index           =   2
               Left            =   60
               TabIndex        =   17
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Prestador"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   16
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "% Casa"
               Height          =   195
               Index           =   21
               Left            =   60
               TabIndex        =   15
               Top             =   795
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Procedimento"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   14
               Top             =   450
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserProcPrestadorInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long
Public lngPRESTADORID           As Long
Public strNomePrestador         As String

Private blnPrimeiraVez          As Boolean

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Valor PrestProcedimento
  LimparCampoTexto txtPrestador
  LimparCampoCombo cboProcedimento
  LimparCampoMask mskPercCasa
  LimparCampoMask mskPercPrest
  LimparCampoMask mskPercRX
  LimparCampoMask mskPercTecRX
  LimparCampoMask mskPercDonoUltra
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserProcPrestadorInc.LimparCampos]", _
            Err.Description
End Sub

Private Sub cboProcedimento_LostFocus()
  Pintar_Controle cboProcedimento, tpCorContr_Normal
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
  Dim objPrestProcedimento      As busSisMed.clsPrestProcedimento
  Dim objGeral                  As busSisMed.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngPROCEDIMENTOID         As Long
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busSisMed.clsGeral
  Set objPrestProcedimento = New busSisMed.clsPrestProcedimento
  'PROCEDIMENTO
  lngPROCEDIMENTOID = 0
  strSql = "SELECT PKID FROM PROCEDIMENTO WHERE PROCEDIMENTO = " & Formata_Dados(cboProcedimento.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngPROCEDIMENTOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'Validar se Procedimento já esta associado ao prestador
  strSql = "SELECT * FROM PRESTADORPROCEDIMENTO " & _
    " WHERE PRESTADORPROCEDIMENTO.PROCEDIMENTOID = " & Formata_Dados(lngPROCEDIMENTOID, tpDados_Longo) & _
    " AND PRESTADORPROCEDIMENTO.PRONTUARIOID = " & Formata_Dados(lngPRESTADORID, tpDados_Longo) & _
    " AND PRESTADORPROCEDIMENTO.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle cboProcedimento, tpCorContr_Erro
    TratarErroPrevisto "Procedimento já associado ao prestador"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objPrestProcedimento = Nothing
    cmdOk.Enabled = True
    SetarFoco cboProcedimento
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar PrestProcedimento
    objPrestProcedimento.AlterarPrestProcedimento lngPKID, _
                                                  lngPROCEDIMENTOID, _
                                                  IIf(mskPercCasa.Text = "", "0", mskPercCasa.Text), _
                                                  IIf(mskPercPrest.Text = "", "0", mskPercPrest.Text), _
                                                  IIf(mskPercRX.Text = "", "0", mskPercRX.Text), _
                                                  IIf(mskPercTecRX.Text = "", "0", mskPercTecRX.Text), _
                                                  IIf(mskPercDonoUltra.Text = "", "0", mskPercDonoUltra.Text)
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir PrestProcedimento
    objPrestProcedimento.InserirPrestProcedimento lngPRESTADORID, _
                                                  lngPROCEDIMENTOID, _
                                                  IIf(mskPercCasa.Text = "", "0", mskPercCasa.Text), _
                                                  IIf(mskPercPrest.Text = "", "0", mskPercPrest.Text), _
                                                  IIf(mskPercRX.Text = "", "0", mskPercRX.Text), _
                                                  IIf(mskPercTecRX.Text = "", "0", mskPercTecRX.Text), _
                                                  IIf(mskPercDonoUltra.Text = "", "0", mskPercDonoUltra.Text)
  End If
  Set objPrestProcedimento = Nothing
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
  If Not Valida_String(cboProcedimento, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o procedimento" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskPercCasa, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o percentual da casa válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskPercPrest, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o percentual do prestador válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskPercRX, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o percentual do dono de RX válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskPercTecRX, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o percentual do técnico de RX válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskPercDonoUltra, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o percentual do dono de Ultrason válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserProcPrestadorInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserProcPrestadorInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco cboProcedimento
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserProcPrestadorInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objPrestProcedimento            As busSisMed.clsPrestProcedimento
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 3915
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  'Limpar Campos
  LimparCampos
  'Beneficiario Convênio
  strSql = "Select PROCEDIMENTO.PROCEDIMENTO from PROCEDIMENTO " & _
    "ORDER BY PROCEDIMENTO.PROCEDIMENTO"
  
  PreencheCombo cboProcedimento, strSql, False, True
  '
  txtPrestador.Text = strNomePrestador
  If Status = tpStatus_Incluir Then
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objPrestProcedimento = New busSisMed.clsPrestProcedimento
    Set objRs = objPrestProcedimento.SelecionarPrestProcedimentoPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      If objRs.Fields("PROCEDIMENTO").Value & "" <> "" Then
        cboProcedimento.Text = objRs.Fields("PROCEDIMENTO").Value & ""
      End If
      INCLUIR_VALOR_NO_MASK mskPercCasa, objRs.Fields("PERCCASA").Value & "", TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskPercPrest, objRs.Fields("PERCPRESTADOR").Value & "", TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskPercRX, objRs.Fields("PERCRX").Value & "", TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskPercTecRX, objRs.Fields("PERCTECRX").Value & "", TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskPercDonoUltra, objRs.Fields("PERCULTRA").Value & "", TpMaskMoeda
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objPrestProcedimento = Nothing
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

Private Sub mskPercCasa_GotFocus()
  Seleciona_Conteudo_Controle mskPercCasa
End Sub
Private Sub mskPercCasa_LostFocus()
  Pintar_Controle mskPercCasa, tpCorContr_Normal
End Sub

Private Sub mskPercDonoUltra_GotFocus()
  Seleciona_Conteudo_Controle mskPercDonoUltra
End Sub
Private Sub mskPercDonoUltra_LostFocus()
  Pintar_Controle mskPercDonoUltra, tpCorContr_Normal
End Sub

Private Sub mskPercPrest_GotFocus()
  Seleciona_Conteudo_Controle mskPercPrest
End Sub
Private Sub mskPercPrest_LostFocus()
  Pintar_Controle mskPercPrest, tpCorContr_Normal
End Sub

Private Sub mskPercRX_GotFocus()
  Seleciona_Conteudo_Controle mskPercRX
End Sub
Private Sub mskPercRX_LostFocus()
  Pintar_Controle mskPercRX, tpCorContr_Normal
End Sub

Private Sub mskPercTecRX_GotFocus()
  Seleciona_Conteudo_Controle mskPercTecRX
End Sub
Private Sub mskPercTecRX_LostFocus()
  Pintar_Controle mskPercTecRX, tpCorContr_Normal
End Sub
