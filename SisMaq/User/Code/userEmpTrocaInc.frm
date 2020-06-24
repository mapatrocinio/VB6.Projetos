VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserEmpTrocaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Empréstimo/troca"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   3390
      Left            =   8430
      ScaleHeight     =   3390
      ScaleWidth      =   1860
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2025
         Left            =   90
         ScaleHeight     =   1965
         ScaleWidth      =   1605
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1665
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   960
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   3105
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   5477
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userEmpTrocaInc.frx":0000
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
         Height          =   2685
         Left            =   120
         TabIndex        =   9
         Top             =   330
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2415
            Index           =   0
            Left            =   120
            ScaleHeight     =   2415
            ScaleWidth      =   7575
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtNome 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   2
               Top             =   810
               Width           =   6165
            End
            Begin VB.ComboBox cboTipoPgto 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   1140
               Width           =   6195
            End
            Begin VB.TextBox txtData 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1320
               TabIndex        =   1
               Text            =   "txtData"
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtTurno 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1320
               TabIndex        =   0
               Text            =   "txtTurno"
               Top             =   150
               Width           =   6165
            End
            Begin MSMask.MaskEdBox mskValor 
               Height          =   255
               Left            =   1320
               TabIndex        =   4
               Top             =   1500
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label6 
               Caption         =   "Nome"
               Height          =   255
               Index           =   0
               Left            =   90
               TabIndex        =   16
               Top             =   810
               Width           =   1455
            End
            Begin VB.Label Label6 
               Caption         =   "Tipo Pgto."
               Height          =   255
               Index           =   1
               Left            =   60
               TabIndex        =   15
               Top             =   1140
               Width           =   1455
            End
            Begin VB.Label Label5 
               Caption         =   "Data"
               Height          =   285
               Index           =   2
               Left            =   60
               TabIndex        =   14
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Turno"
               Height          =   285
               Index           =   0
               Left            =   60
               TabIndex        =   13
               Top             =   150
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Valor"
               Height          =   285
               Index           =   1
               Left            =   60
               TabIndex        =   12
               Top             =   1470
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserEmpTrocaInc"
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
  'EmpTroca
  LimparCampoTexto txtTurno
  LimparCampoTexto txtData
  LimparCampoTexto txtNome
  LimparCampoCombo cboTipoPgto
  LimparCampoMask mskValor
  
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserEmpTrocaInc.LimparCampos]", _
            Err.Description
End Sub



Private Sub cboTipoPgto_LostFocus()
  Pintar_Controle cboTipoPgto, tpCorContr_Normal
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
  Dim objEmpTroca               As busSisMaq.clsEmpTroca
  Dim objGeral                  As busSisMaq.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngTURNOID                As Long
  Dim lngTIPOPGTOID             As Long
  Dim strData                   As String
  Dim strMsg                    As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objEmpTroca = New busSisMaq.clsEmpTroca
  'DONO
  lngTURNOID = RetornaCodTurnoCorrente
  Set objGeral = New busSisMaq.clsGeral
  'TIPOPGTO
  lngTIPOPGTOID = 0
  strSql = "SELECT TIPOPGTO.PKID FROM TIPOPGTO WHERE TIPOPGTO.TIPOPGTO = " & Formata_Dados(cboTipoPgto.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngTIPOPGTOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  '
  If Status = tpStatus_Alterar Then
    'Alterar EmpTroca
    objEmpTroca.AlterarEmpTroca lngPKID, _
                                lngTIPOPGTOID, _
                                mskValor.ClipText, _
                                txtNome.Text
    blnRetorno = True
    blnFechar = True
    Unload Me
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir EmpTroca
    strData = Format(Now, "DD/MM/YYYY hh:mm")
    
    objEmpTroca.InserirEmpTroca lngTURNOID, _
                                lngTIPOPGTOID, _
                                mskValor.ClipText, _
                                strData, _
                                txtNome.Text
    blnRetorno = True
    blnFechar = True
    Unload Me
    '
  End If
  Set objEmpTroca = Nothing
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
  If Not Valida_String(txtNome, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o nome válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(cboTipoPgto, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o tipo de pagamento válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskValor, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o valor válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserEmpTrocaInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserEmpTrocaInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco txtNome
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserEmpTrocaInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objEmpTroca             As busSisMaq.clsEmpTroca
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 3870
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  'Limpar Campos
  LimparCampos
  'TipoPgto
  strSql = "Select TIPOPGTO.TIPOPGTO " & _
        " FROM TIPOPGTO " & _
        " ORDER BY TIPOPGTO.TIPOPGTO"
  PreencheCombo cboTipoPgto, strSql, False, True
  
  If Status = tpStatus_Incluir Then
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objEmpTroca = New busSisMaq.clsEmpTroca
    Set objRs = objEmpTroca.SelecionarEmpTrocaPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      txtTurno.Text = RetornaDescTurnoCorrente(objRs.Fields("TURNOID").Value)
      txtData.Text = Format(objRs.Fields("DATA").Value, "DD/MM/YYYY hh:mm")
      txtNome.Text = objRs.Fields("NOME").Value & ""
      cboTipoPgto.Text = objRs.Fields("TIPOPGTO").Value & ""
      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VALOR").Value, TpMaskMoeda
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objEmpTroca = Nothing
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

Private Sub mskValor_GotFocus()
  Seleciona_Conteudo_Controle mskValor
End Sub
Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub

Private Sub txtNome_GotFocus()
  Seleciona_Conteudo_Controle txtNome
End Sub
Private Sub txtNome_LostFocus()
  Pintar_Controle txtNome, tpCorContr_Normal
End Sub

