VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmServicoPlacaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alterar Veículo"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   2865
      Left            =   8430
      ScaleHeight     =   2865
      ScaleWidth      =   1860
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2055
         Left            =   90
         ScaleHeight     =   1995
         ScaleWidth      =   1605
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   630
         Width           =   1665
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   150
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   990
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   2565
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   4524
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userServicoPlacaInc.frx":0000
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
         Height          =   1965
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   1575
            Index           =   0
            Left            =   120
            ScaleHeight     =   1575
            ScaleWidth      =   7575
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.ComboBox cboVeiculo 
               Height          =   315
               Left            =   1380
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   330
               Width           =   5955
            End
            Begin MSMask.MaskEdBox mskPlaca 
               Height          =   255
               Left            =   1380
               TabIndex        =   0
               TabStop         =   0   'False
               Top             =   60
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   14737632
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               MaxLength       =   8
               Mask            =   "???-####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label10 
               Caption         =   "Placa atual"
               Height          =   255
               Left            =   90
               TabIndex        =   10
               Top             =   60
               Width           =   1245
            End
            Begin VB.Label Label6 
               Caption         =   "Veículo"
               Height          =   255
               Index           =   0
               Left            =   60
               TabIndex        =   9
               Top             =   360
               Width           =   1455
            End
         End
      End
   End
End
Attribute VB_Name = "frmServicoPlacaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPACOTESERVICOID       As Long
Public strCaption               As String
Public strPlaca                 As String

Private blnPrimeiraVez          As Boolean

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'ServicoPlaca
  LimparCampoMask mskPlaca
  LimparCampoCombo cboVeiculo
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmServicoPlacaInc.LimparCampos]", _
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
  Dim objServico           As busElite.clsServico
  Dim objGeral      As busElite.clsGeral
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  Dim lngVEICULOID  As Long
  Dim strPlaca      As String
  
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objServico = New busElite.clsServico
  'VEICULOID
  Set objGeral = New busElite.clsGeral
  strPlaca = ""
  strPlaca = Left(Right(cboVeiculo.Text, 9), 8)
  lngVEICULOID = 0
  strSql = "SELECT PKID FROM VEICULO WHERE VEICULO.PLACA = " & Formata_Dados(strPlaca, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngVEICULOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'Valida se obterve os campos com sucesso
  If lngVEICULOID = 0 Then
    Set objGeral = Nothing
    TratarErroPrevisto "Selecionar o veículo", "cmdOK_Click"
    Pintar_Controle cboVeiculo, tpCorContr_Erro
    SetarFoco cboVeiculo
    Exit Sub
  End If
  Set objGeral = Nothing

  If Status = tpStatus_Incluir Then
    'Alterar ServicoPlaca
    objServico.AlterarVeiculo lngPACOTESERVICOID, _
                              lngVEICULOID
    blnRetorno = True
    blnFechar = True
    Unload Me
    '
  End If
  Set objServico = Nothing
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
  If Not Valida_String(cboVeiculo, TpObrigatorio, False) Then
    strMsg = strMsg & "Selecionar o veículo" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmServicoPlacaInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmServicoPlacaInc.ValidaCampos]", _
            Err.Description
End Function



Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco cboVeiculo
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmServicoPlacaInc.Form_Activate]"
End Sub



Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objServico              As busElite.clsServico
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 3345
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  'Limpar Campos
  LimparCampos
  tabDetalhes_Click 0
  '
  'Combos
  'VAICULO
  strSql = "Select MODELO.NOME + ' (' + VEICULO.PLACA + ')' " & _
      " FROM VEICULO " & _
      " INNER JOIN MODELO ON MODELO.PKID = VEICULO.MODELOID " & _
      "ORDER BY MODELO.NOME, VEICULO.PLACA"
  PreencheCombo cboVeiculo, strSql, False, True
  '
  Me.Caption = Me.Caption & " - " & strCaption
  INCLUIR_VALOR_NO_MASK mskPlaca, strPlaca, TpMaskOutros
  If Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
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



Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    pictrava(0).Enabled = True
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    '
    SetarFoco cboVeiculo
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "Elite.frmServicoPlacaInc.tabDetalhes"
  AmpN
End Sub
