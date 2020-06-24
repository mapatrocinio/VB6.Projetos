VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserAtendeInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de grade de atendimento"
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   90
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1830
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   150
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   3795
      Left            =   120
      TabIndex        =   11
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
      TabPicture(0)   =   "userAtendeInc.frx":0000
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
         Top             =   420
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
            Top             =   150
            Width           =   7575
            Begin VB.ComboBox cboPrestador 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   720
               Width           =   6105
            End
            Begin VB.ComboBox cboDiaDaSemana 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   390
               Width           =   6105
            End
            Begin VB.TextBox txtSala 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtSala"
               Top             =   90
               Width           =   6075
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1290
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   1320
               Width           =   2235
               Begin VB.OptionButton optStatus 
                  Caption         =   "Inativo"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   6
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton optStatus 
                  Caption         =   "Ativo"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   5
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   825
               End
            End
            Begin MSMask.MaskEdBox mskInicio 
               Height          =   255
               Left            =   1320
               TabIndex        =   3
               Top             =   1050
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   16777215
               AutoTab         =   -1  'True
               MaxLength       =   5
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskFim 
               Height          =   255
               Left            =   5820
               TabIndex        =   4
               Top             =   1050
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   16777215
               AutoTab         =   -1  'True
               MaxLength       =   5
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Fim"
               Height          =   195
               Index           =   7
               Left            =   4560
               TabIndex        =   20
               Top             =   1050
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Início"
               Height          =   195
               Index           =   6
               Left            =   60
               TabIndex        =   19
               Top             =   1050
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Sala"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   18
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Prestador"
               Height          =   195
               Index           =   21
               Left            =   60
               TabIndex        =   17
               Top             =   735
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Status"
               Height          =   315
               Index           =   5
               Left            =   60
               TabIndex        =   15
               Top             =   1350
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Dia da semana"
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
Attribute VB_Name = "frmUserAtendeInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long
Public lngSALAID                As Long
Public strDescrSala         As String

Private blnPrimeiraVez          As Boolean

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Valor Plano Convênio
  LimparCampoTexto txtSala
  LimparCampoCombo cboDiaDaSemana
  LimparCampoCombo cboPrestador
  LimparCampoMask mskInicio
  LimparCampoMask mskFim
  optStatus(0).Value = False
  optStatus(1).Value = False
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserAtendeInc.LimparCampos]", _
            Err.Description
End Sub

Private Sub cboDiaDaSemana_LostFocus()
  Pintar_Controle cboDiaDaSemana, tpCorContr_Normal
End Sub

Private Sub cboPrestador_LostFocus()
  Pintar_Controle cboPrestador, tpCorContr_Normal
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
  Dim objAtende          As busSisMed.clsAtende
  Dim objGeral                  As busSisMed.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim strStatus                 As String
  Dim lngDIASDASEMANAID         As Long
  Dim lngPRONTUARIOID           As Long
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busSisMed.clsGeral
  Set objAtende = New busSisMed.clsAtende
  'Status
  If optStatus(0).Value Then
    strStatus = "A"
  Else
    strStatus = "I"
  End If
  'DIASDASEMANA
  lngDIASDASEMANAID = 0
  strSql = "SELECT PKID FROM DIASDASEMANA WHERE DIADASEMANA = " & Formata_Dados(cboDiaDaSemana.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngDIASDASEMANAID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'PRESTADOR
  lngPRONTUARIOID = 0
  strSql = "SELECT PKID FROM PRONTUARIO INNER JOIN PRESTADOR ON PRONTUARIO.PKID = PRESTADOR.PRONTUARIOID " & _
    " Where PRONTUARIO.NOME  = " & Formata_Dados(cboPrestador.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngPRONTUARIOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'Validar se grade de atendimento já cadastrado
  strSql = "SELECT * FROM ATENDE " & _
    " WHERE ATENDE.DIASDASEMANAID = " & Formata_Dados(lngDIASDASEMANAID, tpDados_Longo) & _
    " AND ATENDE.SALAID = " & Formata_Dados(lngSALAID, tpDados_Longo) & _
    " AND ATENDE.HORAINICIO = " & Formata_Dados(mskInicio.Text, tpDados_Texto) & _
    " AND ATENDE.HORATERMINO = " & Formata_Dados(mskFim.Text, tpDados_Texto) & _
    " AND ATENDE.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle cboDiaDaSemana, tpCorContr_Erro
    TratarErroPrevisto "Grade de atendimento já cadastrada"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objAtende = Nothing
    cmdOk.Enabled = True
    SetarFoco cboDiaDaSemana
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar Atende
    objAtende.AlterarAtende lngPKID, _
                            lngSALAID, _
                            lngPRONTUARIOID, _
                            lngDIASDASEMANAID, _
                            mskInicio.Text, _
                            mskFim.Text, _
                            strStatus
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir Atende
    objAtende.InserirAtende lngSALAID, _
                            lngPRONTUARIOID, _
                            lngDIASDASEMANAID, _
                            mskInicio.Text, _
                            mskFim.Text, _
                            strStatus
  End If
  Set objAtende = Nothing
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
  If Not Valida_String(cboDiaDaSemana, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o dia da semana" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(cboPrestador, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o prestador" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Hora(mskInicio, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a hora de início" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Hora(mskFim, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a hora de fim válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Option(optStatus, blnSetarFocoControle) Then
    strMsg = strMsg & "Slecionar o status" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserAtendeInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserAtendeInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco cboDiaDaSemana
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserAtendeInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objAtende           As busSisMed.clsAtende
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
  txtSala.Text = strDescrSala
  'Dias da semana
  strSql = "Select DIADASEMANA from DIASDASEMANA ORDER BY CODIGO"
  PreencheCombo cboDiaDaSemana, strSql, False, True
  'Prestador
  strSql = "Select PRONTUARIO.NOME from PRONTUARIO INNER JOIN PRESTADOR ON PRONTUARIO.PKID = PRESTADOR.PRONTUARIOID ORDER BY PRONTUARIO.NOME"
  PreencheCombo cboPrestador, strSql, False, True
  '
  If Status = tpStatus_Incluir Then
    '
    optStatus(0).Value = True
    'Visible
    optStatus(0).Visible = False
    optStatus(1).Visible = False
    Label5(5).Visible = False
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objAtende = New busSisMed.clsAtende
    Set objRs = objAtende.SelecionarAtendePeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      cboDiaDaSemana.Text = objRs.Fields("DIADASEMANA").Value & ""
      cboPrestador.Text = objRs.Fields("NOME").Value & ""
      INCLUIR_VALOR_NO_MASK mskInicio, objRs.Fields("HORAINICIO").Value & "", TpMaskData
      INCLUIR_VALOR_NO_MASK mskFim, objRs.Fields("HORATERMINO").Value & "", TpMaskData
      If objRs.Fields("STATUS").Value & "" = "A" Then
        optStatus(0).Value = True
        optStatus(1).Value = False
      ElseIf objRs.Fields("STATUS").Value & "" = "I" Then
        optStatus(0).Value = False
        optStatus(1).Value = True
      Else
        optStatus(0).Value = False
        optStatus(1).Value = False
      End If
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objAtende = Nothing
    'Visible
    optStatus(0).Visible = True
    optStatus(1).Visible = True
    Label5(5).Visible = True
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

Private Sub mskFim_GotFocus()
  Seleciona_Conteudo_Controle mskFim
End Sub
Private Sub mskFim_LostFocus()
  Pintar_Controle mskFim, tpCorContr_Normal
End Sub

Private Sub mskInicio_GotFocus()
  Seleciona_Conteudo_Controle mskInicio
End Sub
Private Sub mskInicio_LostFocus()
  Pintar_Controle mskInicio, tpCorContr_Normal
End Sub
