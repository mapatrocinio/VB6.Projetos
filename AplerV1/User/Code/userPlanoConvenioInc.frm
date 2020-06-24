VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmUserPlanoConvenioInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Plano de Convênio"
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   90
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1830
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   150
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   3795
      Left            =   120
      TabIndex        =   14
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
      TabPicture(0)   =   "userPlanoConvenioInc.frx":0000
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
         TabIndex        =   15
         Top             =   420
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2385
            Index           =   0
            Left            =   120
            ScaleHeight     =   2385
            ScaleWidth      =   7575
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtConvenio 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtConvenio"
               Top             =   90
               Width           =   6075
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1290
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   1590
               Width           =   2235
               Begin VB.OptionButton optStatus 
                  Caption         =   "Inativo"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   9
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton optStatus 
                  Caption         =   "Ativo"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   8
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   825
               End
            End
            Begin VB.TextBox txtPlanoConvenio 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   1
               Text            =   "txtPlanoConvenio"
               Top             =   405
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskVrSocio 
               Height          =   255
               Left            =   1320
               TabIndex        =   2
               Top             =   720
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskVrDependente 
               Height          =   255
               Left            =   5820
               TabIndex        =   3
               Top             =   720
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskVrSocioApler 
               Height          =   255
               Left            =   1320
               TabIndex        =   4
               Top             =   1020
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskVrDependenteApler 
               Height          =   255
               Left            =   5820
               TabIndex        =   5
               Top             =   1020
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskInicio 
               Height          =   255
               Left            =   1320
               TabIndex        =   6
               Top             =   1320
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   16777215
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskFim 
               Height          =   255
               Left            =   5820
               TabIndex        =   7
               Top             =   1320
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   16777215
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Fim"
               Height          =   195
               Index           =   7
               Left            =   4560
               TabIndex        =   26
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Início"
               Height          =   195
               Index           =   6
               Left            =   60
               TabIndex        =   25
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Vr. Sócio Apler"
               Height          =   195
               Index           =   4
               Left            =   60
               TabIndex        =   24
               Top             =   1035
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Vr. Dep. Apler"
               Height          =   195
               Index           =   3
               Left            =   4560
               TabIndex        =   23
               Top             =   1020
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Vr. Dependente"
               Height          =   195
               Index           =   2
               Left            =   4560
               TabIndex        =   22
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Convênio"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   21
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Vr. Sócio"
               Height          =   195
               Index           =   21
               Left            =   60
               TabIndex        =   20
               Top             =   735
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Status"
               Height          =   315
               Index           =   5
               Left            =   60
               TabIndex        =   18
               Top             =   1620
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Plano Convênio"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   17
               Top             =   450
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserPlanoConvenioInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long
Public lngCONVENIOID            As Long
Public strDescrConvenio         As String

Private blnPrimeiraVez          As Boolean

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Valor Plano Convênio
  LimparCampoTexto txtConvenio
  LimparCampoTexto txtPlanoConvenio
  LimparCampoMask mskVrSocio
  LimparCampoMask mskVrDependente
  LimparCampoMask mskVrSocioApler
  LimparCampoMask mskVrDependenteApler
  LimparCampoMask mskInicio
  LimparCampoMask mskFim
  optStatus(0).Value = False
  optStatus(1).Value = False
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserPlanoConvenioInc.LimparCampos]", _
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
  Dim objPlanoConvenio          As busApler.clsPlanoConvenio
  Dim objGeral                  As busApler.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim strStatus                 As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busApler.clsGeral
  Set objPlanoConvenio = New busApler.clsPlanoConvenio
  'Status
  If optStatus(0).Value Then
    strStatus = "A"
  Else
    strStatus = "I"
  End If

  'Validar se plano convênio já cadastrado
  strSql = "SELECT * FROM PLANOCONVENIO " & _
    " WHERE PLANOCONVENIO.NOME = " & Formata_Dados(txtPlanoConvenio.Text, tpDados_Texto) & _
    " AND PLANOCONVENIO.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle txtPlanoConvenio, tpCorContr_Erro
    TratarErroPrevisto "Plano do convênio já cadastrado"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objPlanoConvenio = Nothing
    cmdOk.Enabled = True
    SetarFoco txtPlanoConvenio
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar PlanoConvenio
    objPlanoConvenio.AlterarPlanoConvenio lngPKID, _
                                          txtPlanoConvenio.Text, _
                                          mskVrSocio.Text, _
                                          mskVrDependente.Text, _
                                          mskVrSocioApler.Text, _
                                          mskVrDependenteApler.Text, _
                                          mskInicio.Text, _
                                          mskFim.Text, _
                                          strStatus
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir PlanoConvenio
    objPlanoConvenio.InserirPlanoConvenio lngCONVENIOID, _
                                          txtPlanoConvenio.Text, _
                                          mskVrSocio.Text, _
                                          mskVrDependente.Text, _
                                          mskVrSocioApler.Text, _
                                          mskVrDependenteApler.Text, _
                                          mskInicio.Text, _
                                          mskFim.Text
  End If
  Set objPlanoConvenio = Nothing
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
  If Not Valida_String(txtPlanoConvenio, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o plano do convênio" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskVrSocio, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o valor para o sócio" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskVrDependente, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o valor para dependente" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskVrSocioApler, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o valor para o sócio apler" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskVrDependenteApler, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o valor para dependente apler" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Data(mskInicio, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de início" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Data(mskFim, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de fim válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Option(optStatus, blnSetarFocoControle) Then
    strMsg = strMsg & "Slecionar o status" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserPlanoConvenioInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserPlanoConvenioInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco txtPlanoConvenio
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserPlanoConvenioInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objPlanoConvenio           As busApler.clsPlanoConvenio
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
  txtConvenio.Text = strDescrConvenio
  If Status = tpStatus_Incluir Then
    '
    optStatus(0).Value = True
    'Visible
    optStatus(0).Visible = False
    optStatus(1).Visible = False
    Label5(5).Visible = False
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objPlanoConvenio = New busApler.clsPlanoConvenio
    Set objRs = objPlanoConvenio.SelecionarPlanoConvenioPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      txtPlanoConvenio.Text = objRs.Fields("NOME").Value & ""
      INCLUIR_VALOR_NO_MASK mskVrSocio, objRs.Fields("VALORSOCIO").Value & "", TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskVrDependente, objRs.Fields("VALORDEPENDENTE").Value & "", TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskVrSocioApler, objRs.Fields("VALORAPLERSOCIO").Value & "", TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskVrDependenteApler, objRs.Fields("VALORAPLERDEPENDENTE").Value & "", TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskInicio, objRs.Fields("DATAINICIO").Value & "", TpMaskData
      INCLUIR_VALOR_NO_MASK mskFim, objRs.Fields("DATAFIM").Value & "", TpMaskData
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
    Set objPlanoConvenio = Nothing
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

Private Sub mskVrDependente_GotFocus()
  Seleciona_Conteudo_Controle mskVrDependente
End Sub
Private Sub mskVrDependente_LostFocus()
  Pintar_Controle mskVrDependente, tpCorContr_Normal
End Sub

Private Sub mskVrDependenteApler_GotFocus()
  Seleciona_Conteudo_Controle mskVrDependenteApler
End Sub
Private Sub mskVrDependenteApler_LostFocus()
  Pintar_Controle mskVrDependenteApler, tpCorContr_Normal
End Sub

Private Sub mskVrSocio_GotFocus()
  Seleciona_Conteudo_Controle mskVrSocio
End Sub
Private Sub mskVrSocio_LostFocus()
  Pintar_Controle mskVrSocio, tpCorContr_Normal
End Sub

Private Sub mskVrSocioApler_GotFocus()
  Seleciona_Conteudo_Controle mskVrSocioApler
End Sub
Private Sub mskVrSocioApler_LostFocus()
  Pintar_Controle mskVrSocioApler, tpCorContr_Normal
End Sub
Private Sub txtPlanoConvenio_GotFocus()
  Seleciona_Conteudo_Controle txtPlanoConvenio
End Sub
Private Sub txtPlanoConvenio_LostFocus()
  Pintar_Controle txtPlanoConvenio, tpCorContr_Normal
End Sub

