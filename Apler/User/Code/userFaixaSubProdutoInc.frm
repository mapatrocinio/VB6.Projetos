VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserFaixaSubProdutoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Faixa de sub-produto"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4890
      Left            =   8430
      ScaleHeight     =   4890
      ScaleWidth      =   1860
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2055
         Left            =   90
         ScaleHeight     =   1995
         ScaleWidth      =   1605
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2730
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   150
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4665
      Left            =   120
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   8229
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da faixa do sub-produto"
      TabPicture(0)   =   "userFaixaSubProdutoInc.frx":0000
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
         Height          =   3855
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   3585
            Index           =   0
            Left            =   120
            ScaleHeight     =   3585
            ScaleWidth      =   7575
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtSubProduto 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   2
               TabStop         =   0   'False
               Text            =   "txtSubProduto"
               Top             =   690
               Width           =   6075
            End
            Begin VB.TextBox txtProduto 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   1
               TabStop         =   0   'False
               Text            =   "txtProduto"
               Top             =   390
               Width           =   6075
            End
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
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   2250
               Width           =   2235
               Begin VB.OptionButton optStatus 
                  Caption         =   "Inativo"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   11
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton optStatus 
                  Caption         =   "Ativo"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   10
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   825
               End
            End
            Begin VB.TextBox txtFaixa 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   3
               Text            =   "txtFaixa"
               Top             =   1020
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskValor 
               Height          =   255
               Left            =   1320
               TabIndex        =   8
               Top             =   1950
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskPcCusto 
               Height          =   255
               Left            =   5820
               TabIndex        =   9
               Top             =   1950
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDtInicio 
               Height          =   255
               Left            =   1320
               TabIndex        =   4
               Top             =   1350
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDtFim 
               Height          =   255
               Left            =   5820
               TabIndex        =   5
               Top             =   1350
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskFxInicial 
               Height          =   255
               Left            =   1320
               TabIndex        =   6
               Top             =   1650
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskFxFinal 
               Height          =   255
               Left            =   5820
               TabIndex        =   7
               Top             =   1650
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Fx. Inicial"
               Height          =   195
               Index           =   7
               Left            =   60
               TabIndex        =   30
               Top             =   1665
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Fx. Final"
               Height          =   195
               Index           =   6
               Left            =   4560
               TabIndex        =   29
               Top             =   1650
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Dt. Início"
               Height          =   255
               Index           =   1
               Left            =   60
               TabIndex        =   28
               Top             =   1350
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Dt. Fim"
               Height          =   255
               Index           =   2
               Left            =   4560
               TabIndex        =   27
               Top             =   1350
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Sub-Produto"
               Height          =   195
               Index           =   4
               Left            =   60
               TabIndex        =   26
               Top             =   735
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Produto"
               Height          =   195
               Index           =   3
               Left            =   60
               TabIndex        =   25
               Top             =   435
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Pç. Custo"
               Height          =   195
               Index           =   2
               Left            =   4560
               TabIndex        =   24
               Top             =   1950
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Convênio"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   23
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Valor"
               Height          =   195
               Index           =   21
               Left            =   60
               TabIndex        =   22
               Top             =   1965
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Status"
               Height          =   315
               Index           =   5
               Left            =   60
               TabIndex        =   20
               Top             =   2280
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Faixa"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   19
               Top             =   1080
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserFaixaSubProdutoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long
Public lngSUBPRODUTOID          As Long
Public strDescrConvenio         As String
Public strDescrProduto          As String
Public strDescrSubProduto       As String

Private blnPrimeiraVez          As Boolean

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Valor Produto Convênio
  LimparCampoTexto txtConvenio
  LimparCampoTexto txtProduto
  LimparCampoTexto txtSubProduto
  LimparCampoTexto txtFaixa
  LimparCampoMask mskDtInicio
  LimparCampoMask mskDtFim
  LimparCampoMask mskFxInicial
  LimparCampoMask mskFxFinal
  LimparCampoMask mskValor
  LimparCampoMask mskPcCusto
  optStatus(0).Value = False
  optStatus(1).Value = False
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserFaixaSubProdutoInc.LimparCampos]", _
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
  Dim objFaixaSubProduto        As busApler.clsFaixaSubProduto
  Dim objGeral                  As busApler.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim strStatus                 As String
  Dim strDatDesativacao         As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busApler.clsGeral
  Set objFaixaSubProduto = New busApler.clsFaixaSubProduto
  'Status
  If optStatus(0).Value Then
    strStatus = "A"
  Else
    strStatus = "I"
  End If

  'Validar se faixa para sub produto já cadastrada
  strSql = "SELECT * FROM FAIXA " & _
    " WHERE FAIXA.DESCRICAO = " & Formata_Dados(txtFaixa.Text, tpDados_Texto) & _
    " AND FAIXA.SUBPRODUTOID = " & Formata_Dados(lngSUBPRODUTOID, tpDados_Longo) & _
    " AND FAIXA.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle txtFaixa, tpCorContr_Erro
    TratarErroPrevisto "Faixa já cadastrada para sub-produto"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objFaixaSubProduto = Nothing
    cmdOk.Enabled = True
    SetarFoco txtFaixa
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar FaixaSubProduto
    strDatDesativacao = ""
    If strStatus = "I" Then
      strDatDesativacao = Format(Now, "DD/MM/YYYY hh:mm")
    End If
    objFaixaSubProduto.AlterarFaixaSubProduto lngPKID, _
                                              txtFaixa.Text, _
                                              mskDtInicio.Text, _
                                              IIf(mskDtFim.Text = "__/__/____", "", mskDtFim.Text), _
                                              strDatDesativacao, _
                                              mskFxInicial.Text, _
                                              mskFxFinal.Text, _
                                              mskValor.Text, _
                                              mskPcCusto.Text, _
                                              strStatus
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir FaixaSubProduto
    objFaixaSubProduto.InserirFaixaSubProduto lngSUBPRODUTOID, _
                                              txtFaixa.Text, _
                                              mskDtInicio.Text, _
                                              IIf(mskDtFim.Text = "__/__/____", "", mskDtFim.Text), _
                                              mskFxInicial.Text, _
                                              mskFxFinal.Text, _
                                              mskValor.Text, _
                                              mskPcCusto.Text
  End If
  Set objFaixaSubProduto = Nothing
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
  If Not Valida_String(txtFaixa, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a faixa do sub-produto" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Data(mskDtInicio, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher data inicial da faixa" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Data(mskDtFim, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher data final da faixa" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskFxInicial, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a faixa inicial da faixa" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskFxFinal, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher faixa final da faixa" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskValor, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o valor válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskPcCusto, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o preço de custo válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Option(optStatus, blnSetarFocoControle) Then
    strMsg = strMsg & "Slecionar o status" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserFaixaSubProdutoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserFaixaSubProdutoInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco txtFaixa
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserFaixaSubProdutoInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objFaixaSubProduto      As busApler.clsFaixaSubProduto
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 5370
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  'Limpar Campos
  LimparCampos
  txtConvenio.Text = strDescrConvenio
  txtProduto.Text = strDescrProduto
  txtSubProduto.Text = strDescrSubProduto
  If Status = tpStatus_Incluir Then
    '
    optStatus(0).Value = True
    'Visible
    optStatus(0).Visible = False
    optStatus(1).Visible = False
    Label5(5).Visible = False
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objFaixaSubProduto = New busApler.clsFaixaSubProduto
    Set objRs = objFaixaSubProduto.SelecionarFaixaSubProdutoPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      txtFaixa.Text = objRs.Fields("DESCRICAO").Value & ""
      INCLUIR_VALOR_NO_MASK mskDtInicio, objRs.Fields("DATAINICIO").Value & "", TpMaskData
      INCLUIR_VALOR_NO_MASK mskDtFim, objRs.Fields("DATAFIM").Value & "", TpMaskData
      INCLUIR_VALOR_NO_MASK mskFxInicial, objRs.Fields("FXINICIAL").Value & "", TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskFxFinal, objRs.Fields("FXFINAL").Value & "", TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VALOR").Value & "", TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskPcCusto, objRs.Fields("PRCUSTO").Value & "", TpMaskMoeda
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
    Set objFaixaSubProduto = Nothing
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

Private Sub mskDtFim_GotFocus()
  Seleciona_Conteudo_Controle mskDtFim
End Sub
Private Sub mskDtFim_LostFocus()
  Pintar_Controle mskDtFim, tpCorContr_Normal
End Sub

Private Sub mskDtInicio_GotFocus()
  Seleciona_Conteudo_Controle mskDtInicio
End Sub
Private Sub mskDtInicio_LostFocus()
  Pintar_Controle mskDtInicio, tpCorContr_Normal
End Sub

Private Sub mskFxFinal_GotFocus()
  Seleciona_Conteudo_Controle mskFxFinal
End Sub
Private Sub mskFxFinal_LostFocus()
  Pintar_Controle mskFxFinal, tpCorContr_Normal
End Sub

Private Sub mskFxInicial_GotFocus()
  Seleciona_Conteudo_Controle mskFxInicial
End Sub
Private Sub mskFxInicial_LostFocus()
  Pintar_Controle mskFxInicial, tpCorContr_Normal
End Sub

Private Sub mskPcCusto_GotFocus()
  Seleciona_Conteudo_Controle mskPcCusto
End Sub
Private Sub mskPcCusto_LostFocus()
  Pintar_Controle mskPcCusto, tpCorContr_Normal
End Sub
Private Sub mskValor_GotFocus()
  Seleciona_Conteudo_Controle mskValor
End Sub
Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
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
    SetarFoco txtFaixa
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "Apler.frmUserFaixaSubProdutoInc.tabDetalhes"
  AmpN
End Sub

Private Sub txtFaixa_GotFocus()
  Seleciona_Conteudo_Controle txtFaixa
End Sub
Private Sub txtFaixa_LostFocus()
  Pintar_Controle txtFaixa, tpCorContr_Normal
End Sub
