VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmUserReceitaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Receitas"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5385
      Left            =   8250
      ScaleHeight     =   5385
      ScaleWidth      =   1860
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   90
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   16
         Top             =   3180
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5175
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da despesa"
      TabPicture(0)   =   "userReceitaInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cadastro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4545
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   7575
         Begin VB.TextBox txtCodigo 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
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
            Height          =   3795
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   7335
            Begin VB.ComboBox cboContrato 
               Height          =   315
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   2010
               Width           =   5655
            End
            Begin VB.CommandButton cmdConsultar 
               Caption         =   "&Z"
               Height          =   880
               Left            =   3960
               Style           =   1  'Graphical
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   180
               Width           =   1335
            End
            Begin VB.TextBox txtCheque 
               Height          =   285
               Left            =   5280
               MaxLength       =   15
               TabIndex        =   10
               Top             =   2730
               Width           =   1935
            End
            Begin VB.ComboBox cboLivro 
               Height          =   315
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   2730
               Width           =   2175
            End
            Begin VB.ComboBox cboFormaPgto 
               Height          =   315
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   2370
               Width           =   2175
            End
            Begin VB.TextBox txtDescricao 
               Height          =   525
               Left            =   1560
               MaxLength       =   100
               MultiLine       =   -1  'True
               TabIndex        =   6
               Text            =   "userReceitaInc.frx":001C
               Top             =   1440
               Width           =   5655
            End
            Begin MSMask.MaskEdBox mskGrupo 
               Height          =   255
               Left            =   1560
               TabIndex        =   1
               Top             =   240
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskSubGrupo 
               Height          =   255
               Left            =   2160
               TabIndex        =   2
               Top             =   240
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskValorRecebido 
               Height          =   255
               Left            =   5280
               TabIndex        =   12
               Top             =   3090
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskData 
               Height          =   255
               Index           =   0
               Left            =   1560
               TabIndex        =   11
               Top             =   3090
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskValorReceber 
               Height          =   255
               Left            =   1560
               TabIndex        =   5
               Top             =   1080
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskData 
               Height          =   255
               Index           =   1
               Left            =   1560
               TabIndex        =   4
               Top             =   720
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label6 
               Caption         =   "Contrato"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   2010
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "<----------------------"
               Height          =   255
               Left            =   2850
               TabIndex        =   30
               Top             =   240
               Width           =   1005
            End
            Begin VB.Label lblCheque 
               Caption         =   "Número cheque"
               Height          =   255
               Left            =   3960
               TabIndex        =   29
               Top             =   2730
               Width           =   1215
            End
            Begin VB.Label lblLivro 
               Caption         =   "Livro"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   2730
               Width           =   1215
            End
            Begin VB.Label Label4 
               Caption         =   "Valor a Receber"
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   1080
               Width           =   1425
            End
            Begin VB.Label Label3 
               Caption         =   "Dt. Venc."
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label8 
               Caption         =   "Valor Recebido"
               Height          =   255
               Left            =   3960
               TabIndex        =   25
               Top             =   3090
               Width           =   1095
            End
            Begin VB.Label Da 
               Caption         =   "Dt. Rceb."
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   3090
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "Forma Pgto."
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   2370
               Width           =   1215
            End
            Begin VB.Label Label2 
               Caption         =   "Grupo/Sub Grupo"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label9 
               Caption         =   "Descrição"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   1440
               Width           =   1215
            End
         End
         Begin VB.Label Label44 
            Caption         =   "Sequencial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "frmUserReceitaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                         As tpStatus
Public lngRECEITAID                   As Long
Public bRetorno                       As Boolean
Public blnPrimeiraVez                 As Boolean
Public bFechar                        As Boolean
Public strTipo                        As String
Public strTipoCtaRecebidas            As String



Private Sub cboContrato_LostFocus()
  Pintar_Controle cboContrato, tpCorContr_Normal
End Sub

Private Sub cboFormaPgto_Click()
  If UCase(cboFormaPgto.Text) = "CHEQUE" Then
    lblLivro.Enabled = True
    cboLivro.Enabled = True
    lblCheque.Enabled = True
    txtCheque.Enabled = True
  Else
    lblLivro.Enabled = False
    cboLivro.Enabled = False
    lblCheque.Enabled = False
    txtCheque.Enabled = False
    txtCheque.Text = ""
    cboLivro.ListIndex = -1
  End If
End Sub

Private Sub cboFormaPgto_LostFocus()
  Pintar_Controle cboFormaPgto, tpCorContr_Normal
End Sub

Private Sub cboLivro_LostFocus()
  Pintar_Controle cboLivro, tpCorContr_Normal
End Sub

Private Sub cmdCancelar_Click()
  bFechar = True
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdConsultar_Click()
  On Error GoTo trata
  frmUserGrupoDespesaCons.QuemChamou = 2
  frmUserGrupoDespesaCons.Show vbModal
  SetarFoco mskGrupo
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdOK_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim strMsgErro              As String
  Dim objRs                   As ADODB.Recordset
  Dim objReceita              As busSisContas.clsReceita
  Dim clsGer                  As busSisContas.clsGeral
  Dim lngFORMAPGTOID          As Long
  Dim lngSUBGRUPOID           As Long
  Dim lngLIVROID              As Long
  Dim lngSEQUENCIALEXTERNO    As Long
  Dim lngEMPRESACTRTOID       As Long
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração da Receita
    If Not ValidaCampos Then Exit Sub
    'Valida se unidade de estoque já cadastrada
    Set clsGer = New busSisContas.clsGeral
    If cboLivro.Text <> "" Then
      strSql = "Select PKID From LIVRO WHERE NUMEROLIVRO = " & Formata_Dados(cboLivro.Text, tpDados_Texto, tpNulo_NaoAceita) & _
                  " AND PARCEIROID = " & Formata_Dados(glParceiroId, tpDados_Longo)
      Set objRs = clsGer.ExecutarSQL(strSql)
      If objRs.EOF Then
        objRs.Close
        Set objRs = Nothing
        Set clsGer = Nothing
        TratarErroPrevisto "Livro não cadastrado", "cmdOK_Click"
        Exit Sub
        
      Else
        lngLIVROID = objRs.Fields("PKID").Value
      End If
      objRs.Close
      Set objRs = Nothing
    End If
    strSql = "Select EMPRESACTRTO.PKID From EMPRESACTRTO INNER JOIN EMPRESA ON EMPRESA.PKID = EMPRESACTRTO.EMPRESAID " & _
            " WHERE EMPRESA.PARCEIROID = " & Formata_Dados(glParceiroId, tpDados_Longo) & _
            " AND EMPRESACTRTO.DESCRICAO = " & Formata_Dados(cboContrato.Text, tpDados_Texto)
    Set objRs = clsGer.ExecutarSQL(strSql)
    If objRs.EOF Then
      'objRs.Close
      'Set objRs = Nothing
      'Set clsGer = Nothing
      'TratarErroPrevisto "Forma de Pagamento não cadastrada", "cmdOK_Click"
      'Exit Sub
      lngEMPRESACTRTOID = 0
    Else
      lngEMPRESACTRTOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    strSql = "Select PKID From FORMAPGTO WHERE FORMAPGTO = " & Formata_Dados(cboFormaPgto.Text, tpDados_Texto, tpNulo_NaoAceita)
    Set objRs = clsGer.ExecutarSQL(strSql)
    If objRs.EOF Then
      'objRs.Close
      'Set objRs = Nothing
      'Set clsGer = Nothing
      'TratarErroPrevisto "Forma de Pagamento não cadastrada", "cmdOK_Click"
      'Exit Sub
      lngFORMAPGTOID = 0
    Else
      lngFORMAPGTOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    strSql = "Select SUBGRUPODESPESA.PKID From GRUPODESPESA INNER JOIN SUBGRUPODESPESA ON GRUPODESPESA.PKID = SUBGRUPODESPESA.GRUPODESPESAID " & _
      " WHERE GRUPODESPESA.PARCEIROID = " & Formata_Dados(glParceiroId, tpDados_Longo) & _
      " AND GRUPODESPESA.CODIGO = " & Formata_Dados(mskGrupo.Text, tpDados_Texto, tpNulo_NaoAceita) & _
      " AND SUBGRUPODESPESA.CODIGO = " & Formata_Dados(mskSubGrupo.Text, tpDados_Texto, tpNulo_NaoAceita)
    Set objRs = clsGer.ExecutarSQL(strSql)
    If objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set clsGer = Nothing
      TratarErroPrevisto "Grupo/Subgrupo não cadastrado", "cmdOK_Click"
      SetarFoco mskGrupo
      Exit Sub
      
    Else
      lngSUBGRUPOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set clsGer = Nothing

    Set objReceita = New busSisContas.clsReceita
    
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objReceita.AlterarReceita IIf(Len(mskData(0).ClipText) = 0, "", mskData(0).Text), _
                                mskData(1).Text, _
                                mskValorReceber.Text, _
                                lngLIVROID, _
                                txtCheque.Text, _
                                lngRECEITAID, _
                                txtDescricao.Text, _
                                mskValorRecebido, _
                                lngSUBGRUPOID, _
                                lngFORMAPGTOID, _
                                gsNomeUsu, _
                                lngEMPRESACTRTOID
      Set objReceita = Nothing
      bRetorno = True
      bFechar = True
      Unload Me
    ElseIf Status = tpStatus_Incluir Then
      'Código para inclusão
      '
      lngSEQUENCIALEXTERNO = CLng(RetornaGravaSequencialV1("SEQUENCIALREC" & glParceiroId))
      objReceita.IncluirReceita mskData(1).Text, _
                                mskValorReceber.Text, _
                                lngLIVROID, _
                                txtCheque.Text, _
                                strTipo, _
                                IIf(Len(mskData(0).ClipText) = 0, "", mskData(0).Text), _
                                txtDescricao.Text, _
                                mskValorRecebido, _
                                lngSUBGRUPOID, _
                                lngFORMAPGTOID, _
                                gsNomeUsu, _
                                lngSEQUENCIALEXTERNO, _
                                lngEMPRESACTRTOID, _
                                glParceiroId
      MsgBox "Sequencial: " & lngSEQUENCIALEXTERNO, vbOKOnly, TITULOSISTEMA
      'Limpar campos
      LimparCampoMask mskData(1)
      LimparCampoMask mskValorReceber
      LimparCampoTexto txtCheque
      LimparCampoTexto txtCodigo
      LimparCampoMask mskData(0)
      LimparCampoMask mskValorRecebido
      LimparCampoTexto txtDescricao
      cboFormaPgto.ListIndex = -1
      LimparCampoMask mskGrupo
      LimparCampoMask mskSubGrupo
      cboContrato.ListIndex = -1
      SetarFoco mskGrupo
      Set objReceita = Nothing
      bRetorno = True
    End If
    'Set objReceita = Nothing
    'bFechar = True
    'Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  Dim strMsg              As String
  Dim objSubGrupoDespesa  As busSisContas.clsSubGrupoDespesa
  Dim strTipoReceita      As String
  '
  If Not Valida_Moeda(mskGrupo, TpObrigatorio) Then
    strMsg = strMsg & "Informar o Grupo válido" & vbCrLf
    Pintar_Controle mskGrupo, tpCorContr_Erro
  End If
  If Not Valida_Moeda(mskSubGrupo, TpObrigatorio) Then
    strMsg = strMsg & "Informar o Sub Grupo válido" & vbCrLf
    Pintar_Controle mskSubGrupo, tpCorContr_Erro
  End If
  If Not Valida_String(cboContrato, TpObrigatorio) Then
    strMsg = strMsg & "Selecionar o contrato" & vbCrLf
    Pintar_Controle cboContrato, tpCorContr_Erro
  End If
  If strMsg = "" Then
    'Digitou Grupo e subGrupo, Preencher
    Set objSubGrupoDespesa = New busSisContas.clsSubGrupoDespesa
    strTipoReceita = objSubGrupoDespesa.SelecionarTipoGrupo(mskGrupo.Text, _
                                                            mskSubGrupo.Text)
    Set objSubGrupoDespesa = Nothing
  End If
  If strTipoReceita <> "C" Then
    strMsg = strMsg & "Informar o Grupo do tipo Crédito" & vbCrLf
    Pintar_Controle mskGrupo, tpCorContr_Erro
  End If
  
  If Not Valida_Data(mskData(1), TpObrigatorio) Then
    strMsg = strMsg & "Informar a data de vencimento válida" & vbCrLf
    Pintar_Controle mskData(1), tpCorContr_Erro
  End If
  If Not Valida_Moeda(mskValorReceber, TpObrigatorio) Then
    strMsg = strMsg & "Informar o valor a receber válido" & vbCrLf
    Pintar_Controle mskValorReceber, tpCorContr_Erro
  End If
  If Len(txtDescricao.Text) = 0 Then
    strMsg = strMsg & "Informar a descrição da receita válida" & vbCrLf
    Pintar_Controle txtDescricao, tpCorContr_Erro
  End If
  If strTipoCtaRecebidas = "S" Then
    If Not Valida_Data(mskData(0), TpNaoObrigatorio) Then
      strMsg = strMsg & "Informar a data de recebimento válida" & vbCrLf
      Pintar_Controle mskData(0), tpCorContr_Erro
    End If
    If Not Valida_Moeda(mskValorRecebido, TpNaoObrigatorio) Then
      strMsg = strMsg & "Informar o valor recebido válido" & vbCrLf
      Pintar_Controle mskValorRecebido, tpCorContr_Erro
    End If
  Else
    If Not Valida_Data(mskData(0), TpObrigatorio) Then
      strMsg = strMsg & "Informar a data de recebimento válida" & vbCrLf
      Pintar_Controle mskData(0), tpCorContr_Erro
    End If
    If Not Valida_Moeda(mskValorRecebido, TpObrigatorio) Then
      strMsg = strMsg & "Informar o valor recebido válido" & vbCrLf
      Pintar_Controle mskValorRecebido, tpCorContr_Erro
    End If
  End If
  If strMsg = "" Then
    If Len(mskData(0).ClipText) > 0 Then
      If Len(cboFormaPgto.Text) = 0 Then
        strMsg = strMsg & "Selecionar a forma de recebimento" & vbCrLf
        Pintar_Controle cboFormaPgto, tpCorContr_Erro
      End If
      If cboFormaPgto = "CHEQUE" Then
        If Len(cboLivro.Text) = 0 Then
          strMsg = strMsg & "Selecionar o livro" & vbCrLf
          Pintar_Controle cboLivro, tpCorContr_Erro
        End If
        If Len(txtCheque.Text) = 0 Then
          strMsg = strMsg & "Informar o número do cheuqe válido" & vbCrLf
          Pintar_Controle txtCheque, tpCorContr_Erro
        End If
      End If
      If Not Valida_Moeda(mskValorRecebido, TpObrigatorio) Then
        strMsg = strMsg & "Informar o valor recebido válido" & vbCrLf
        Pintar_Controle mskValorRecebido, tpCorContr_Erro
      End If
    Else
      If Not Valida_Moeda(mskValorRecebido, TpNaoObrigatorio) Then
        strMsg = strMsg & "Informar o valor recebido válido" & vbCrLf
        Pintar_Controle mskValorRecebido, tpCorContr_Erro
      End If
    End If
  End If
  '
'  If strTipo = "A" Then
'    If Not Valida_Data(mskData(0), TpObrigatorio) Then
'      strMsg = strMsg & "Informar a data de pagamento válida" & vbCrLf
'      Pintar_Controle mskData(0), tpCorContr_Erro
'    End If
'  End If
'  If Not Valida_Moeda(mskValorRecebido, TpObrigatorio) Then
'    strMsg = strMsg & "Informar o valor pago válido" & vbCrLf
'    Pintar_Controle mskValorRecebido, tpCorContr_Erro
'  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserReceitaInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    If strTipo = "T" Then
      SetarFoco txtDescricao
    Else
      SetarFoco mskGrupo
    End If
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserReceitaInc.Form_Activate]"
End Sub



Private Sub Form_Load()
On Error GoTo trata
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim objReceita      As busSisContas.clsReceita
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 5865
  Me.Width = 10200
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  LerFigurasAvulsas cmdConsultar, "FiltrarDown.ico", "FiltrarDown.ico", "Pesquisar Grupo/Sub Grupo"
  '
  strSql = "SELECT EMPRESACTRTO.DESCRICAO FROM EMPRESACTRTO " & _
                " INNER JOIN EMPRESA ON EMPRESA.PKID = EMPRESACTRTO.EMPRESAID " & _
                " WHERE EMPRESA.PARCEIROID = " & Formata_Dados(glParceiroId, tpDados_Longo) & _
                " ORDER BY EMPRESACTRTO.DESCRICAO;"
  PreencheCombo cboContrato, strSql, False, True
  strSql = "SELECT FORMAPGTO FROM FORMAPGTO ORDER BY FORMAPGTO;"
  PreencheCombo cboFormaPgto, strSql, False, True
  strSql = "SELECT NUMEROLIVRO FROM LIVRO " & _
              " WHERE PARCEIROID = " & Formata_Dados(glParceiroId, tpDados_Longo) & _
              " ORDER BY NUMEROLIVRO;"
  PreencheCombo cboLivro, strSql, False, True
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    LimparCampoMask mskData(1)
    LimparCampoMask mskValorReceber
    LimparCampoTexto txtCheque
    LimparCampoTexto txtCodigo
    LimparCampoMask mskData(0)
    LimparCampoMask mskValorRecebido
    LimparCampoTexto txtDescricao
    'cboFormaPgto.Text = "DINHEIRO"
    LimparCampoMask mskGrupo
    LimparCampoMask mskSubGrupo
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objReceita = New busSisContas.clsReceita
    Set objRs = objReceita.SelecionarReceita(lngRECEITAID)
    '
    If Not objRs.EOF Then
      txtCodigo.Text = objRs.Fields("SEQUENCIAL").Value & ""
      INCLUIR_VALOR_NO_MASK mskData(1), objRs.Fields("DT_VENCIMENTO").Value, TpMaskData
      INCLUIR_VALOR_NO_MASK mskValorReceber, objRs.Fields("VR_RECEBER").Value, TpMaskMoeda
      If objRs.Fields("DESCR_CONTRATO").Value & "" <> "" Then
        cboContrato.Text = objRs.Fields("DESCR_CONTRATO").Value & ""
      End If
      If objRs.Fields("NUMEROLIVRO").Value & "" <> "" Then
        cboLivro.Text = objRs.Fields("NUMEROLIVRO").Value & ""
      End If
      txtCheque.Text = objRs.Fields("NUMEROCHEQUE").Value & ""
      INCLUIR_VALOR_NO_MASK mskData(0), objRs.Fields("DT_RECEBIMENTO").Value, TpMaskData
      INCLUIR_VALOR_NO_MASK mskValorRecebido, objRs.Fields("VR_RECEBIDO").Value, TpMaskMoeda
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
      If objRs.Fields("DESCRFORMAPGTO").Value & "" <> "" Then
        cboFormaPgto.Text = objRs.Fields("DESCRFORMAPGTO").Value & ""
      End If
      INCLUIR_VALOR_NO_MASK mskGrupo, objRs.Fields("CODIGOGRUPODESPESA").Value, TpMaskOutros
      INCLUIR_VALOR_NO_MASK mskSubGrupo, objRs.Fields("CODIGOSUBGRUPODESPESA").Value, TpMaskOutros
    End If
    Set objReceita = Nothing
    If strTipo = "T" Then
      mskData(1).Enabled = False
      mskValorReceber.Enabled = False
      txtCheque.Enabled = False
      mskData(0).Enabled = False
      mskValorRecebido.Enabled = False
      'txtDescricao
      cboFormaPgto.Enabled = False
      mskGrupo.Enabled = False
      mskSubGrupo.Enabled = False
      cmdConsultar.Enabled = False
    End If
  End If
  '
  If Status = tpStatus_Consultar Then
    cmdOk.Visible = False
  End If
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

Private Sub mskData_GotFocus(Index As Integer)
  Selecionar_Conteudo mskData(Index)
End Sub

Private Sub mskData_LostFocus(Index As Integer)
  Pintar_Controle mskData(Index), tpCorContr_Normal
End Sub

Private Sub mskGrupo_GotFocus()
  Selecionar_Conteudo mskGrupo
End Sub

Private Sub mskGrupo_LostFocus()
  Pintar_Controle mskGrupo, tpCorContr_Normal
End Sub

Private Sub mskSubGrupo_GotFocus()
  Selecionar_Conteudo mskSubGrupo
End Sub

Private Sub mskSubGrupo_LostFocus()
  Dim objSubGrupoDespesa  As busSisContas.clsSubGrupoDespesa
  Dim strTipo             As String
  On Error GoTo trata
  Pintar_Controle mskSubGrupo, tpCorContr_Normal
  If Valida_Moeda(mskGrupo, TpObrigatorio) And Valida_Moeda(mskSubGrupo, TpObrigatorio) Then
    'Digitou Grupo e subGrupo, Preencher
    Set objSubGrupoDespesa = New busSisContas.clsSubGrupoDespesa
    strTipo = objSubGrupoDespesa.SelecionarTipoGrupo(mskGrupo.Text, _
                                                     mskSubGrupo.Text)
    If strTipo = "C" Then SetarFoco mskData(1)
    Set objSubGrupoDespesa = Nothing
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub mskValorReceber_GotFocus()
  Selecionar_Conteudo mskValorReceber
End Sub

Private Sub mskValorReceber_LostFocus()
  Pintar_Controle mskValorReceber, tpCorContr_Normal
End Sub

Private Sub mskValorRecebido_GotFocus()
  Selecionar_Conteudo mskValorRecebido
End Sub

Private Sub mskValorRecebido_LostFocus()
  Pintar_Controle mskValorRecebido, tpCorContr_Normal
End Sub

Private Sub txtCheque_GotFocus()
  Selecionar_Conteudo txtCheque
End Sub

Private Sub txtCheque_LostFocus()
  Pintar_Controle txtCheque, tpCorContr_Normal
End Sub

Private Sub txtDescricao_GotFocus()
  Selecionar_Conteudo txtDescricao
End Sub

Private Sub txtDescricao_LostFocus()
  Pintar_Controle txtDescricao, tpCorContr_Normal
End Sub

