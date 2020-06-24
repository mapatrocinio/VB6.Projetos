VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserDespesaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de despesa"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5145
      Left            =   8250
      ScaleHeight     =   5145
      ScaleWidth      =   1860
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2145
         Left            =   30
         ScaleHeight     =   2085
         ScaleWidth      =   1605
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2790
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1020
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   150
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4815
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da despesa"
      TabPicture(0)   =   "userDespesaInc.frx":0000
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
         Height          =   4215
         Left            =   240
         TabIndex        =   16
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
            Height          =   3525
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   7335
            Begin VB.CheckBox chkPgto 
               Caption         =   "Efetuar mais de um pagamento, ou pagar em formas diferentes"
               Height          =   405
               Left            =   4740
               TabIndex        =   4
               Top             =   630
               Width           =   2535
            End
            Begin VB.ComboBox cboFuncionario 
               Height          =   315
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   2520
               Width           =   4515
            End
            Begin VB.CommandButton cmdConsultar 
               Caption         =   "&Z"
               Height          =   800
               Left            =   3960
               Style           =   1  'Graphical
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   1680
               Width           =   800
            End
            Begin VB.ComboBox cboFormaPgto 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   2040
               Width           =   2175
            End
            Begin VB.CheckBox chkVale 
               Caption         =   "Vale"
               Height          =   195
               Left            =   3720
               TabIndex        =   3
               Top             =   720
               Width           =   855
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               ScaleHeight     =   255
               ScaleWidth      =   3255
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   360
               Width           =   3255
               Begin MSMask.MaskEdBox mskData 
                  Height          =   255
                  Index           =   0
                  Left            =   1440
                  TabIndex        =   1
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  _Version        =   393216
                  BackColor       =   14737632
                  AutoTab         =   -1  'True
                  MaxLength       =   16
                  Mask            =   "##/##/#### ##:##"
                  PromptChar      =   "_"
               End
               Begin VB.Label Da 
                  Caption         =   "Dt. Pgto."
                  Height          =   255
                  Left            =   0
                  TabIndex        =   19
                  Top             =   0
                  Width           =   1335
               End
            End
            Begin VB.TextBox txtDescricao 
               Height          =   525
               Left            =   1560
               MaxLength       =   50
               MultiLine       =   -1  'True
               TabIndex        =   5
               Text            =   "userDespesaInc.frx":001C
               Top             =   1080
               Width           =   5655
            End
            Begin MSMask.MaskEdBox mskValor 
               Height          =   255
               Left            =   1560
               TabIndex        =   2
               Top             =   720
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskGrupo 
               Height          =   255
               Left            =   1560
               TabIndex        =   6
               Top             =   1680
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   2
               Mask            =   "##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskSubGrupo 
               Height          =   255
               Left            =   1920
               TabIndex        =   7
               Top             =   1680
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   2
               Mask            =   "##"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Caption         =   "Funcionário"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   26
               Top             =   2520
               Width           =   1215
            End
            Begin VB.Label Label3 
               Caption         =   "<----------------------"
               Height          =   255
               Left            =   2520
               TabIndex        =   25
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "Forma Pgto."
               Enabled         =   0   'False
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   24
               Top             =   2040
               Width           =   1215
            End
            Begin VB.Label Label2 
               Caption         =   "Grupo/Sub Grupo"
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   1680
               Width           =   1455
            End
            Begin VB.Label Label9 
               Caption         =   "Descrição"
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label8 
               Caption         =   "Valor"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   720
               Width           =   735
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
            TabIndex        =   22
            Top             =   240
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "frmUserDespesaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                         As tpStatus
Public lngDESPESAID                   As Long
Public bRetorno                       As Boolean
Public blnPrimeiraVez                 As Boolean
Public bFechar                        As Boolean
Public strTipo                        As String


Private Sub cboFormaPgto_LostFocus()
  Pintar_Controle cboFormaPgto, tpCorContr_Normal
End Sub

Private Sub cboFuncionario_LostFocus()
  Pintar_Controle cboFuncionario, tpCorContr_Normal
End Sub

Private Sub chkVale_Click()
  On Error Resume Next
  'Tratar campo
  If chkVale.Value = 1 Then
    Label1(1).Enabled = True
    cboFuncionario.Enabled = True
  Else
    Label1(1).Enabled = False
    cboFuncionario.Enabled = False
    cboFuncionario.ListIndex = -1
  End If

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
  frmUserGrupoDespesaCons.QuemChamou = 0
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
  Dim objDespesa              As busSisMaq.clsDespesa
  Dim clsGer                  As busSisMaq.clsGeral
  Dim lngCCId                 As Long
  Dim lngFORMAPGTOID          As Long
  Dim lngSUBGRUPOID           As Long
  Dim lngFUNCIONARIOID        As Long
  Dim strSequencial           As String
  Dim curMargemConsignavel    As Currency
  Dim curValorDespesaJaLanc   As Currency
  Dim curValorDespesa         As Currency
  Dim strDtInicial            As String
  Dim strDtFinal              As String
  Dim datDtInicial            As Date
  Dim datDtFinal              As Date
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração da Despesa
    If Not ValidaCampos Then Exit Sub
    'Valida se unidade de estoque já cadastrada
    Set clsGer = New busSisMaq.clsGeral
    'Obter funcionário
    lngFUNCIONARIOID = 0
    If chkVale.Value = 1 Then 'And blnTrabalhaComFuncAssoc = True Then
      strSql = "Select PKID FROM PESSOA WHERE NOME = " & Formata_Dados(cboFuncionario.Text, tpDados_Texto)
      Set objRs = clsGer.ExecutarSQL(strSql)
      If objRs.EOF Then
        objRs.Close
        Set objRs = Nothing
        Set clsGer = Nothing
        TratarErroPrevisto "Funcionário não cadastrado", "cmdOK_Click"
        Exit Sub
      Else
        lngFUNCIONARIOID = objRs.Fields("PKID").Value
      End If
      objRs.Close
      Set objRs = Nothing
'''      'Validar se funcionário atingiu margem de consignação
'''      strSql = "Select MARGEMCONSIGNAVEL From FUNCIONARIO WHERE PKID = " & Formata_Dados(lngFUNCIONARIOID, tpDados_Longo)
'''      Set objRs = clsGer.ExecutarSQL(strSql)
'''      curMargemConsignavel = 0
'''      If Not objRs.EOF Then
'''        curMargemConsignavel = IIf(IsNull(objRs.Fields("MARGEMCONSIGNAVEL").Value), 0, objRs.Fields("MARGEMCONSIGNAVEL").Value)
'''      End If
'''      objRs.Close
'''      Set objRs = Nothing
'''      datDtInicial = CDate(Year(Now) & "/" & Month(Now) & "/" & glDiaFechaFolha & " 00:00")
'''      datDtFinal = datDtInicial
'''      datDtInicial = DateAdd("D", 1, datDtInicial)
'''      datDtFinal = DateAdd("M", 1, datDtFinal)
'''      '
'''      strDtInicial = Format(datDtInicial, "DD/MM/YYYY")
'''      strDtFinal = Format(datDtFinal, "DD/MM/YYYY")
'''      strDtInicial = strDtInicial & " 00:00"
'''      strDtFinal = strDtFinal & " 23:59"
'''      strSql = "Select SUM(DESPESA.VR_PAGO) AS VALORLANCADO From DESPESA " & _
'''        " WHERE DESPESA.FUNCIONARIOID = " & Formata_Dados(lngFUNCIONARIOID, tpDados_Longo) & _
'''        " AND DESPESA.PKID <> " & Formata_Dados(lngDESPESAID, tpDados_Longo) & _
'''        " AND DT_PAGAMENTO < " & Formata_Dados(strDtFinal, tpDados_DataHora) & _
'''        " AND DT_PAGAMENTO > " & Formata_Dados(strDtInicial, tpDados_DataHora)
'''
'''      Set objRs = clsGer.ExecutarSQL(strSql)
'''      curValorDespesaJaLanc = 0
'''      If Not objRs.EOF Then
'''        curValorDespesaJaLanc = IIf(IsNull(objRs.Fields("VALORLANCADO").Value), 0, objRs.Fields("VALORLANCADO").Value)
'''      End If
'''      objRs.Close
'''      Set objRs = Nothing
'''      'De posse do total já lançado de vales
'''      curValorDespesa = CCur(mskValor.Text)
'''      If (curValorDespesa + curValorDespesaJaLanc) > curMargemConsignavel Then
'''        strMsgErro = "Despesa lançada para funcionário ultrapassa seu valor consignável." & vbCrLf & vbCrLf
'''        strMsgErro = strMsgErro & "Valor Consignável  - " & Format(curMargemConsignavel, "###,##0.00") & vbCrLf
'''        strMsgErro = strMsgErro & "Já lançado este mês - " & Format(curValorDespesaJaLanc, "###,##0.00") & vbCrLf
'''        strMsgErro = strMsgErro & "Despesa + lançado - " & Format(curValorDespesaJaLanc + curValorDespesa, "###,##0.00") & vbCrLf
'''        strMsgErro = strMsgErro & "Máximo que ainda pode ser lançado - " & Format(curMargemConsignavel - curValorDespesaJaLanc, "###,##0.00") & vbCrLf
'''        TratarErroPrevisto strMsgErro, "cmdOK_Click"
'''        Exit Sub
'''      End If
    End If
    
    strSql = "Select PKID From FORMAPGTO WHERE FORMAPGTO = " & Formata_Dados(cboFormaPgto.Text, tpDados_Texto, tpNulo_NaoAceita)
    Set objRs = clsGer.ExecutarSQL(strSql)
    If objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set clsGer = Nothing
      TratarErroPrevisto "Forma de Pagamento não cadastrada", "cmdOK_Click"
      Exit Sub
      
    Else
      lngFORMAPGTOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    lngSUBGRUPOID = 0
    If mskGrupo.ClipText <> "" Or mskSubGrupo.ClipText <> "" Then
      strSql = "Select SUBGRUPODESPESA.PKID From GRUPODESPESA INNER JOIN SUBGRUPODESPESA ON GRUPODESPESA.PKID = SUBGRUPODESPESA.GRUPODESPESAID " & _
        "WHERE GRUPODESPESA.CODIGO = " & Formata_Dados(mskGrupo.Text, tpDados_Texto, tpNulo_NaoAceita) & _
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
    End If
    '
    Set clsGer = Nothing
'''    'PEDE SENHA SUPERIOR
'''    '----------------------------
'''    '----------------------------
'''    'Pede Senha Superior (Diretor, Gerente ou Administrador)
'''    If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
'''      'Só pede senha superior se quem estiver logado não for superior
'''      frmUserLoginSup.Show vbModal
'''
'''      If Len(Trim(gsNomeUsuLib)) = 0 Then
'''        TratarErroPrevisto "É necessário a confirmação com senha superior para incluir ou alterar uma despesa.", "[frmUserDespesaInc.cmdOk_Click]"
'''        Exit Sub
'''      End If
'''      '
'''      'Capturou Nome do Usuário, continua processo
'''    Else
'''      gsNomeUsuLib = gsNomeUsu
'''    End If
'''    '--------------------------------
'''    '--------------------------------
    
    Set objDespesa = New busSisMaq.clsDespesa
    
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objDespesa.AlterarDespesa lngDESPESAID, _
                                txtDescricao.Text, _
                                mskValor, _
                                IIf(chkVale.Value = 1, "S", "N"), _
                                lngSUBGRUPOID, _
                                lngFORMAPGTOID, _
                                gsNomeUsuLib, _
                                lngFUNCIONARIOID

      bRetorno = True
      bFechar = True
      'IMP_COMPROV_DESPESA lngDESPESAID, gsNomeEmpresa, 1, RetornaDescTurnoCorrente, blnTrabalhaComFuncAssoc
      Unload Me
      Exit Sub
    ElseIf Status = tpStatus_Incluir Then
      'Código para inclusão
      '
      strSequencial = RetornaGravaCampoSequencial("SEQUENCIALDESP")
      '
      lngDESPESAID = objDespesa.IncluirDespesa(strSequencial, _
                                               strTipo, _
                                               RetornaCodTurnoCorrente, _
                                               Format(Now, "DD/MM/YYYY hh:mm"), _
                                               txtDescricao.Text, _
                                               mskValor.Text, _
                                               IIf(chkVale.Value = 1, "S", "N"), _
                                               lngSUBGRUPOID, _
                                               lngFORMAPGTOID, _
                                               gsNomeUsu, _
                                               gsNomeUsuLib, _
                                               lngFUNCIONARIOID)
      '
      bRetorno = True
      INCLUIR_VALOR_NO_MASK mskValor, "", TpMaskMoeda
      chkVale.Value = 0
      chkPgto.Value = 0
      txtDescricao.Text = ""
      mskGrupo.Text = "__"
      mskSubGrupo.Text = "__"
      'cboFormaPgto.ListIndex = 0
    End If
    Set objDespesa = Nothing
    'IMPRIMIR BOLETA DE VENDAS
    'IMP_COMPROV_DESPESA lngDESPESAID, gsNomeEmpresa, 1, RetornaDescTurnoCorrente, blnTrabalhaComFuncAssoc
    SetarFoco mskValor
    If Status = tpStatus_Incluir Then lngDESPESAID = 0
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  Dim strMsg     As String
  '
  If strTipo = "A" Then
    If Not Valida_Data(mskData(0), TpObrigatorio) Then
      strMsg = strMsg & "Informar a data de pagamento válida" & vbCrLf
      Pintar_Controle mskData(0), tpCorContr_Erro
    End If
  End If
  If Not Valida_Moeda(mskValor, TpObrigatorio) Then
    strMsg = strMsg & "Informar o valor pago válido" & vbCrLf
    Pintar_Controle mskValor, tpCorContr_Erro
  End If
  If Len(txtDescricao.Text) = 0 Then
    strMsg = strMsg & "Informar a descrição da despesa válida" & vbCrLf
    Pintar_Controle txtDescricao, tpCorContr_Erro
  End If
  If Len(cboFormaPgto.Text) = 0 Then
    strMsg = strMsg & "Selecionar a forma de pagamento" & vbCrLf
    Pintar_Controle cboFormaPgto, tpCorContr_Erro
  End If
  '
  If chkVale.Value = False Then
    If Not Valida_Moeda(mskGrupo, TpObrigatorio) Then
      strMsg = strMsg & "Informar o Grupo válido" & vbCrLf
      Pintar_Controle mskGrupo, tpCorContr_Erro
    End If
    If Not Valida_Moeda(mskSubGrupo, TpObrigatorio) Then
      strMsg = strMsg & "Informar o Sub Grupo válido" & vbCrLf
      Pintar_Controle mskSubGrupo, tpCorContr_Erro
    End If
  Else
    'Verifica se selecionou funcionário
    If Len(cboFuncionario.Text) = 0 Then
      strMsg = strMsg & "Selecionar o funcionário" & vbCrLf
      Pintar_Controle cboFuncionario, tpCorContr_Erro
    End If
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserDespesaInc.ValidaCampos]"
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
      SetarFoco mskValor
    Else
      SetarFoco mskData(0)
    End If
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserDespesaInc.Form_Activate]"
End Sub



Private Sub Form_Load()
On Error GoTo trata
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim objDespesa      As busSisMaq.clsDespesa
  '
  
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 5520
  Me.Width = 10200
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  LerFigurasAvulsas cmdConsultar, "Filtrar.ico", "FiltrarDown.ico", "Pesquisar Grupo/Sub Grupo"
  '
  strSql = "SELECT FORMAPGTO FROM FORMAPGTO ORDER BY FORMAPGTO;"
  PreencheCombo cboFormaPgto, strSql, False, True
  strSql = "SELECT NOME FROM PESSOA " & _
      " INNER JOIN FUNCIONARIO ON FUNCIONARIO.PESSOAID = PESSOA.PKID " & _
      " ORDER BY NOME;"
  PreencheCombo cboFuncionario, strSql, False, True
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    LimparCampoTexto txtCodigo
    LimparCampoMask mskData(0)
    LimparCampoMask mskValor
    LimparCampoTexto txtDescricao
    cboFormaPgto.Text = "DINHEIRO"
    LimparCampoMask mskGrupo
    LimparCampoMask mskSubGrupo
    LimparCampoCheck chkVale
    LimparCampoCheck chkPgto
    '
  ElseIf Status = tpStatus_Alterar Then
    'Pega Dados do Banco de dados
    Set objDespesa = New busSisMaq.clsDespesa
    Set objRs = objDespesa.SelecionarDespesa(lngDESPESAID)
    '
    If Not objRs.EOF Then
      txtCodigo.Text = objRs.Fields("SEQUENCIAL").Value & ""
      If objRs.Fields("VALE").Value & "" = "S" Then
        chkVale.Value = 1
      Else
        chkVale.Value = 0
      End If
      If objRs.Fields("TOT_PGTO_OUTRA_FORMA").Value = 0 Then
        chkPgto.Value = 0
      Else
        chkPgto.Value = 1
      End If
      INCLUIR_VALOR_NO_MASK mskData(0), objRs.Fields("DT_PAGAMENTO").Value, TpMaskData
      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VR_PAGO").Value, TpMaskMoeda
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
      cboFormaPgto.Text = objRs.Fields("DESCRFORMAPGTO").Value & ""
      INCLUIR_VALOR_NO_MASK mskGrupo, objRs.Fields("CODIGOGRUPODESPESA").Value, TpMaskOutros
      INCLUIR_VALOR_NO_MASK mskSubGrupo, objRs.Fields("CODIGOSUBGRUPODESPESA").Value, TpMaskOutros
      'Tratar campo
      chkVale_Click
      If objRs.Fields("NOME_FUNCIONARIO").Value & "" <> "" Then
        cboFuncionario.Text = objRs.Fields("NOME_FUNCIONARIO").Value & ""
      End If
    End If
    Set objDespesa = Nothing
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
  If Not bFechar Then Cancel = True
End Sub

Private Sub mskData_GotFocus(Index As Integer)
  Selecionar_Conteudo mskData(0)
End Sub

Private Sub mskData_LostFocus(Index As Integer)
  Pintar_Controle mskData(0), tpCorContr_Normal
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
  Pintar_Controle mskSubGrupo, tpCorContr_Normal
End Sub

Private Sub mskValor_GotFocus()
  Selecionar_Conteudo mskValor
End Sub

Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub

Private Sub txtDescricao_GotFocus()
  Selecionar_Conteudo txtDescricao
End Sub

Private Sub txtDescricao_LostFocus()
  Pintar_Controle txtDescricao, tpCorContr_Normal
End Sub

