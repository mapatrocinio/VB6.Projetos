VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserCartaoDebitoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de cart�o de d�bito"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   8520
      ScaleHeight     =   2565
      ScaleWidth      =   1860
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2055
         Left            =   0
         ScaleHeight     =   1995
         ScaleWidth      =   1605
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   360
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   4048
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do cart�o de d�bito"
      TabPicture(0)   =   "userCartaoDebitoInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picTrava(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.PictureBox picTrava 
         BorderStyle     =   0  'None
         Height          =   1665
         Index           =   0
         Left            =   120
         ScaleHeight     =   1665
         ScaleWidth      =   7695
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   480
         Width           =   7695
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
            Height          =   1635
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   7695
            Begin VB.PictureBox Picture3 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1620
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   1260
               Width           =   2235
               Begin VB.OptionButton optConvenio 
                  Caption         =   "Sim"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   3
                  Top             =   0
                  Width           =   1065
               End
               Begin VB.OptionButton optConvenio 
                  Caption         =   "N�o"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   4
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1095
               End
            End
            Begin VB.TextBox txtNumero 
               Height          =   285
               Left            =   1590
               MaxLength       =   3
               TabIndex        =   0
               Top             =   270
               Width           =   885
            End
            Begin VB.TextBox txtNome 
               Height          =   285
               Left            =   1590
               MaxLength       =   50
               TabIndex        =   1
               Top             =   600
               Width           =   5895
            End
            Begin MSMask.MaskEdBox mskPercDesc 
               Height          =   255
               Left            =   1590
               TabIndex        =   2
               Top             =   960
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   6
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Conv�nio"
               Height          =   315
               Index           =   12
               Left            =   180
               TabIndex        =   16
               Top             =   1290
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "N�mero"
               Height          =   255
               Index           =   0
               Left            =   150
               TabIndex        =   14
               Top             =   270
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Perc. tx. Admin."
               Height          =   255
               Index           =   1
               Left            =   180
               TabIndex        =   13
               Top             =   930
               Width           =   1305
            End
            Begin VB.Label Label6 
               Caption         =   "Nome"
               Height          =   255
               Index           =   3
               Left            =   180
               TabIndex        =   12
               Top             =   600
               Width           =   1455
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserCartaoDebitoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                     As tpStatus
Public lngCARTAOID                As Long
Public bRetorno                   As Boolean
Public bFechar                    As Boolean
Public sTitulo                    As String
Public intQuemChamou              As Integer
Private blnPrimeiraVez            As Boolean



Private Sub cmdCancelar_Click()
  bFechar = True
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim strMsgErro              As String
  Dim objRs                   As ADODB.Recordset
  Dim objCartaoDebito         As busSisMed.clsCartaoDebito
  Dim objGer                  As busSisMed.clsGeral
  Dim strConvenio             As String
  
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclus�o/Altera��o de Grupo card�pio
    If Not ValidaCampos Then Exit Sub
    'Valida se cart�o j� cadastrado
    Set objGer = New busSisMed.clsGeral
    strSql = "Select * From CARTAO WHERE NOME = " & Formata_Dados(txtNome.Text, tpDados_Texto, tpNulo_Aceita) & _
      " AND PKID <> " & Formata_Dados(lngCARTAOID, tpDados_Longo, tpNulo_NaoAceita)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGer = Nothing
      TratarErroPrevisto "Cart�o j� cadastrado", "cmdOK_Click"
      Pintar_Controle txtNome, tpCorContr_Erro
      SetarFoco txtNome
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGer = Nothing
    '
    Set objCartaoDebito = New busSisMed.clsCartaoDebito
    '
    If optConvenio(0).Value Then
      strConvenio = "S"
    ElseIf optConvenio(1).Value Then
      strConvenio = "N"
    Else
      strConvenio = ""
    End If
    If Status = tpStatus_Alterar Then
      'C�digo para altera��o
      '
      '
      objCartaoDebito.AlterarCartaoDebito lngCARTAOID, _
                              txtNome.Text, _
                              txtNumero.Text, _
                              mskPercDesc, _
                              strConvenio
                            
      bRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Pega Informa��es para inserir
      '
      objCartaoDebito.InserirCartaoDebito txtNome.Text, _
                              txtNumero.Text, _
                              mskPercDesc, _
                              strConvenio
      '
      bRetorno = True
    End If
    Set objCartaoDebito = Nothing
    bFechar = True
    Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg     As String
  '
  If Not IsNumeric(txtNumero.Text) Then
    strMsg = strMsg & "Informar o n�mero do cart�o v�lido" & vbCrLf
    Pintar_Controle txtNumero, tpCorContr_Erro
    SetarFoco txtNumero
  End If
  If Len(Trim(txtNumero.Text)) <> 3 Then
    strMsg = strMsg & "Informar o n�mero do cart�o cm tr�s digitos" & vbCrLf
    Pintar_Controle txtNumero, tpCorContr_Erro
    SetarFoco txtNumero
  End If
  If Len(txtNome.Text) = 0 Then
    strMsg = strMsg & "Informar o nome do Cart�o" & vbCrLf
    Pintar_Controle txtNome, tpCorContr_Erro
    SetarFoco txtNome
  End If
  If Not Valida_Moeda(mskPercDesc, TpObrigatorio) Then
    strMsg = strMsg & "Informar a taxa do cart�o v�lida" & vbCrLf
    Pintar_Controle mskPercDesc, tpCorContr_Erro
    SetarFoco mskPercDesc
  End If
  If strMsg = "" Then
    If CCur(mskPercDesc) < 0 Or CCur(mskPercDesc) > 100 Then
      strMsg = strMsg & "Preencher a taxa de administra��o do Cart�o com valor de 0 a 100 %" & vbCrLf
      Pintar_Controle mskPercDesc, tpCorContr_Erro
      SetarFoco mskPercDesc
    End If
  End If
  If Not Valida_Option(optConvenio, False) Then
    strMsg = strMsg & "Slecionar conv�nio" & vbCrLf
    SetarFoco optConvenio(0)
  End If
  
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserCartaoDebitoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    tabDetalhes.Tab = 0
    SetarFoco txtNumero
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserCartaoDebitoInc.Form_Activate]"
End Sub

Private Sub txtNome_GotFocus()
  Selecionar_Conteudo txtNome
End Sub

Private Sub txtNome_LostFocus()
  Pintar_Controle txtNome, tpCorContr_Normal
End Sub

Private Sub txtNumero_GotFocus()
  Selecionar_Conteudo txtNumero
End Sub

Private Sub txtNumero_LostFocus()
  Pintar_Controle txtNumero, tpCorContr_Normal
End Sub

Private Sub mskPercDesc_GotFocus()
  Selecionar_Conteudo mskPercDesc
End Sub

Private Sub mskPercDesc_LostFocus()
  Pintar_Controle mskPercDesc, tpCorContr_Normal
End Sub


Private Sub Form_Load()
On Error GoTo trata
  Dim objRs     As ADODB.Recordset
  Dim strSql    As String
  Dim objCartaoDebito As busSisMed.clsCartaoDebito
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 2940
  Me.Width = 10470
  CenterForm Me
  blnPrimeiraVez = True
  optConvenio(0).Value = False
  optConvenio(1).Value = False
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclus�o, Inclui o Pedido
    txtNome.Text = ""
    txtNumero.Text = ""
    INCLUIR_VALOR_NO_MASK mskPercDesc, "", TpMaskMoeda
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objCartaoDebito = New busSisMed.clsCartaoDebito
    Set objRs = objCartaoDebito.ListarCartaoDebito(lngCARTAOID)
    '
    If Not objRs.EOF Then
      txtNome.Text = objRs.Fields("NOME").Value & ""
      txtNumero.Text = objRs.Fields("NUMERO").Value & ""
      '
      INCLUIR_VALOR_NO_MASK mskPercDesc, objRs.Fields("PERCTAXAADMIN").Value, TpMaskMoeda
      If objRs.Fields("INDCONVENIO").Value & "" = "S" Then
        optConvenio(0).Value = True
        optConvenio(1).Value = False
      ElseIf objRs.Fields("INDCONVENIO").Value & "" = "N" Then
        optConvenio(0).Value = False
        optConvenio(1).Value = True
      Else
        optConvenio(0).Value = False
        optConvenio(1).Value = False
      End If
      
    End If
    Set objCartaoDebito = Nothing
  End If
  
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub



Private Sub Form_Unload(Cancel As Integer)
  If Not bFechar Then Cancel = True
End Sub
