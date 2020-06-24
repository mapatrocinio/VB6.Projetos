VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserMovCaixaInc 
   Caption         =   "Movimentação do Caixa"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   11265
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   11265
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4965
      Width           =   11265
      Begin VB.CommandButton cmdOk 
         Caption         =   "ENTER"
         Default         =   -1  'True
         Height          =   880
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   880
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4725
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8334
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   529
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Totais"
      TabPicture(0)   =   "userMovCaixaInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraAluno"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraAluno 
         Height          =   4245
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   10695
         Begin VB.Frame Frame1 
            Height          =   2265
            Left            =   120
            TabIndex        =   23
            Top             =   990
            Width           =   10335
            Begin MSMask.MaskEdBox mskRetDin 
               Height          =   255
               Left            =   1680
               TabIndex        =   1
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskRetChq 
               Height          =   255
               Left            =   5040
               TabIndex        =   2
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskRetCar 
               Height          =   255
               Left            =   8430
               TabIndex        =   3
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskRetPen 
               Height          =   255
               Left            =   5040
               TabIndex        =   5
               Top             =   540
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDepDin 
               Height          =   255
               Left            =   1680
               TabIndex        =   7
               Top             =   840
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDepChq 
               Height          =   255
               Left            =   5040
               TabIndex        =   8
               Top             =   840
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDepCar 
               Height          =   255
               Left            =   8430
               TabIndex        =   9
               Top             =   840
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDepPen 
               Height          =   255
               Left            =   5040
               TabIndex        =   11
               Top             =   1140
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskRetCarDeb 
               Height          =   255
               Left            =   1680
               TabIndex        =   4
               Top             =   540
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDepCarDeb 
               Height          =   255
               Left            =   1680
               TabIndex        =   10
               Top             =   1140
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskRetFat 
               Height          =   255
               Left            =   8430
               TabIndex        =   6
               Top             =   540
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDepFat 
               Height          =   255
               Left            =   8430
               TabIndex        =   12
               Top             =   1140
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label7 
               Caption         =   "Deposito Penhor"
               Height          =   255
               Left            =   6840
               TabIndex        =   39
               Top             =   1110
               Width           =   1335
            End
            Begin VB.Label Label6 
               Caption         =   "Retirada Fatura"
               Height          =   255
               Left            =   6840
               TabIndex        =   38
               Top             =   540
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Deposito Cartão Déb"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   1140
               Width           =   1575
            End
            Begin VB.Label Label4 
               Caption         =   "Retirada Cartão Déb"
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   540
               Width           =   1455
            End
            Begin VB.Label Label10 
               Caption         =   "Retirada Dinheiro"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label11 
               Caption         =   "Retirada Cheque"
               Height          =   255
               Left            =   3480
               TabIndex        =   30
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label12 
               Caption         =   "Retirada Cartão Cred"
               Height          =   255
               Left            =   6840
               TabIndex        =   29
               Top             =   240
               Width           =   1485
            End
            Begin VB.Label Label13 
               Caption         =   "Deposito Dinheiro"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   840
               Width           =   1455
            End
            Begin VB.Label Label14 
               Caption         =   "Deposito Cheque"
               Height          =   255
               Left            =   3480
               TabIndex        =   27
               Top             =   810
               Width           =   1455
            End
            Begin VB.Label Label15 
               Caption         =   "Deposito Cartão Créd"
               Height          =   255
               Left            =   6840
               TabIndex        =   26
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label Label17 
               Caption         =   "Retirada Penhor"
               Height          =   255
               Left            =   3480
               TabIndex        =   25
               Top             =   540
               Width           =   1215
            End
            Begin VB.Label Label18 
               Caption         =   "Deposito Penhor"
               Height          =   255
               Left            =   3480
               TabIndex        =   24
               Top             =   1110
               Width           =   1335
            End
         End
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   735
            Index           =   0
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   10455
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   240
            Width           =   10455
            Begin VB.TextBox txtFuncionario 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtFuncionario"
               Top             =   0
               Width           =   3585
            End
            Begin VB.TextBox txtData 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7440
               Locked          =   -1  'True
               TabIndex        =   34
               TabStop         =   0   'False
               Text            =   "txtData"
               Top             =   360
               Width           =   1935
            End
            Begin VB.TextBox txtUsuario 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4440
               Locked          =   -1  'True
               TabIndex        =   32
               TabStop         =   0   'False
               Text            =   "txtUsuario"
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txtValor 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   21
               TabStop         =   0   'False
               Text            =   "txtValor"
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txtTurno 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   15
               TabStop         =   0   'False
               Text            =   "txtTurno"
               Top             =   0
               Width           =   4245
            End
            Begin VB.Label Label3 
               Caption         =   "Data Mov."
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
               Left            =   6000
               TabIndex        =   35
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label2 
               Caption         =   "Usuário Resp."
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
               Left            =   3000
               TabIndex        =   33
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Valor do Caixa"
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
               Left            =   0
               TabIndex        =   22
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label21 
               Caption         =   "Turno Corrente"
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
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Width           =   1455
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserMovCaixaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public vrTotCaixa             As Currency
Public vrTotCaixaDin          As Currency
Public vrTotCaixaPen          As Currency
Public vrTotCaixaFat          As Currency
Public vrTotCaixaChq          As Currency
Public vrTotCaixaCar          As Currency
Public vrTotCaixaCarDeb       As Currency

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim sSql As String
  Dim objGeral As busSisMed.clsGeral
  
  Dim vrRetDin      As Currency
  Dim vrRetChq      As Currency
  Dim vrRetCar      As Currency
  Dim vrRetCarDeb   As Currency
  Dim vrRetPen      As Currency
  Dim vrRetFat      As Currency
  
  Dim vrDepDin        As Currency
  Dim vrDepChq        As Currency
  Dim vrDepCar        As Currency
  Dim vrDepCarDeb     As Currency
  Dim vrDepPen        As Currency
  Dim vrDepFat        As Currency
  '
  Dim vrTotMovCaixa As Currency
  '
  If Not ValidaCampos Then Exit Sub
  '
  If Len(mskRetDin.ClipText) <> 0 Then vrRetDin = CCur(mskRetDin.Text)
  If Len(mskRetChq.ClipText) <> 0 Then vrRetChq = CCur(mskRetChq.Text)
  If Len(mskRetCar.ClipText) <> 0 Then vrRetCar = CCur(mskRetCar.Text)
  If Len(mskRetCarDeb.ClipText) <> 0 Then vrRetCarDeb = CCur(mskRetCarDeb.Text)
  If Len(mskRetPen.ClipText) <> 0 Then vrRetPen = CCur(mskRetPen.Text)
  If Len(mskRetFat.ClipText) <> 0 Then vrRetFat = CCur(mskRetFat.Text)
  '
  If Len(mskDepDin.ClipText) <> 0 Then vrDepDin = CCur(mskDepDin.Text)
  If Len(mskDepChq.ClipText) <> 0 Then vrDepChq = CCur(mskDepChq.Text)
  If Len(mskDepCar.ClipText) <> 0 Then vrDepCar = CCur(mskDepCar.Text)
  If Len(mskDepCarDeb.ClipText) <> 0 Then vrDepCarDeb = CCur(mskDepCarDeb.Text)
  If Len(mskDepPen.ClipText) <> 0 Then vrDepPen = CCur(mskDepPen.Text)
  If Len(mskDepFat.ClipText) <> 0 Then vrDepFat = CCur(mskDepFat.Text)
  '
  vrTotMovCaixa = vrDepDin + vrDepChq + vrDepCar + vrDepCarDeb + vrDepPen + vrDepFat
  vrTotMovCaixa = vrTotMovCaixa - vrRetDin - vrRetChq - vrRetCar - vrRetCarDeb - vrRetPen - vrRetFat
    
  If (vrTotCaixa + vrTotMovCaixa) < 0 Then
    MsgBox "O Valor Total da Movimentação do Caixa não pode ultrapassar o valor total no caixa"
    Exit Sub
  End If
  If (vrTotCaixaDin + vrDepDin - vrRetDin) < 0 Then
    MsgBox "O Valor Total da Movimentação do Caixa em dinheiro não pode ultrapassar o valor total de dinheiro no caixa"
    Exit Sub
  End If
  If (vrTotCaixaChq + vrDepChq - vrRetChq) < 0 Then
    MsgBox "O Valor Total da Movimentação do Caixa em Cheque não pode ultrapassar o valor total de cheque no caixa"
    Exit Sub
  End If
  If (vrTotCaixaCar + vrDepCar - vrRetCar) < 0 Then
    MsgBox "O Valor Total da Movimentação do Caixa em Cartão não pode ultrapassar o valor total de Cartão no caixa"
    Exit Sub
  End If
  If (vrTotCaixaCarDeb + vrDepCarDeb - vrRetCarDeb) < 0 Then
    MsgBox "O Valor Total da Movimentação do Caixa em Cartão de débito não pode ultrapassar o valor total de Cartão de débito no caixa"
    Exit Sub
  End If
  If (vrTotCaixaPen + vrDepPen - vrRetPen) < 0 Then
    MsgBox "O Valor Total da Movimentação do Caixa em Penhor não pode ultrapassar o valor total em Penhor no caixa"
    Exit Sub
  End If
  If (vrTotCaixaFat + vrDepFat - vrRetFat) < 0 Then
    MsgBox "O Valor Total da Movimentação do Caixa em Fatura não pode ultrapassar o valor total em Fatura no caixa"
    Exit Sub
  End If
  
  'Se está tudo ok, Grava no Banco de dados
  Set objGeral = New busSisMed.clsGeral
  sSql = "INSERT INTO SANGRIA (TURNOID, DATA, RESPONSAVEL, VRRETDIN, VRRETCHQ, VRRETCAR, VRRETCARDEB, VRRETPEN, VRDEPDIN, VRDEPCHQ, VRDEPCAR, VRDEPCARDEB, VRDEPPEN, VRRETFAT, VRDEPFAT) VALUES (" & _
    Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo) & _
    ", " & Formata_Dados(Format(Now, "DD/MM/YYYY hh:mm"), tpDados_DataHora) & _
    ", " & Formata_Dados(gsNomeUsuLib, tpDados_Texto, tpNulo_NaoAceita, 30) & _
    ", " & Formata_Dados(mskRetDin, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(mskRetChq, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(mskRetCar.Text, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(mskRetCarDeb.Text, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(mskRetPen.Text, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(mskDepDin.Text, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(mskDepChq.Text, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(mskDepCar.Text, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(mskDepCarDeb.Text, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(mskDepPen.Text, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(mskRetFat.Text, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(mskDepFat.Text, tpDados_Moeda, tpNulo_Aceita) & _
    ")"
  objGeral.ExecutarSQLAtualizacao sSql
  Set objGeral = Nothing
  '
  IMP_COMPROV_SANGRIA gsNomeEmpresa, 1, gsNomeUsu
  'LOG

  Unload Me
  '
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub


Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim Msg As String
  '
  If Not Valida_Moeda(mskRetDin, TpNaoObrigatorio, , False) Then
    Msg = "Informar o valor da Retirada em  Dinheiro válido"
  ElseIf Not Valida_Moeda(mskRetChq, TpNaoObrigatorio, , False) Then
    Msg = "Informar o valor da Retirada em Cheque válido"
  ElseIf Not Valida_Moeda(mskRetCar, TpNaoObrigatorio, , False) Then
    Msg = "Informar o valor da Retirada em Cartão de crédito válido"
  ElseIf Not Valida_Moeda(mskRetCarDeb, TpNaoObrigatorio, , False) Then
    Msg = "Informar o valor da Retirada em Cartão de débito válido"
  ElseIf Not Valida_Moeda(mskRetPen, TpNaoObrigatorio, , False) Then
    Msg = "Informar o valor da Retirada em Penhor válido"
    
  ElseIf Not Valida_Moeda(mskDepDin, TpNaoObrigatorio, , False) Then
    Msg = "Informar o valor do Depósito em Dinheiro válido"
  ElseIf Not Valida_Moeda(mskDepChq, TpNaoObrigatorio, , False) Then
    Msg = "Informar o valor do Depósito em Cheque válido"
  ElseIf Not Valida_Moeda(mskDepCar, TpNaoObrigatorio, , False) Then
    Msg = "Informar o valor do Depósito em Cartão de crédito válido"
  ElseIf Not Valida_Moeda(mskDepCarDeb, TpNaoObrigatorio, , False) Then
    Msg = "Informar o valor do Depósito em Cartão de débito válido"
  ElseIf Not Valida_Moeda(mskDepPen, TpNaoObrigatorio, , False) Then
    Msg = "Informar o valor do Depósito em Penhor válido"
  ElseIf Len(mskRetDin.ClipText) = 0 And Len(mskRetChq.ClipText) = 0 And Len(mskRetCar.ClipText) = 0 And Len(mskRetCarDeb.ClipText) = 0 And Len(mskRetPen.ClipText) = 0 And Len(mskRetFat.ClipText) = 0 And _
    Len(mskDepDin.ClipText) = 0 And Len(mskDepChq.ClipText) = 0 And Len(mskDepCar.ClipText) = 0 And Len(mskDepCarDeb.ClipText) = 0 And Len(mskDepPen.ClipText) = 0 And Len(mskDepFat.ClipText) = 0 Then
    Msg = "Deverá ser informado um depósito ou uma retirada válida"
    
  End If
  If Len(Msg) <> 0 Then
    MsgBox "Os seguintes erros ocorreram: " & vbCrLf & vbCrLf & Msg, vbExclamation, TITULOSISTEMA
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  ValidaCampos = False
End Function


Private Sub Form_Load()
  On Error GoTo trata
  AmpS
  'Me.Caption = Me.Caption & sgFormCaption
  Me.Width = 11385
  Me.Height = 6405
  CenterForm Me
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, cmdCancelar
  '
  txtTurno.Text = RetornaDescTurnoCorrente
  txtFuncionario.Text = gsNomeUsuCompleto
  txtValor.Text = Format(vrTotCaixa, "###,###,##0.00")
  txtUsuario.Text = gsNomeUsuLib
  txtData.Text = DataHoraAtualFormatada
  '
  If Not gbTrabComDepSangria Then
    'Não Trabalha com depósito na Sangria
    Label13.Enabled = False
    Label14.Enabled = False
    Label15.Enabled = False
    Label5.Enabled = False
    Label18.Enabled = False
    Label7.Enabled = False
    mskDepDin.Enabled = False
    mskDepChq.Enabled = False
    mskDepCar.Enabled = False
    mskDepCarDeb.Enabled = False
    mskDepPen.Enabled = False
    mskDepFat.Enabled = False
  End If
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub
