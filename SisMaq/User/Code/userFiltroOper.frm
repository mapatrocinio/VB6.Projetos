VERSION 5.00
Begin VB.Form frmUserFiltroOper 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aplicar filtro em cliente para vizualização de cheques"
   ClientHeight    =   3285
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Preencher dados para consulta de unidades"
      Height          =   2115
      Left            =   60
      TabIndex        =   3
      Top             =   90
      Width           =   6555
      Begin VB.TextBox txtSequencial 
         Height          =   285
         Left            =   1350
         TabIndex        =   0
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sequencial"
         Height          =   165
         Left            =   120
         TabIndex        =   4
         Top             =   390
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   880
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2310
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "ENTER"
      Default         =   -1  'True
      Height          =   880
      Left            =   2010
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2310
      Width           =   1215
   End
End
Attribute VB_Name = "frmUserFiltroOper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strSqlWhereTurno   As String
Public strSqlWhereLocacao As String
Option Explicit


Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdConfirmar_Click()
  On Error GoTo trata
  '
  If Not Valida_Moeda(txtSequencial, TpObrigatorio, True) Then
    TratarErroPrevisto "Entrar com o sequencial válido"
    Exit Sub
  End If
  If txtSequencial.Text <> "" Then
    strSqlWhereTurno = "TURNO.PKID IN (SELECT DISTINCT TURNOENTRADAID FROM LOCACAO WHERE SEQUENCIAL = " & Formata_Dados(txtSequencial.Text, tpDados_Longo) & ")"
    strSqlWhereLocacao = "LOCACAO.PKID IN (SELECT DISTINCT LOCACAO.PKID FROM LOCACAO WHERE SEQUENCIAL = " & Formata_Dados(txtSequencial.Text, tpDados_Longo) & ")"
  End If
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Activate()
  SetarFoco txtSequencial
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  CenterForm Me
  LerFiguras Me, tpBmp_Vazio, cmdConfirmar, cmdCancelar
  LimparCampoTexto txtSequencial
  strSqlWhereTurno = ""
  strSqlWhereLocacao = ""
  
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub


Private Sub txtSequencial_LostFocus()
  Seleciona_Conteudo_Controle txtSequencial
End Sub

Private Sub txtSequencial_GotFocus()
  Seleciona_Conteudo_Controle txtSequencial
End Sub

