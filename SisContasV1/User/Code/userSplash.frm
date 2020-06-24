VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Start"
   ClientHeight    =   3735
   ClientLeft      =   1980
   ClientTop       =   360
   ClientWidth     =   6375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3735
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSistema 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Carregando Sistema"
      Enabled         =   0   'False
      Height          =   225
      Left            =   3570
      TabIndex        =   4
      Top             =   900
      Width           =   2010
   End
   Begin VB.CheckBox chkSenha 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Obtendo Senha"
      Enabled         =   0   'False
      Height          =   210
      Left            =   480
      TabIndex        =   3
      Top             =   1350
      Width           =   1995
   End
   Begin VB.CheckBox chkBanco 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lendo Banco de Dados"
      Enabled         =   0   'False
      Height          =   210
      Left            =   480
      TabIndex        =   2
      Top             =   1110
      Width           =   2040
   End
   Begin VB.CheckBox chkConexao 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Carregando Conexão"
      Enabled         =   0   'False
      Height          =   210
      Left            =   480
      TabIndex        =   1
      Top             =   885
      Width           =   2100
   End
   Begin VB.Timer Timer2 
      Left            =   3720
      Top             =   1005
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   6240
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6345
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public Sub LerFiguras()
   '
   Me.Picture = LoadPicture(gsBMPPath & "Xa_2.jpg")
   '
End Sub

Private Sub Form_Load()
  lblTitulo.Caption = TITULOSISTEMA
  chkConexao.Value = 1
  chkBanco.Value = 1
  CapturaParametrosRegistro 0
  chkSenha.Value = 1
  '
  LerFiguras
  '
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMDI.stbPrinc.Panels(1).Text = gsNomeUsu
  frmMDI.stbPrinc.Panels(2).Text = gsNivel
End Sub

Private Sub Timer2_Timer()
  Dim RetVal As Variant
  Unload Me
  frmMDI.Show
  '
  'Captura_Config
  'HabServDesp
  '
  
'''  'Eugenio, para tirar a proteção, comente o código abaixo até antes de End Sub
'''  'Depois vá em Project/References e desmarque as referencias para Protec
'''  '---------------------------------------------------------------
'''  '----------------
'''  'Proteção do sistema
'''  '----------------
'''  Dim clsProtec As clsProtec
'''  Set clsProtec = New clsProtec
'''  '----------------
'''  'Verifica Proteção do sistema
'''  '-------------------------
'''  'Valida primeira vez que entrou no sistema
'''  If Not clsProtec.Valida_Primeira_Vez(gsBDadosPath & nomeBDados, App.Path) Then
'''    End
'''    Exit Sub
'''  End If
'''  'Válida Equipamento
'''  If Not clsProtec.Valida_Estacao(gsBDadosPath & nomeBDados) Then
'''    End
'''    Exit Sub
'''  End If
'''  '----------------
'''  'Valida se sistema expirou
'''  If Not clsProtec.Valida_Chave(gsBDadosPath & nomeBDados, "S", gsNivel) Then
'''    End
'''    Exit Sub
'''  End If
'''  '----------------
'''  'Atualizar data Atual do sistema
'''  clsProtec.Atualiza_Chave_Data_Atual gsBDadosPath & nomeBDados
'''  'Mata o arquivo fisicamene
'''  clsProtec.Trata_Arquivo_Fisico App.Path
'''  Set clsProtec = Nothing
'''  '-----------------
'''  '------------ FIM
'''  '----------------
  
End Sub
