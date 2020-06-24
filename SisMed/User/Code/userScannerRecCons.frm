VERSION 5.00
Begin VB.Form frmUserScannerRecCons 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receita Scanner"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtReceita 
      Height          =   8385
      Left            =   1410
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "userScannerRecCons.frx":0000
      Top             =   60
      Width           =   7005
   End
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   8520
      Left            =   8550
      ScaleHeight     =   8520
      ScaleWidth      =   1860
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   1215
         Left            =   90
         ScaleHeight     =   1155
         ScaleWidth      =   1605
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   7200
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   150
            Width           =   1335
         End
      End
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Recceita"
      Height          =   195
      Index           =   13
      Left            =   90
      TabIndex        =   4
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "frmUserScannerRecCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnRetorno               As Boolean
Public blnFechar                As Boolean
Private blnPrimeiraVez          As Boolean

Public strDescricao             As String


Private Sub cmdCancelar_Click()
  blnFechar = True
  blnRetorno = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco cmdCancelar
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserScannerRecCons.Form_Activate]"
End Sub


Private Sub Form_Load()
  Dim objRegister       As busSisMed.clsRegistro
  On Error GoTo trata
  '
  
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 9000
  Me.Width = 10500
  CenterForm Me
  blnPrimeiraVez = True
  txtReceita.Text = strDescricao
  '
  LerFiguras Me, tpBmp_Vazio, , , cmdCancelar
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
End Sub

