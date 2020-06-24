VERSION 5.00
Begin VB.Form frmUserScannerCons 
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
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   8520
      Left            =   8550
      ScaleHeight     =   8520
      ScaleWidth      =   1860
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.CommandButton cmdMaisVertical 
         Caption         =   "+"
         Height          =   405
         Left            =   540
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Aumentar botões Unidades Verticalmente"
         Top             =   1140
         Width           =   400
      End
      Begin VB.CommandButton cmdMenosVertical 
         Caption         =   "-"
         Height          =   405
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Diminuir botões Unidades Verticalmente"
         Top             =   1140
         Width           =   400
      End
      Begin VB.CommandButton cmdMenosHorizontal 
         Caption         =   "-"
         Height          =   405
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Diminuir botões Unidades Horizontalmente"
         Top             =   360
         Width           =   400
      End
      Begin VB.CommandButton cmdMaisHorizontal 
         Caption         =   "+"
         Height          =   405
         Left            =   510
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Aumentar botões Unidades Horizontalmente"
         Top             =   360
         Width           =   400
      End
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   1215
         Left            =   90
         ScaleHeight     =   1155
         ScaleWidth      =   1605
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   7200
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   150
            Width           =   1335
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Vertical"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Horizontal"
         Height          =   195
         Index           =   29
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Image Image1 
      Height          =   8445
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "frmUserScannerCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnRetorno               As Boolean
Public blnFechar                As Boolean
Private blnPrimeiraVez          As Boolean

'BOTÕES NOVOS
Public lngTotalBotoes           As Long
Public intLarguraPadrao         As Integer
Public intAlturaPadrao          As Integer


Private Sub cmdCancelar_Click()
  blnFechar = True
  blnRetorno = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub



Private Sub RedimensionaBotoes()
  On Error GoTo trata
  Image1.Width = intLarguraPadrao
  Image1.Height = intAlturaPadrao
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source, _
            Err.Description
End Sub
Private Sub cmdMaisHorizontal_Click()
  On Error GoTo trata
  Dim objRegister As busSisMed.clsRegistro
  'If intLarguraPadrao > 3000 Then Exit Sub
  intLarguraPadrao = intLarguraPadrao + 50
  cmdMaisHorizontal.ToolTipText = intLarguraPadrao
  cmdMenosHorizontal.ToolTipText = intLarguraPadrao
  RedimensionaBotoes
  'Saving Settings
  Set objRegister = New busSisMed.clsRegistro
  objRegister.SalvarChaveRegistro TITULOSISTEMA, _
                                  "WidthButtonScanner", _
                                  intLarguraPadrao & ""
  Set objRegister = Nothing

  
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source, _
            Err.Description
End Sub

Private Sub cmdMaisVertical_Click()
  On Error GoTo trata
  Dim objRegister As busSisMed.clsRegistro
  'If intAlturaPadrao > 3000 Then Exit Sub
  intAlturaPadrao = intAlturaPadrao + 50
  cmdMaisVertical.ToolTipText = intAlturaPadrao
  cmdMenosVertical.ToolTipText = intAlturaPadrao
  RedimensionaBotoes
  
  Set objRegister = New busSisMed.clsRegistro
  objRegister.SalvarChaveRegistro TITULOSISTEMA, _
                                  "HeightButtonScanner", _
                                  intAlturaPadrao & ""
  Set objRegister = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source, _
            Err.Description
End Sub

Private Sub cmdMenosHorizontal_Click()
  On Error GoTo trata
  Dim objRegister As busSisMed.clsRegistro
  If intLarguraPadrao < 1000 Then Exit Sub
  intLarguraPadrao = intLarguraPadrao - 50
  cmdMenosHorizontal.ToolTipText = intLarguraPadrao
  cmdMaisHorizontal.ToolTipText = intLarguraPadrao
  RedimensionaBotoes
  'Saving Settings
  Set objRegister = New busSisMed.clsRegistro
  objRegister.SalvarChaveRegistro TITULOSISTEMA, _
                                  "WidthButtonScanner", _
                                  intLarguraPadrao & ""
  Set objRegister = Nothing
  
  
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source, _
            Err.Description
End Sub

Private Sub cmdMenosVertical_Click()
  On Error GoTo trata
  Dim objRegister As busSisMed.clsRegistro
  If intAlturaPadrao < 1000 Then Exit Sub
  intAlturaPadrao = intAlturaPadrao - 50
  cmdMenosVertical.ToolTipText = intAlturaPadrao
  cmdMaisVertical.ToolTipText = intAlturaPadrao
  RedimensionaBotoes
  'Saving Settings
  Set objRegister = New busSisMed.clsRegistro
  objRegister.SalvarChaveRegistro TITULOSISTEMA, _
                                  "HeightButtonScanner", _
                                  intAlturaPadrao & ""
  Set objRegister = Nothing
  
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source, _
            Err.Description
End Sub
Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco cmdCancelar
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserScannerCons.Form_Activate]"
End Sub


Private Sub Form_Load()
  Dim objRegister       As busSisMed.clsRegistro
  On Error GoTo trata
  '
  
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  If Me.ActiveControl Is Nothing Then
    Me.Top = 580
    Me.Left = 1
    Me.WindowState = 2 'Maximizado
  End If
  'Me.Height = 9000
  'Me.Width = 10500
  'CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, , , cmdCancelar
  '
  'NOVO BOTÕES NOVOS
  'Captura os dados do registro
  Set objRegister = New busSisMed.clsRegistro
  
  
  intLarguraPadrao = IIf(Not IsNumeric(objRegister.RetornarChaveRegistro(TITULOSISTEMA, _
                                                       "WidthButtonScanner")), 0, objRegister.RetornarChaveRegistro(TITULOSISTEMA, _
                                                       "WidthButtonScanner"))
  intAlturaPadrao = IIf(Not IsNumeric(objRegister.RetornarChaveRegistro(TITULOSISTEMA, _
                                                      "HeightButtonScanner")), 0, objRegister.RetornarChaveRegistro(TITULOSISTEMA, _
                                                      "HeightButtonScanner"))
  Set objRegister = Nothing
  'Default
  'Width 8415
  'Height 8445
  If intLarguraPadrao = 0 Or intAlturaPadrao = 0 Then
    Image1.Width = 8415
    Image1.Height = 8445
    intLarguraPadrao = 8415
    intAlturaPadrao = 8445
    
  Else
    Image1.Width = intLarguraPadrao
    Image1.Height = intAlturaPadrao
  End If
  cmdMenosHorizontal.ToolTipText = intLarguraPadrao
  cmdMaisHorizontal.ToolTipText = intLarguraPadrao
  cmdMenosVertical.ToolTipText = intAlturaPadrao
  cmdMaisVertical.ToolTipText = intAlturaPadrao
  'FIM NOVO BOTÕES
  
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
End Sub

