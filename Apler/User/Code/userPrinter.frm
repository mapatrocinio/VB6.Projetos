VERSION 5.00
Begin VB.Form frmUserPrinter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Impressoras"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstPrinter 
      Height          =   2985
      Left            =   0
      TabIndex        =   3
      Top             =   60
      Width           =   4185
   End
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   3165
      Left            =   4290
      ScaleHeight     =   3165
      ScaleWidth      =   1860
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   1215
         Left            =   90
         ScaleHeight     =   1155
         ScaleWidth      =   1635
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1695
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   120
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmUserPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public blnFechar    As Boolean
Public strKey       As String
Private Sub cmdOk_Click()
  On Error GoTo trata
  If Not ValidaCampos Then Exit Sub
  Set Printer = GetDefaultPrinter(lstPrinter.Text)
''''  Printer.Print "Default printer is: " + Printer.DeviceName
''''  Printer.Print "Driver name is: " + Printer.DriverName
''''  Printer.Print "Port is: " + Printer.Port
''''  Printer.EndDoc
  'Slavar no register
  SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
            Key:=strKey, setting:=lstPrinter.Text

  blnFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg        As String
  Dim blnSetarFoco  As Boolean
  '
  blnSetarFoco = True
  If Not Valida_String(lstPrinter, TpObriga.TpObrigatorio, blnSetarFoco) Then
    strMsg = strMsg & "Selecionar uma impressora da lista" & vbCrLf
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserPrinter.ValidaCampos]"
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

Private Sub Form_Activate()
  SetarFoco lstPrinter
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  AmpS
  Me.Height = 3645
  Me.Width = 6240
  CenterForm Me
  LerFiguras Me, tpBmp_Vazio, pbtnImprimir:=cmdOk
  blnFechar = False
  GetPrinters lstPrinter
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo trata
  AmpS
  If blnFechar = False Then
    Cancel = True
    AmpN
    Exit Sub
  End If
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub lstPrinter_GotFocus()
  Selecionar_Conteudo lstPrinter
End Sub

Private Sub lstPrinter_LostFocus()
  Pintar_Controle lstPrinter, tpCorContr_Normal
End Sub

