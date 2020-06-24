VERSION 5.00
Begin VB.Form frmRegistrarEstacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Estação"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informações de Registro de estação"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4335
      Begin VB.TextBox txtChave 
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtNroEstacao 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Chave"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Nro. da Estação"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"RegistrarEstacao.frx":0000
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmRegistrarEstacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sVolumeKey       As String
Public bRegistroValido  As Boolean
Public sBanco           As String

Private Sub cmdCancela_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim rs          As ADODB.Recordset
  Dim rs1         As ADODB.Recordset
  Dim sSql        As String
  Dim sChave1     As String
  Dim objGeral    As busSisLoc.clsGeral
  '
  If Len(Trim(txtChave.Text)) = 0 Then
    MsgBox "Por favor, entre com o número da chave !", vbOKOnly, gsTituloDll
    Exit Sub
  ElseIf Not IsNumeric(txtChave.Text) Then
    MsgBox "Por favor, entre com o valor da chave númerico !", vbOKOnly, gsTituloDll
    Exit Sub
  ElseIf Len(txtChave.Text) <> 6 Then
    MsgBox "Por favor, entre com o valor da chave com 6 posições !", vbOKOnly, gsTituloDll
    Exit Sub
  End If
  'validar Chave
  If Not Validar_Chave_HD(txtNroEstacao.Text, txtChave.Text) Then
    MsgBox "Chave inválida !", vbOKOnly, gsTituloDll
    Exit Sub
  Else
    'Chave é Válida
    'Gravar Registro
    Set objGeral = New busSisLoc.clsGeral
    SaveSetting appname:=sAppName, section:=sSection, _
              Key:=sKey1, setting:=Encripta(txtChave.Text)
    'Gravar Register na base de dados
    '-------------------------------
    sSql = "Select count(*) from Gerencial"
    Set rs = objGeral.ExecutarSQL(sSql)
    If rs(0) = 0 Then
      sSql = "INSERT INTO Gerencial(Chave1) VALUES (" & Formata_Dados(Encripta(txtChave.Text), tpDados_Texto, tpNulo_Aceita) & ");"
      objGeral.ExecutarSQLAtualizacao sSql
    Else
      sSql = "Select Chave1 from Gerencial;"
      Set rs1 = objGeral.ExecutarSQL(sSql)
      '
      If rs1.EOF Then
        sChave1 = ""
      Else
        sChave1 = rs1(0) & ""
      End If
      rs1.Close
      Set rs1 = Nothing
      'Verifica se Máquina já está gravada
      If InStr(1, sChave1, Encripta(txtChave.Text)) = 0 Then
        'Estação não está cadastrada
        sSql = "UPDATE Gerencial Set Chave1 = " & Formata_Dados(sChave1 & Encripta(txtChave.Text), tpDados_Texto, tpNulo_Aceita) & ";"
        objGeral.ExecutarSQLAtualizacao sSql
      End If
    End If
    rs.Close
    Set rs = Nothing
    Set objGeral = Nothing
    '-------------------------------
    bRegistroValido = True
    Unload Me
  End If
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".frmRegistrarEstacao.cmdOk_Click]", _
            Err.Description
End Sub

Private Sub Form_Activate()
  On Error GoTo trata
  txtChave.SetFocus
  txtNroEstacao.Text = sVolumeKey
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".frmRegistrarEstacao.Form_Activate]", _
            Err.Description
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  CenterForm Me
  bRegistroValido = False
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".frmRegistrarEstacao.Form_Load]", _
            Err.Description
End Sub
