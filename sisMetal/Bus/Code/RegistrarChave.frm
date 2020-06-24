VERSION 5.00
Begin VB.Form frmRegistrarChave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Chave"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informações de Registro de estação"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4335
      Begin VB.TextBox txtDataAtual 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtChave 
         Height          =   285
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtNroEstacao 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Data Atual"
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
         TabIndex        =   9
         Top             =   360
         Width           =   1215
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
         Top             =   1080
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
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"RegistrarChave.frx":0000
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
Attribute VB_Name = "frmRegistrarChave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iQtdVezes        As Integer
Public sVolumeKey       As String
Public sDataAtual       As String
Public bRegistroValido  As Boolean
Public sBanco           As String
Public sEstourouPrazo   As String

Private Sub cmdCancela_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim rs            As ADODB.Recordset
  Dim rs1           As ADODB.Recordset
  Dim sSql          As String
  Dim sChave1       As String
  Dim sDataIni      As String
  Dim sDataSO       As String
  Dim sHoraAtual    As String
  Dim sQtdVezes     As String
  Dim lqtdCchaves   As Long
  Dim objGeral      As busSisMetal.clsGeral
  '
  If Len(Trim(txtChave.Text)) = 0 Then
    MsgBox "Por favor, entre com o número da chave !", vbOKOnly, gsTituloDll
    Exit Sub
  ElseIf Not IsNumeric(txtChave.Text) Then
    MsgBox "Por favor, entre com o valor da chave númerico !", vbOKOnly, gsTituloDll
    Exit Sub
  ElseIf Len(txtChave.Text) <> 9 Then
    MsgBox "Por favor, entre com o valor da chave com 9 posições !", vbOKOnly, gsTituloDll
    Exit Sub
  End If
  'validar Chave
  If Not Validar_Chave(txtNroEstacao.Text, txtChave.Text) Then
    iQtdVezes = iQtdVezes + 1
    If iQtdVezes > 2 Then
      MsgBox "Atenção !" & vbCrLf & vbCrLf & "Você tentou entrar no sistema com uma chave inválida por três vezes. Caso continue, certifique-se que a chave é válida, caso contrário o sistema não voltará mais a funcionar. Obrigado !", vbOKOnly, gsTituloDll
      Unload Me
    Else
      MsgBox "Chave inválida !", vbOKOnly, gsTituloDll
      Exit Sub
    End If
  Else
    'Chave é Válida
    'Verifica duplicidade da chave
    Set objGeral = New busSisMetal.clsGeral
    sSql = "Select count(*) from Chave Where Chave1 = " & Formata_Dados(Encripta(txtChave.Text), tpDados_Texto, tpNulo_Aceita) & ""
    Set rs = objGeral.ExecutarSQL(sSql)
    If rs(0) <> 0 Then
      MsgBox "Chave já cadastrada !", vbOKOnly, gsTituloDll
      rs.Close
      Set rs = Nothing
      Set objGeral = Nothing
      Exit Sub
    End If
    rs.Close
    Set rs = Nothing
    '
    sSql = "INSERT INTO Chave(Chave1) VALUES (" & Formata_Dados(Encripta(txtChave.Text), tpDados_Texto, tpNulo_Aceita) & ");"
    objGeral.ExecutarSQLAtualizacao sSql
    '
    If sEstourouPrazo = "S" Then
      sSql = "Select Count(*) From Chave;"
      Set rs1 = objGeral.ExecutarSQL(sSql)
      lqtdCchaves = rs1(0)
      rs1.Close
      Set rs1 = Nothing
      '
      sQtdVezes = Format(lqtdCchaves, "000000") & CalcDv(Format(lqtdCchaves, "000000"))
      '
      sSql = "UPDATE Gerencial set Chave4 = " & Formata_Dados(Encripta(sQtdVezes), tpDados_Texto, tpNulo_Aceita) & ";"
      objGeral.ExecutarSQLAtualizacao sSql
    Else
      sDataIni = Format(Now, "DD") & CalcDv(Format(Now, "DD"))
      sDataIni = sDataIni & Format(Now, "MM") & CalcDv(Format(Now, "MM"))
      sDataIni = sDataIni & Format(Now, "YYYY") & CalcDv(Format(Now, "YYYY"))
      sDataIni = sDataIni & CalcDv(Mid(sDataIni, 1, 10))
      sDataIni = sDataIni & CalcDv(Mid(sDataIni, 1, 12))
      sDataIni = sDataIni & CalcDv(Mid(sDataIni, 1, 3))
      '
      sDataSO = Format(Now, "DD") & CalcDv(Format(Now, "DD"))
      sDataSO = sDataSO & Format(Now, "MM") & CalcDv(Format(Now, "MM"))
      sDataSO = sDataSO & Format(Now, "YYYY") & CalcDv(Format(Now, "YYYY"))
      sDataSO = sDataSO & CalcDv(sDataSO)
      sDataSO = sDataSO & CalcDv(Mid(sDataSO, 1, 6))
      sDataSO = sDataSO & CalcDv(sDataSO)
      sHoraAtual = Format(Now, "hhmmss")
      sDataSO = sDataSO & sHoraAtual & CalcDv(sHoraAtual)
      
      '
      sQtdVezes = "000001" & CalcDv("000001")
      '
      sSql = "UPDATE Gerencial set Chave2 = " & Formata_Dados(Encripta(sDataIni), tpDados_Texto, tpNulo_Aceita) & ", Chave3 = " & Formata_Dados(Encripta(sDataSO), tpDados_Texto, tpNulo_Aceita) & ", Chave4 = " & Formata_Dados(Encripta(sQtdVezes), tpDados_Texto, tpNulo_Aceita) & ";"
      objGeral.ExecutarSQLAtualizacao sSql
    End If
    '-------------------------------
    bRegistroValido = True
    Set objGeral = Nothing
    Unload Me
  End If
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".frmRegistrarChave.cmdok_Click]", _
            Err.Description
End Sub

Private Sub Form_Activate()
  On Error GoTo trata
  txtNroEstacao.Text = sVolumeKey
  txtDataAtual.Text = sDataAtual
  If sEstourouPrazo = "S" Then
    Label1.Caption = "Atenção !!! Encerrou o prazo de validade de seu sistema. Por favor informe o Nro da Estação e a Data Atual do seu computador ao suporte para adquirir uma chave que será válida por mais 1 mês"
  End If
  txtChave.SetFocus
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".frmRegistrarChave.Form_Activate]", _
            Err.Description
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  CenterForm Me
  bRegistroValido = False
  iQtdVezes = 0
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".frmRegistrarChave.Form_Load]", _
            Err.Description
End Sub

