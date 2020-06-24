VERSION 5.00
Begin VB.Form frmManutAssuntoCadastrar 
   Caption         =   "Cadastrar Assunto"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2250
   ScaleWidth      =   8160
   Begin VB.CommandButton btnFechar 
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6900
      TabIndex        =   6
      Top             =   540
      Width           =   735
   End
   Begin VB.CommandButton btmInserir 
      Caption         =   "Inserir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6900
      TabIndex        =   5
      Top             =   150
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   6435
      Begin VB.TextBox TxtAssunto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1770
         TabIndex        =   4
         Top             =   750
         Width           =   4425
      End
      Begin VB.ComboBox SelAssunto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1770
         TabIndex        =   2
         Top             =   270
         Width           =   4425
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição do Assunto:  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         TabIndex        =   3
         Top             =   780
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Assunto:  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   930
         TabIndex        =   1
         Top             =   300
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmManutAssuntoCadastrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim seq_assunto As Integer
Dim descr_assunto As String
Dim Sql As String
Dim i As Integer
Dim Rs As ADODB.Recordset

Private Sub btmInserir_Click()

On Error GoTo TrataErro
    If frmManutAssuntoCadastrar.TxtAssunto.Text = "" Then
        MsgBox "O campo 'Descrição do Assunto' deve ser informado!", vbExclamation
        Exit Sub
    End If
    
    'Set Rs = New ADODB.Recordset
    Sql = "select max(seq_assunto) as assunto from assunto "
    Set Rs = ExecutaSqlRs(Sql)
    If Not Rs.EOF Then
        seq_assunto = Rs("assunto")
        seq_assunto = seq_assunto + 1
    End If
    Rs.Close
    
    Sql = "insert into assunto(seq_assunto, descr_assunto, " & _
            "ind_status, dta_atualizacao, sig_usuario)" & _
            " values(" & seq_assunto & ",'" & TxtAssunto.Text & "'," & _
            "'S',date(),'tofelipe')"
    ExecutaSql (Sql)
    
    
    '********************************************************
    'Para carregar os assuntos cadastrados
    '********************************************************
    Sql = "select seq_assunto, descr_assunto from assunto"
       
    Set Rs = ExecutaSqlRs(Sql)
    
    SelAssunto.Clear
    Call CarregaAssunto(Rs, frmManutAssuntoCadastrar)
    Rs.Close
    Set Rs = Nothing
    MsgBox "Assunto cadastrado com sucesso!", vbInformation
    
    Exit Sub
    
TrataErro:
    MsgBox Err.Description, vbCritical
    
End Sub

Private Sub btnFechar_Click()
Unload frmManutAssuntoCadastrar
End Sub

Private Sub Form_Load()

Dim Sql As String

On Error GoTo TrataErro
Set Rs = New ADODB.Recordset

frmManutAssuntoCadastrar.Width = 8280
frmManutAssuntoCadastrar.Height = 2655

'********************************************************
'Para carregar os assuntos cadastrados
'********************************************************
Sql = "select seq_assunto, descr_assunto from assunto"
   
Set Rs = ExecutaSqlRs(Sql)

SelAssunto.Clear
Call CarregaAssunto(Rs, frmManutAssuntoCadastrar)
Rs.Close

Set cn = Nothing
Set Rs = Nothing
Exit Sub
TrataErro:
    MsgBox Err.Description, vbCritical, "Erro no aplicativo"
    Set cn = Nothing
    Set Rs = Nothing
End Sub

Private Sub SelAssunto_Click()
    If SelAssunto.ListIndex = -1 Then
        Exit Sub
    End If
    descr_assunto = SelAssunto.List(SelAssunto.ListIndex)
    TxtAssunto.Text = descr_assunto
End Sub
