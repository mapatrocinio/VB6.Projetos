VERSION 5.00
Begin VB.Form frmManutChklstCadastrar 
   Caption         =   "Cadastrar Checklist"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2400
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
      Left            =   7200
      TabIndex        =   5
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
      Left            =   7200
      TabIndex        =   4
      Top             =   150
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   2085
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   6915
      Begin VB.TextBox TxtChklst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1890
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1170
         Width           =   4875
      End
      Begin VB.ComboBox SelChklst 
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
         Left            =   1890
         TabIndex        =   6
         Top             =   750
         Width           =   4875
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
         Left            =   1890
         TabIndex        =   2
         Top             =   330
         Width           =   4875
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descrição do Ítem:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   510
         TabIndex        =   7
         Top             =   1275
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Checklists Cadastrados:  "
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
         Top             =   825
         Width           =   1785
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Assunto:"
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
         Left            =   960
         TabIndex        =   1
         Top             =   375
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmManutChklstCadastrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim seq_checklist As Variant
Dim seq_assunto As Variant
Dim descr_checklist As String
Dim Rs As ADODB.Recordset
Dim Sql As String
Private Sub btmInserir_Click()

On Error GoTo TrataErro
    '****************************************************
    'validação dos campos para o cadastro
    '****************************************************
    If SelAssunto.ListIndex = -1 Then
        MsgBox "Um assunto deve ser informado.", vbInformation
        Exit Sub
    End If
    If Trim(TxtChklst.Text) = "" Then
        MsgBox "O campo 'Descrição do ítem' deve ser informado.", vbInformation
        Exit Sub
    End If
    
    seq_assunto = SelAssunto.ItemData(SelAssunto.ListIndex)
    descr_checklist = Trim(TxtChklst.Text)
    '****************************************************
    'Verifica qual o maior seq_checklist para adicionar um novo registro
    '****************************************************
    Sql = " select iif(max(seq_checklist) = null,0,max(seq_checklist)) " & _
            " as checklist from checklist where seq_assunto = " & seq_assunto
    Set Rs = ExecutaSqlRs(Sql)
    If Not Rs.EOF Then
        seq_checklist = Rs("checklist")
        seq_checklist = seq_checklist + 1
    End If
    Rs.Close
    
    Sql = " insert into checklist(seq_assunto, seq_checklist, descr_checklist, " & _
            " dta_atualizacao, sig_usuario, ind_desativacao) values" & _
            " (" & seq_assunto & "," & seq_checklist & ",'" & descr_checklist & "'," & _
            " date(), 'tofelipe','A')"
    ExecutaSql (Sql)
    
    '***********************************************************
    'Para atualizar a combo com os checklists cadastrados
    '***********************************************************
    
    Sql = "select seq_checklist,descr_checklist " & _
            " From checklist where ind_desativacao  = 'A' and" & _
            " seq_assunto=  " & seq_assunto
    Set Rs = ExecutaSqlRs(Sql)
    If Not Rs.EOF Then
        SelChklst.Clear
        CarregaChecklist Rs, frmManutChklstCadastrar
    End If
    Rs.Close
    
    MsgBox "Checklist cadastrado com sucesso!", vbInformation
    
    Exit Sub
    
TrataErro:
    MsgBox Err.Description, vbCritical, "Erro no Aplicativo"
    Set cn = Nothing
    Set Rs = Nothing
End Sub

Private Sub btnFechar_Click()
Unload frmManutChklstCadastrar
End Sub

Private Sub Form_Load()

Dim Sql As String

On Error GoTo TrataErro
Set Rs = New ADODB.Recordset

frmManutChklstCadastrar.Width = 8280
frmManutChklstCadastrar.Height = 2805


'********************************************************
'Para carregar os assuntos cadastrados
'********************************************************
Sql = "select seq_assunto, descr_assunto from assunto"
   
Set Rs = ExecutaSqlRs(Sql)

SelAssunto.Clear
Call CarregaAssunto(Rs, frmManutChklstCadastrar)
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
On Error GoTo TrataErro
    If SelAssunto.ListIndex = -1 Then
        Exit Sub
    End If

    seq_assunto = SelAssunto.ItemData(SelAssunto.ListIndex)
    Sql = "select seq_checklist,descr_checklist " & _
            " From checklist where ind_desativacao  = 'A' and" & _
            " seq_assunto=  " & seq_assunto
    Set Rs = ExecutaSqlRs(Sql)
    If Not Rs.EOF Then
        SelChklst.Clear
        CarregaChecklist Rs, frmManutChklstCadastrar
    End If
    Rs.Close
    Exit Sub
TrataErro:
    MsgBox Err.Description, vbCritical, "Erro no aplicativo"
    Set cn = Nothing
    Set Rs = Nothing
End Sub

Private Sub SelChklst_Click()
    If SelChklst.ListIndex = -1 Then
        Exit Sub
    End If
    
    descr_checklist = SelChklst.List(SelChklst.ListIndex)
    TxtChklst.Text = descr_checklist
End Sub
