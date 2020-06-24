VERSION 5.00
Begin VB.Form frmManutPergCadastrar 
   Caption         =   "Cadastrar Perguntas"
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
      Begin VB.TextBox TxtItemSa 
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
      Begin VB.ComboBox SelItemSa 
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
         Caption         =   "Descrição da Pergunta:"
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
         Left            =   150
         TabIndex        =   7
         Top             =   1275
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Perguntas Cadastradas:  "
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
Attribute VB_Name = "frmManutPergCadastrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim seq_item_sa As Variant
Dim seq_assunto As Integer
Dim descr_item_sa As String
Dim Rs As ADODB.Recordset
Dim Sql As String
Private Sub btmInserir_Click()

On Error GoTo TrataErro
    '****************************************************
    'validação dos campos para o cadastro
    '****************************************************
    If SelAssunto.ListIndex = -1 Then
        MsgBox "O campo 'Assunto' deve ser informado.", vbInformation
        Exit Sub
    End If
    If Trim(TxtItemSa.Text) = "" Then
        MsgBox "O campo 'Descrição da Pergunta' deve ser informado.", vbInformation
        Exit Sub
    End If
    
    seq_assunto = SelAssunto.ItemData(SelAssunto.ListIndex)
    descr_item_sa = Trim(TxtItemSa.Text)
    
    '****************************************************
    'Verifica qual o maior seq_item_sa para adicionar um novo registro
    '****************************************************
    Sql = " select iif(max(seq_item_sa) = null,0,max(seq_item_sa)) " & _
            " as itemsa from item_sa where seq_assunto = " & seq_assunto
    Set Rs = ExecutaSqlRs(Sql)
    If Not Rs.EOF Then
        seq_seq_item_sa = Rs("itemsa")
        seq_item_sa = seq_item_sa + 1
    End If
    Rs.Close
    Sql = " insert into item_sa(seq_assunto, seq_item_sa,descr_item_sa," & _
            " ind_status,ind_item_extra,dta_atualizacao, sig_usuario) values" & _
            " (" & seq_assunto & "," & seq_item_sa & ",'" & descr_item_sa & "','A'," & _
            "'N',date(),'tofelipe')"

    ExecutaSql (Sql)
    
    '***********************************************************
    'Para atualizar a combo com as perguntas cadastradas
    '***********************************************************
    
    Sql = "select seq_item_sa,descr_item_sa,ind_item_extra    " & _
            " From item_sa where seq_assunto= " & seq_assunto & _
            " and ind_status = 'A' and ind_item_extra = 'N' "
     
    Set Rs = ExecutaSqlRs(Sql)
    If Not Rs.EOF Then
        SelItemSa.Clear
        CarregaPergunta Rs, frmManutPergCadastrar
    End If
    Rs.Close
    
    MsgBox "Pergunta cadastrada com sucesso!", vbInformation
    
    Exit Sub
    
TrataErro:
    MsgBox Err.Description, vbCritical, "Erro no Aplicativo"
    Set cn = Nothing
    Set Rs = Nothing
End Sub

Private Sub btnFechar_Click()
Unload frmManutPergCadastrar
End Sub

Private Sub Form_Load()

Dim Sql As String

On Error GoTo TrataErro
Set Rs = New ADODB.Recordset

frmManutPergCadastrar.Width = 8280
frmManutPergCadastrar.Height = 2805


'********************************************************
'Para carregar os assuntos cadastrados
'********************************************************
Sql = "select seq_assunto, descr_assunto from assunto"
   
Set Rs = ExecutaSqlRs(Sql)

SelAssunto.Clear
Call CarregaAssunto(Rs, frmManutPergCadastrar)
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
    Sql = "select seq_item_sa,descr_item_sa,ind_item_extra    " & _
            " From item_sa where seq_assunto= " & seq_assunto & _
            " and ind_status = 'A' and ind_item_extra = 'N' "
     
    Set Rs = ExecutaSqlRs(Sql)
    If Not Rs.EOF Then
        SelItemSa.Clear
        CarregaPergunta Rs, frmManutPergCadastrar
    End If
    Rs.Close
    Exit Sub
TrataErro:
    MsgBox Err.Description, vbCritical, "Erro no aplicativo"
    Set cn = Nothing
    Set Rs = Nothing
End Sub

Private Sub SelItemSa_Click()
    If SelItemSa.ListIndex = -1 Then
        Exit Sub
    End If
    
    descr_item_sa = SelItemSa.List(SelItemSa.ListIndex)
    TxtItemSa.Text = descr_item_sa

End Sub
