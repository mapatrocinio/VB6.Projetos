VERSION 5.00
Begin VB.Form frmManutTemplateCadastrar 
   Caption         =   "Cadastrar Templates"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4125
   ScaleWidth      =   8160
   Begin VB.ComboBox SelPalavraChave 
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
      Left            =   2040
      TabIndex        =   8
      Top             =   1230
      Width           =   3255
   End
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
      Left            =   7260
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
      Left            =   7260
      TabIndex        =   4
      Top             =   150
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   3945
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   6915
      Begin VB.CommandButton btnAdicionar 
         Caption         =   "Adicionar ao Texto"
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
         Left            =   5250
         MaskColor       =   &H00C00000&
         TabIndex        =   13
         Top             =   1200
         Width           =   1515
      End
      Begin VB.TextBox TxtTemplate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   1890
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2010
         Width           =   4875
      End
      Begin VB.ComboBox SelTemplate 
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
         TabIndex        =   10
         Top             =   1590
         Width           =   4875
      End
      Begin VB.ComboBox SelTopicoDoc 
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
      Begin VB.ComboBox SelTipoTemplate 
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Left            =   1065
         TabIndex        =   11
         Top             =   2115
         Width           =   780
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Templates Cadastrados:"
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
         Left            =   90
         TabIndex        =   9
         Top             =   1710
         Width           =   1755
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Campos chave:"
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
         Left            =   720
         TabIndex        =   7
         Top             =   1275
         Width           =   1125
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tópico do Documento:"
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
         Left            =   240
         TabIndex        =   3
         Top             =   870
         Width           =   1605
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Template:"
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
         Left            =   585
         TabIndex        =   1
         Top             =   450
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmManutTemplateCadastrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim seq_tipo_documento As Integer
Dim seq_topico_doc As Integer
Dim seq_template As Variant
Dim descr_template As String
Dim campo_chave As String
Dim Rs As ADODB.Recordset
Dim Sql As String
Private Sub btmInserir_Click()

On Error GoTo TrataErro
    If SelTipoTemplate.ListIndex = -1 Then
        MsgBox "Para cadastrar um template, o tipo do documento deve ser informado.", vbExclamation
        Exit Sub
    End If
    If SelTopicoDoc.ListIndex = -1 Then
        MsgBox "Para cadastrar um template, o tópico do documento deve ser informado.", vbExclamation
        Exit Sub
    End If
    If Trim(TxtTemplate.Text) = "" Then
        MsgBox "O campo 'Descrição' deve ser informado.", vbExclamation
        Exit Sub
    End If
    Sql = " select iif(max(seq_template)=null,0,max(seq_template)) as template" & _
            " from template_doc_auditoria where seq_tipo_documento = " & _
            seq_tipo_documento & " and seq_topico_doc =" & seq_topico_doc

    Set Rs = ExecutaSqlRs(Sql)
    If Not Rs.EOF Then
        seq_template = Rs("Template")
        seq_template = seq_template + 1
    End If
    Rs.Close
    
    seq_tipo_documento = SelTipoTemplate.ItemData(SelTipoTemplate.ListIndex)
    seq_topico_doc = SelTopicoDoc.ItemData(SelTopicoDoc.ListIndex)
    descr_template = Trim(TxtTemplate.Text)
    
    Sql = " insert into template_doc_auditoria(seq_tipo_documento," & _
            " seq_topico_doc, seq_template, descr_template," & _
            " dta_atualizacao, sig_usuario, ind_desativacao)" & _
            " values (" & seq_tipo_documento & "," & seq_topico_doc & "," & _
            seq_template & ",'" & descr_template & "',date(),'tofelipe','A')"
    
    ExecutaSql (Sql)
    
    
    Sql = "select seq_template, descr_template " & _
          " From template_doc_auditoria " & _
          " where  seq_tipo_documento = " & seq_tipo_documento & _
          " and seq_topico_doc= " & seq_topico_doc & _
          " and ind_desativacao='A'"
  
    Set Rs = ExecutaSqlRs(Sql)
    If Not Rs.EOF Then
        SelTemplate.Clear
        CarregaTemplate Rs, frmManutTemplateCadastrar
    End If
    Rs.Close
    
    MsgBox "Template cadastrado com sucesso.", vbInformation
    
    Exit Sub
    
    
TrataErro:
    MsgBox Err.Description, vbCritical, "Erro no Aplicativo"
    Set cn = Nothing
    Set Rs = Nothing
End Sub

Private Sub btnAdicionar_Click()
    If SelPalavraChave.ListIndex = -1 Then
        Exit Sub
    End If
    campo_chave = SelPalavraChave.List(SelPalavraChave.ListIndex)
    descr_template = TxtTemplate.Text
    descr_template = descr_template & " " & campo_chave
    TxtTemplate.Text = descr_template
End Sub

Private Sub btnFechar_Click()
Unload frmManutTemplateCadastrar
End Sub


Private Sub Form_Load()

Dim Sql As String

On Error GoTo TrataErro
Set Rs = New ADODB.Recordset

frmManutTemplateCadastrar.Width = 8280
frmManutTemplateCadastrar.Height = 4530

'*****************************************
'Carregar a combo com os tipos de template
'*****************************************
Sql = "select seq_tipo_documento, descr_tipo_documento    " & _
        " From tipo_doc_auditoria"
Set Rs = ExecutaSqlRs(Sql)
If Not Rs.EOF Then
    SelTipoTemplate.Clear
    CarregaTipoTemplate Rs, frmManutTemplateCadastrar
End If
Rs.Close

Sql = "select campo_chave from campochave"
Set Rs = ExecutaSqlRs(Sql)
If Not Rs.EOF Then
    SelPalavraChave.Clear
    CarregaCampoChave Rs, frmManutTemplateCadastrar
End If


Set cn = Nothing
Set Rs = Nothing
Exit Sub
TrataErro:
    MsgBox Err.Description, vbCritical, "Erro no aplicativo"
    Set cn = Nothing
    Set Rs = Nothing
End Sub

Private Sub SelTemplate_Click()
    If SelTemplate.ListIndex = -1 Then
        Exit Sub
    End If
    descr_template = SelTemplate.List(SelTemplate.ListIndex)
    TxtTemplate.Text = descr_template
    
    
End Sub

Private Sub SelTipoTemplate_Click()
    If SelTipoTemplate.ListIndex = -1 Then
        Exit Sub
    End If
    
    '***************************************
    'Carregar os tópicos do template de acordo com o tipo de documento
    '***************************************
    seq_tipo_documento = SelTipoTemplate.ItemData(SelTipoTemplate.ListIndex)
    SelTopicoDoc.Clear
    SelTemplate.Clear
    Sql = "select seq_topico_doc,descr_conteudo " & _
            " From topico_doc_auditoria" & _
            " where seq_tipo_Documento= " & seq_tipo_documento
    Set Rs = ExecutaSqlRs(Sql)
    If Not Rs.EOF Then
        CarregaTopicoTemplate Rs, frmManutTemplateCadastrar
    End If
    Rs.Close
    
    If SelTopicoDoc.ListIndex = -1 Then
        Exit Sub
    End If
    
    seq_topico_doc = SelTopicoDoc.ItemData(SelTopicoDoc.ListIndex)
    
    Sql = "select seq_template, descr_template " & _
          " From template_doc_auditoria " & _
          " where  seq_tipo_documento = " & seq_tipo_documento & _
          " and seq_topico_doc= " & seq_topico_doc & _
          " and ind_desativacao='A'"
  
    Set Rs = ExecutaSqlRs(Sql)
    If Not Rs.EOF Then
        CarregaTemplate Rs, frmManutTemplateCadastrar
    End If
    Rs.Close
    
End Sub

Private Sub SelTopicoDoc_Click()

    If SelTipoTemplate.ListIndex = -1 Then
        Exit Sub
    End If
    seq_tipo_documento = SelTipoTemplate.ItemData(SelTipoTemplate.ListIndex)
    SelTemplate.Clear
    If SelTopicoDoc.ListIndex = -1 Then
        Exit Sub
    End If
    
    seq_topico_doc = SelTopicoDoc.ItemData(SelTopicoDoc.ListIndex)
    
    Sql = "select seq_template, descr_template " & _
          " From template_doc_auditoria " & _
          " where  seq_tipo_documento = " & seq_tipo_documento & _
          " and seq_topico_doc= " & seq_topico_doc & _
          " and ind_desativacao='A'"
  
    Set Rs = ExecutaSqlRs(Sql)
    If Not Rs.EOF Then
        CarregaTemplate Rs, frmManutTemplateCadastrar
    End If
    Rs.Close

End Sub
