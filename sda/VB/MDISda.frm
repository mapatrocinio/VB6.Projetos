VERSION 5.00
Begin VB.MDIForm MDIsda 
   BackColor       =   &H8000000C&
   Caption         =   "SDA - Sistema de Auditoria"
   ClientHeight    =   2910
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3885
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDISda.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu OpManut 
      Caption         =   "&Manutenção"
      Begin VB.Menu OpManutTI 
         Caption         =   "&Tabelas Internas"
         Begin VB.Menu OpManutAssunto 
            Caption         =   "&Assunto"
            Begin VB.Menu OpManutAssuntoCadastrar 
               Caption         =   "Cadastrar"
            End
         End
         Begin VB.Menu OpManutChklst 
            Caption         =   "&Checklist"
            Begin VB.Menu OpManutChklstCadastrar 
               Caption         =   "Cadastrar"
            End
         End
         Begin VB.Menu OpManutPerg 
            Caption         =   "&Perguntas"
            Begin VB.Menu OpManutPergCadastrar 
               Caption         =   "Cadastrar"
            End
         End
         Begin VB.Menu OpManutPergChklst 
            Caption         =   "Per&guntas x Checklist"
            Begin VB.Menu OpManutPergChklstCadastrar 
               Caption         =   "Cadastrar"
            End
         End
         Begin VB.Menu OpManutTmp 
            Caption         =   "&Templates"
            Begin VB.Menu OpManutTmpCadastrar 
               Caption         =   "Cadastrar"
            End
         End
      End
   End
   Begin VB.Menu OpProcesso 
      Caption         =   "&Processo"
      Begin VB.Menu OpProcessoConsultar 
         Caption         =   "&Consultar"
      End
   End
   Begin VB.Menu OpSa 
      Caption         =   "&SA"
      Begin VB.Menu OpSaConsultar 
         Caption         =   "&Consultar"
      End
   End
   Begin VB.Menu OpSair 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "MDIsda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OpManutAssuntoCadastrar_Click()
Load frmManutAssuntoCadastrar
End Sub

Private Sub OpManutChklstCadastrar_Click()
Load frmManutChklstCadastrar
End Sub

Private Sub OpManutPergCadastrar_Click()
Load frmManutPergCadastrar
End Sub

Private Sub OpManutPergChklstCadastrar_Click()
Load frmManutPergChklstCadastrar
End Sub

Private Sub OpManutTmpCadastrar_Click()
Load frmManutTemplateCadastrar
End Sub

Private Sub OpProcessoConsultar_Click()
Load frmProcessoConsultar
End Sub

Private Sub OpSaConsultar_Click()
    Load frmSAConsultar
End Sub

Private Sub OpSair_Click()
Unload MDIsda
End Sub
