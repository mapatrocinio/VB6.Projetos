VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  
  Dim objWordDoc As Word.Document
  Dim objWord  As Word.Application
  
  Set objWord = New Word.Application
  Set objWordDoc = New Word.Document
  
  Set objWordDoc = objWord.Documents.Add()
  objWord.Visible = True
  
              objWord.Selection.MoveRight 1, 1
              objWord.Selection.TypeParagraph
                    
              objWordDoc.Tables.Add objWord.Selection.Range, 1, 1
              objWord.Selection.Tables(1).Columns(1).SetWidth 60, 1
              'Seleciona Coluna
              objWord.Selection.Tables(1).Columns(1).Select
              
              'On Error Resume Next
              
                      objWord.Selection.Tables(1).Cell(1, 1).Select
                      
                          objWord.Selection.Tables(1).Columns.Add
                          
                      objWord.Selection.Tables(1).Cell(1, 2).Select
                      
                          objWord.Selection.Shading.Texture = 100
                          objWord.Selection.ParagraphFormat.Alignment = 2
                          objWord.Selection.Cells.VerticalAlignment = 1
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
                          If "S" = "S" Then
                            objWord.Selection.Font.Bold = -1
                          Else
                            objWord.Selection.Font.Bold = 0
                          End If
                          objWord.Selection.TypeText "Janeiro"
                          'Borda
                          
                          objWord.Selection.Tables(1).Columns.Add
                          
                      objWord.Selection.Tables(1).Cell(1, 3).Select
                      
                          objWord.Selection.Shading.Texture = 100
                          objWord.Selection.ParagraphFormat.Alignment = 2
                          objWord.Selection.Cells.VerticalAlignment = 1
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
                          If "S" = "S" Then
                            objWord.Selection.Font.Bold = -1
                          Else
                            objWord.Selection.Font.Bold = 0
                          End If
                          objWord.Selection.TypeText "fevereiro"
                          'Borda
                          
                          objWord.Selection.Tables(1).Columns.Add
                          
                      objWord.Selection.Tables(1).Cell(1, 4).Select
                      
                          objWord.Selection.Shading.Texture = 100
                          objWord.Selection.ParagraphFormat.Alignment = 2
                          objWord.Selection.Cells.VerticalAlignment = 1
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
                          If "S" = "S" Then
                            objWord.Selection.Font.Bold = -1
                          Else
                            objWord.Selection.Font.Bold = 0
                          End If
                          objWord.Selection.TypeText "Março"
                          'Borda
                          
                          objWord.Selection.Tables(1).Columns.Add
                          
                      objWord.Selection.Tables(1).Cell(1, 5).Select
                      
                          objWord.Selection.Shading.Texture = 100
                          objWord.Selection.ParagraphFormat.Alignment = 2
                          objWord.Selection.Cells.VerticalAlignment = 1
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
                          If "S" = "S" Then
                            objWord.Selection.Font.Bold = -1
                          Else
                            objWord.Selection.Font.Bold = 0
                          End If
                          objWord.Selection.TypeText "Abril"
                          'Borda
                          
                      objWord.Selection.Tables(1).Rows.Add
                      
                      objWord.Selection.Tables(1).Cell(2, 1).Select
                      
                          objWord.Selection.Shading.Texture = 100
                          objWord.Selection.ParagraphFormat.Alignment = 1
                          objWord.Selection.Cells.VerticalAlignment = 1
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
                          If "S" = "S" Then
                            objWord.Selection.Font.Bold = -1
                          Else
                            objWord.Selection.Font.Bold = 0
                          End If
                          objWord.Selection.TypeText "1"
                          
                      objWord.Selection.Tables(1).Cell(2, 2).Select
                      
                          objWord.Selection.Shading.Texture = 0
                          objWord.Selection.ParagraphFormat.Alignment = 2
                          objWord.Selection.Cells.VerticalAlignment = 1
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
                          If "N" = "S" Then
                            objWord.Selection.Font.Bold = -1
                          Else
                            objWord.Selection.Font.Bold = 0
                          End If
                          objWord.Selection.TypeText "214,39"
                          
                      objWord.Selection.Tables(1).Cell(2, 3).Select
                      
                          objWord.Selection.Shading.Texture = 0
                          objWord.Selection.ParagraphFormat.Alignment = 2
                          objWord.Selection.Cells.VerticalAlignment = 1
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
                          If "N" = "S" Then
                            objWord.Selection.Font.Bold = -1
                          Else
                            objWord.Selection.Font.Bold = 0
                          End If
                          objWord.Selection.TypeText "214,39"
                          
                      objWord.Selection.Tables(1).Cell(2, 4).Select
                      
                          objWord.Selection.Shading.Texture = 0
                          objWord.Selection.ParagraphFormat.Alignment = 2
                          objWord.Selection.Cells.VerticalAlignment = 1
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
                          If "N" = "S" Then
                            objWord.Selection.Font.Bold = -1
                          Else
                            objWord.Selection.Font.Bold = 0
                          End If
                          objWord.Selection.TypeText "215,90"
                          
                      objWord.Selection.Tables(1).Cell(2, 5).Select
                      
                          objWord.Selection.Shading.Texture = 0
                          objWord.Selection.ParagraphFormat.Alignment = 2
                          objWord.Selection.Cells.VerticalAlignment = 1
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
                          If "N" = "S" Then
                            objWord.Selection.Font.Bold = -1
                          Else
                            objWord.Selection.Font.Bold = 0
                          End If
                          objWord.Selection.TypeText "216,00"
                          
                      objWord.Selection.Tables(1).Rows.Add
                      
                      objWord.Selection.Tables(1).Cell(3, 1).Select
                      
                          objWord.Selection.Shading.Texture = 100
                          objWord.Selection.ParagraphFormat.Alignment = 1
                          objWord.Selection.Cells.VerticalAlignment = 1
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
                          If "S" = "S" Then
                            objWord.Selection.Font.Bold = -1
                          Else
                            objWord.Selection.Font.Bold = 0
                          End If
                          objWord.Selection.TypeText "2"
                          
                      objWord.Selection.Tables(1).Cell(3, 2).Select
                      
                          objWord.Selection.Shading.Texture = 0
                          objWord.Selection.ParagraphFormat.Alignment = 2
                          objWord.Selection.Cells.VerticalAlignment = 1
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
                          If "N" = "S" Then
                            objWord.Selection.Font.Bold = -1
                          Else
                            objWord.Selection.Font.Bold = 0
                          End If
                          objWord.Selection.TypeText "217,00"
                          
                      objWord.Selection.Tables(1).Cell(3, 3).Select
                      
                          objWord.Selection.Shading.Texture = 0
                          objWord.Selection.ParagraphFormat.Alignment = 2
                          objWord.Selection.Cells.VerticalAlignment = 1
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
                          If "N" = "S" Then
                            objWord.Selection.Font.Bold = -1
                          Else
                            objWord.Selection.Font.Bold = 0
                          End If
                          objWord.Selection.TypeText "218,00"
                          
                      objWord.Selection.Tables(1).Cell(3, 4).Select
                      
                          objWord.Selection.Shading.Texture = 0
                          objWord.Selection.ParagraphFormat.Alignment = 2
                          objWord.Selection.Cells.VerticalAlignment = 1
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
                          If "N" = "S" Then
                            objWord.Selection.Font.Bold = -1
                          Else
                            objWord.Selection.Font.Bold = 0
                          End If
                          objWord.Selection.TypeText "219,00"
                          
                      objWord.Selection.Tables(1).Cell(3, 5).Select
                      
                          objWord.Selection.Shading.Texture = 0
                          objWord.Selection.ParagraphFormat.Alignment = 2
                          objWord.Selection.Cells.VerticalAlignment = 1
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
                          If "N" = "S" Then
                            objWord.Selection.Font.Bold = -1
                          Else
                            objWord.Selection.Font.Bold = 0
                          End If
                          objWord.Selection.TypeText "220"
                          
                         objWord.Selection.Tables(1).Columns(1).Delete
    
              'NOVO - APÓS MONTAGEM DA TABELA,
              'VARRER COLUNAS PARA QUEBRA-LA EM DIVERSAS TABELAS
              blnSairLoop = True
                    
              Do While blnSairLoop
                  For intCol = 1 To objWord.Selection.Tables(1).Columns.Count Step 8
                      'intCol varia 1, 9, 17, etc...
                      'Para cada bloco de tabelas ....
                      If objWord.Selection.Tables(1).Columns.Count > 8 Then
                          'caso exceda o tamanho máximo de cols na página,
                          'recortar o restante para montar nova tabela
                          'Selecionar
                          objWord.Selection.Tables(1).Cell(1, 9).Select
                          objWord.Selection.EndKey 10, 1
                          objWord.Selection.EndOf 9, 1
                          'Recortar
                          objWord.Selection.Cut
                          'Reorganiza as colunas da tabela
                          
  'Dim lngColCabec
  'Dim lngLinCabec
  'Dim strTextoAnterior
  'Dim strTextoAnteriorOutroNivel
  'Dim lngColAnteiror
  'Dim lngColMarge
  'Dim blnSelecionarUltimaLinha
  On Error Resume Next
  
                          'Adicionar parágrafo
                          objWord.Selection.MoveDown 5, 100
                          objWord.Selection.TypeParagraph
                          'Copiar
                          objWord.Selection.Paste

                      Else
                          blnSairLoop = False
                          'Reorganiza as colunas da tabela
                          
  'Dim lngColCabec
  'Dim lngLinCabec
  'Dim strTextoAnterior
  'Dim strTextoAnteriorOutroNivel
  'Dim lngColAnteiror
  'Dim lngColMarge
  'Dim blnSelecionarUltimaLinha
  On Error Resume Next
  
                      End If
                  Next
              Loop
              'objWord.Selection.MoveDown 5, 100
              objWord.Selection.EndKey 6
              objWord.Selection.TypeParagraph
  
  
  'objWord.Quit
  objWord.Visible = True
  Set objWordDoc = Nothing
  Set objWord = Nothing
  
End Sub
