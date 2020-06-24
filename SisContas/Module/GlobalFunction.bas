Attribute VB_Name = "mdlGlobalFunction"
Option Explicit
Public Const TITULOSISTEMA = "Sistema Gerenciador de Contas"

Public Enum TpTipoBD
  TpTipoBD_ACCESS = 0
  TpTipoBD_SQL = 1
End Enum

'Public Const intBancoDados As Integer = TpTipoBD.TpTipoBD_ACCESS
Public Const intBancoDados As Integer = TpTipoBD.TpTipoBD_SQL

Public Enum tpAceitaNulo
  tpNulo_Aceita
  tpNulo_NaoAceita
End Enum

Public Enum TpTipoDados
  tpDados_Texto '0 a 255
  tpDados_Memo 'Sem Limite
  tpDados_Inteiro '-32767 a 32767
  tpDados_Longo 'Sem limite
  tpDados_DataHora 'MM/DD/YYYY hh:mm:ss
  tpDados_Moeda '121212.98
  tpDados_Boolean '121212.98
End Enum

Public Function Tira_Plic(pValor As String) As Variant
  On Error Resume Next
  Tira_Plic = Replace(pValor, "'", "''")
End Function




Public Function Formata_Dados(pValor As Variant, _
                              pTipoDados As TpTipoDados, _
                              Optional pAceitaNulo As tpAceitaNulo = tpNulo_Aceita, _
                              Optional pTamanhoCampo As Integer) As Variant
  On Error GoTo trata
  '
  Dim vRetorno As Variant
  Dim sData As String
  '
  Select Case pTipoDados
  Case TpTipoDados.tpDados_Boolean
    If intBancoDados = TpTipoBD.TpTipoBD_ACCESS Then
      If pValor Then
        vRetorno = "True"
      Else
        vRetorno = "False"
      End If
    ElseIf intBancoDados = TpTipoBD.TpTipoBD_SQL Then
      If pValor Then
        vRetorno = 1
      Else
        vRetorno = 0
      End If
    End If
  Case TpTipoDados.tpDados_Texto
    If intBancoDados = TpTipoBD.TpTipoBD_ACCESS Then
      If Len(Trim(pValor & "")) = 0 Then
        If pAceitaNulo = tpNulo_Aceita Then
          vRetorno = "Null"
        Else
          vRetorno = "' '"
        End If
      Else
        vRetorno = "'" & Tira_Plic(Trim(pValor & "")) & "'"
      End If
    ElseIf intBancoDados = TpTipoBD.TpTipoBD_SQL Then
      If Len(Trim(pValor & "")) = 0 Then
        If pAceitaNulo = tpNulo_Aceita Then
          vRetorno = "Null"
        Else
          vRetorno = "' '"
        End If
      Else
        vRetorno = "'" & Tira_Plic(Trim(pValor & "")) & "'"
      End If
    End If
  Case TpTipoDados.tpDados_Longo
    If intBancoDados = TpTipoBD.TpTipoBD_ACCESS Then
      If Not IsNumeric(pValor) Then
        If pAceitaNulo = tpNulo_Aceita Then
          vRetorno = "Null"
        Else
          vRetorno = "0"
        End If
      Else
        vRetorno = CLng(pValor)
      End If
    ElseIf intBancoDados = TpTipoBD.TpTipoBD_SQL Then
      If Not IsNumeric(pValor) Then
        If pAceitaNulo = tpNulo_Aceita Then
          vRetorno = "Null"
        Else
          vRetorno = "0"
        End If
      Else
        vRetorno = CLng(pValor)
      End If
    End If
  Case TpTipoDados.tpDados_DataHora
    If intBancoDados = TpTipoBD.TpTipoBD_ACCESS Then
      'Converter para Data
      sData = ""
      If Len(pValor & "") = 10 Then
        If Mid(pValor & "", 1, 2) <> "__" Then
          'Data no Formato DD/MM/YYYY
          sData = "#" & Mid(pValor, 4, 2) & "/" & Mid(pValor, 1, 2) & "/" & Mid(pValor, 7, 4) & "#"
        Else
          sData = "null"
        End If
      ElseIf Len(pValor & "") = 16 Then
        'Data no Formato DD/MM/YYYY hh:mm
        If Mid(pValor & "", 1, 2) <> "__" Then
          'Data no Formato DD/MM/YYYY
          sData = "#" & Mid(pValor, 4, 2) & "/" & Mid(pValor, 1, 2) & "/" & Mid(pValor, 7, 4) & " " & Mid(pValor, 12, 2) & ":" & Mid(pValor, 15, 2) & "#"
        Else
          sData = "null"
        End If
      Else
        sData = ""
      End If
      If Len(sData) = 0 Then
        If pAceitaNulo = tpNulo_Aceita Then
          vRetorno = "Null"
        Else
          vRetorno = "#01/01/1900#"
        End If
      Else
        vRetorno = sData
      End If
    ElseIf intBancoDados = TpTipoBD.TpTipoBD_SQL Then
      sData = ""
      If Len(pValor & "") = 10 Then
        If Mid(pValor & "", 1, 2) <> "__" Then
          'Data no Formato DD/MM/YYYY
          sData = "Convert(DateTime, '" & Mid(pValor, 1, 2) & "/" & Mid(pValor, 4, 2) & "/" & Mid(pValor, 7, 4) & "')"
        Else
          sData = "null"
        End If
      ElseIf Len(pValor & "") = 5 Then
        If pValor & "" <> "__:__" Then
          'Data no Formato hh:mm
          sData = "Convert(DateTime, '01/01/1900 " & pValor & "')"
        Else
          sData = "null"
        End If
      ElseIf Len(pValor & "") = 16 Then
        'Data no Formato DD/MM/YYYY hh:mm
        If Mid(pValor & "", 1, 2) <> "__" Then
          'Data no Formato DD/MM/YYYY
          sData = "Convert(DateTime, '" & Mid(pValor, 1, 2) & "/" & Mid(pValor, 4, 2) & "/" & Mid(pValor, 7, 4) & " " & Mid(pValor, 12, 2) & ":" & Mid(pValor, 15, 2) & "')"
        Else
          sData = "null"
        End If
      ElseIf Len(pValor & "") = 19 Then
        'Data no Formato DD/MM/YYYY hh:mm
        If Mid(pValor & "", 1, 2) <> "__" Then
          'Data no Formato DD/MM/YYYY
          sData = "Convert(DateTime, '" & Mid(pValor, 1, 2) & "/" & Mid(pValor, 4, 2) & "/" & Mid(pValor, 7, 4) & " " & Mid(pValor, 12, 2) & ":" & Mid(pValor, 15, 2) & ":" & Mid(pValor, 18, 2) & "')"
        Else
          sData = "null"
        End If
      Else
        sData = ""
      End If
      If Len(sData) = 0 Then
        vRetorno = "Null"
      Else
        vRetorno = sData
      End If
    End If
  Case TpTipoDados.tpDados_Moeda
    If intBancoDados = TpTipoBD.TpTipoBD_ACCESS Then
      If Not IsNumeric(pValor) Then
        If pAceitaNulo = tpNulo_Aceita Then
          vRetorno = "Null"
        Else
          vRetorno = "0"
        End If
      Else
        vRetorno = Replace(pValor, ".", "")
        vRetorno = Replace(vRetorno, ",", ".")
      End If
    ElseIf intBancoDados = TpTipoBD.TpTipoBD_SQL Then
      If Not IsNumeric(pValor) Then
        If pAceitaNulo = tpNulo_Aceita Then
          vRetorno = "Null"
        Else
          vRetorno = "0"
        End If
      Else
        vRetorno = Replace(pValor, ".", "")
        vRetorno = Replace(vRetorno, ",", ".")
      End If
    End If
  Case Else
  End Select
  '
  Formata_Dados = vRetorno
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[GlobalFunction.Formata_Dados]", _
            Err.Description
End Function



