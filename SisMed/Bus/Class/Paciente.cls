VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarPacientePeloPkid(lngPRONTUARIOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objPaciente As datSisMed.clsPaciente
  Set objPaciente = New datSisMed.clsPaciente
  
  'Listar Paciente
  Set SelecionarPacientePeloPkid = objPaciente.SelecionarPacientePeloPkid(lngPRONTUARIOID)
  '
  Set objPaciente = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Sub InserirPaciente(lngPRONTUARIOID As Long)
  On Error GoTo trata
  Dim objPaciente     As datSisMed.clsPaciente
  '
  Set objPaciente = New datSisMed.clsPaciente
  'Inserir Paciente
  objPaciente.InserirPaciente lngPRONTUARIOID
  '
  Set objPaciente = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarPaciente(lngPRONTUARIOID As Long)
  On Error GoTo trata
  Dim objPaciente     As datSisMed.clsPaciente
  '
  Set objPaciente = New datSisMed.clsPaciente
  'Atualizar Paciente
  objPaciente.AlterarPaciente lngPRONTUARIOID
  '
  Set objPaciente = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


