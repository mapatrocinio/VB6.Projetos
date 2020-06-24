VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub SalvaParametrosRegisterUsuario(strTituloSistema As String, _
                                          strNomeUsu As String, _
                                          strNivel As String)
  On Error GoTo trata
  SalvarChaveRegistro strTituloSistema, _
                      "Usuario", _
                      strNomeUsu
  SalvarChaveRegistro strTituloSistema, _
                      "Nivel", _
                      strNivel
  SalvarChaveRegistro strTituloSistema, _
                      "Ativo", _
                      "S"
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsRegistro.SalvaParametrosRegisterUsuario]", _
            Err.Description
End Sub

Public Sub SalvaParametrosRegisterBMP(strTituloSistema As String, _
                                      strBMP As String)
  On Error GoTo trata
  SalvarChaveRegistro strTituloSistema, _
                      "BMP", _
                      strBMP
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsRegistro.SalvaParametrosRegisterBMP]", _
            Err.Description
End Sub
Public Sub SalvaParametrosRegisterNivel(strTituloSistema As String)
  On Error GoTo trata
  SalvarChaveRegistro strTituloSistema, _
                      "Nivel", _
                      ""
  SalvarChaveRegistro strTituloSistema, _
                      "Ativo", _
                      "N"
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsRegistro.SalvaParametrosRegisterNivel]", _
            Err.Description
End Sub

Public Sub CapturaParametrosRegister(ByRef strTituloSistema As String, _
                                     ByRef strReportPath As String, _
                                     ByRef strAppPath As String, _
                                     ByRef strNomeUsu As String, _
                                     ByRef strNomeEmpresa As String, _
                                     ByRef strBMPPath As String, _
                                     ByRef strIconsPath As String, _
                                     ByRef strBMP As String, _
                                     ByRef strPathBackup As String, _
                                     ByRef strNomeServidorBD As String)
    
    
''  strBDadosPath = RetornarChaveRegistro(strTituloSistema, _
''                                        "CaminhoDB")
  strReportPath = RetornarChaveRegistro(strTituloSistema, _
                                        "CaminhoReport")
  strAppPath = RetornarChaveRegistro(strTituloSistema, _
                                        "CaminhoApp")
  strNomeUsu = RetornarChaveRegistro(strTituloSistema, _
                                        "Usuario")
  strNomeEmpresa = RetornarChaveRegistro(strTituloSistema, _
                                        "Empresa")
  strBMPPath = RetornarChaveRegistro(strTituloSistema, _
                                        "CaminhoBMP")
  strIconsPath = RetornarChaveRegistro(strTituloSistema, _
                                        "CaminhoIcons")
  strBMP = RetornarChaveRegistro(strTituloSistema, _
                                        "BMP")
  strPathBackup = RetornarChaveRegistro(strTituloSistema, _
                                        "CaminhoBackUp")
  strNomeServidorBD = RetornarChaveRegistro(strTituloSistema, _
                                        "ServidorBD")
  On Error GoTo trata
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsRegistro.CapturaParametrosRegister]", _
            Err.Description
End Sub

Public Sub SalvarChaveRegistro(strTitulosSistema As String, _
                               strKey As String, _
                               strValor As String)
  On Error GoTo trata
  SaveSetting AppName:=strTitulosSistema, section:="Iniciar", _
            Key:=strKey, setting:=strValor
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsRegistro.SalvarChaveRegistro]", _
            Err.Description
End Sub

Public Function RetornarChaveRegistro(strTitulosSistema As String, _
                                      strKey As String)
  On Error GoTo trata
  RetornarChaveRegistro = GetSetting(AppName:=strTitulosSistema, section:="Iniciar", _
                          Key:=strKey)
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsRegistro.RetornarChaveRegistro]", _
            Err.Description
End Function

Public Sub TratarDimensaoBotoes(strTitulosSistema As String)
  On Error GoTo trata
  Dim strAltura     As String
  Dim strLargura    As String
  'Captura os dados do registro
  strAltura = RetornarChaveRegistro(strTitulosSistema, _
                                    "HeightButton")
  strLargura = RetornarChaveRegistro(strTitulosSistema, _
                                     "WidthButton")
  If Len(Trim(strAltura & "")) = 0 Then
    SalvarChaveRegistro strTitulosSistema, _
                        "HeightButton", _
                        "1000"
  End If
  If Len(Trim(strLargura & "")) = 0 Then
    SalvarChaveRegistro strTitulosSistema, _
                        "WidthButton", _
                        "1000"
  End If
  
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsRegistro.TratarDimensaoBotoes]", _
            Err.Description
End Sub



