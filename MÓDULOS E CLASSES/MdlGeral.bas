Attribute VB_Name = "MdlGeral"
Option Explicit
   Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
   Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
      
   Global sysSistema             As String
   Global sysUsuario             As String
   Global sysPerfil              As String
      
   Public Cnn                    As New cConecta
   Public SqlSrv                 As New CConSql
      
   Public RsConsulta             As New ADODB.Recordset
   Public RsTipoCombo            As New ADODB.Recordset
   Public Rst                    As New ADODB.Recordset
   Public Rs                     As New ADODB.Recordset
   Public Rsr                    As New ADODB.Recordset
   
   Public VDataServ              As Date
   Public VHoraServ              As String
   Public Autoriza               As Boolean
   Public CONTINUA               As Boolean
   Public sBuscaCnpjCNOME        As String
   Public sBuscaCnpjCTRAN        As String
   Public FormularioAtual        As Form
   Public FrmAtual               As Form
   
   Public Impres                 As Printer
   
   Public vsql                   As String
   Public CmdSql                 As String
   Public strCombo               As String
   
   Public I                      As Long
   
   Public MnNumce                As Long
   Public sBuscaProg             As String
   Public nEmpresa               As Long
   
   Public SC_FUNCIONARIO         As Long
   Public SC_CENTRODECUSTO       As Integer
   Public SC_SENHA               As String
   Public SC_NOMEFUNCIONARIO     As String
   Public SC_EMPRESA             As Integer
   Public SC_MAQNAME             As String
   
   Public SC_TERMINAL_SERVICES   As Boolean
   Public SC_FISCOCHKHABILITADO  As Boolean
   Public SC_NOMEDEPTO           As String
   
   Public NRC_Ano                As Integer
   Public NRC_Nro                As Integer
   Public NroVez                 As Integer
   
   Public FormMdi                As MDIForm
   Public RetemPisCof            As Boolean
   Public NomeMunic              As String
   Public UnidFederal            As String
   Public Msg                    As String
   
   '  Variaveis com Valores Constante
   Public Const Nkms             As String = "Mensagem do Sistema"
   Public Const MsgSys           As String = "Mensagem do Sistema"
'=======================================================================
' Variaveis para uso em Relatórios ( Impressão )
   Public ArqDir     As String ', NomeRel As String Esta variavel existe mod MakeFile
   Public Hora       As String
   Public Altura     As Integer
   Public Largura    As Integer
   Public Polegada   As Integer
   Public CtLin      As Integer
   Public CtPag      As Integer
   Public TotLin     As Integer
   Public Centimetro As Integer
   Public Tam        As Integer

  ' Variavies P/ Uso do FiscoCheck
   Public Chk_Situcao         As String
   Public Chk_InscEst         As String
   Public Chk_MsgErro         As String
   Public Chk_Bairro          As String
   Public Chk_Cep             As Long
   
   'Variaveis formulario pesquisa
   
   Public Glb_PesqNome      As String
   Public Glb_PesqCodigo   As String

Public Sub Marca(FRM As Form)
   On Error Resume Next
   If TypeOf FRM.ActiveControl Is TextBox Or _
      TypeOf FRM.ActiveControl Is MaskEdBox Then
      If FRM.ActiveControl.Enabled Then
         FRM.ActiveControl.SelStart = 0
         FRM.ActiveControl.SelLength = Len(Trim(FRM.ActiveControl))
      End If
   End If
End Sub
      
Public Sub ShowForm(FORMULARIO As Form, _
           Optional FixaForm As Boolean)
   FORMULARIO.Hide
   FORMULARIO.WindowState = 1
   FORMULARIO.WindowState = 0
   Set FormularioAtual = FORMULARIO
   
   FORMULARIO.Show IIf(Not FixaForm, 0, vbModal)
   Wcenter FORMULARIO
End Sub

Sub Wcenter(Janela As Form)
   If Janela.WindowState <> 1 Then
      If Janela.MDIChild = True Then
         If Janela.WindowState <> 2 Then
            Janela.Top = (FormMdi.ScaleHeight / 2) - (Janela.Height / 2)
            Janela.Left = (FormMdi.ScaleWidth / 2) - (Janela.Width / 2)
         End If
      End If
   End If
End Sub

Public Function SoNro(numero As String, _
                      Tecla As Integer, _
             Optional Decimais As Integer) As Integer
   Dim Posicao    As Integer
   Dim contador   As Integer
   Dim Tamanho    As Integer
   
   If Tecla = 46 Then Tecla = 44 ' Troca ponto por virgula

   Select Case Tecla
   Case vbKey0 To vbKey9
   Case 44
   Case vbKeyReturn, vbKeyTab
      SoNro = Tecla
      Exit Function
   Case vbKeyBack
      SoNro = Tecla
      Exit Function
   Case Else
      SoNro = 0
      Exit Function
   End Select
   
   Posicao = InStr(1, numero, ",", vbBinaryCompare)
     
   If Posicao > 0 Then
      If Chr(Tecla) = "," Then
         SoNro = 0
      Else
         contador = 0
         Tamanho = Len(numero)
         contador = Tamanho - Posicao
         
         If contador >= Decimais Then
            SoNro = 0
         Else
            SoNro = Tecla
         End If
      End If
   Else
      If Decimais = 0 Then
         If Chr(Tecla) = "," Then
            SoNro = 0
         Else
            SoNro = Tecla
         End If
      Else
         SoNro = Tecla
      End If
   End If
End Function

Public Sub CpoTexto(ByRef CodigoAsc As Integer, _
                 Optional Numerico As Boolean, _
                 Optional NroDecimal As Boolean, _
                 Optional Texto As String)
   If CodigoAsc = 8 Then Exit Sub
   If CodigoAsc = 13 Then Sendkeys "{TAB}": Exit Sub
   If Numerico = False And CodigoAsc = 39 Then
      CodigoAsc = 96
   ElseIf Numerico = True Then
      If NroDecimal = True Then
         If CodigoAsc = 46 Or CodigoAsc = 44 Then
            If InStr(1, Texto, ",") > 0 Then
               CodigoAsc = 0
            Else
               CodigoAsc = 44
               Exit Sub
            End If
         End If
      End If
      If CodigoAsc < 48 Or CodigoAsc > 57 Then
         If CodigoAsc <> 22 And CodigoAsc <> 3 Then CodigoAsc = 0
      End If
   End If
End Sub

Public Function Data_Pend_Pedido(cnpj As String) As Long
   If Mid(cnpj, 1, 8) = "61490561" Then
      Data_Pend_Pedido = Format("01/01/1900", "YYYYMMDD")
   Else
      Data_Pend_Pedido = Format(Date - 120, "YYYYMMDD") ' alterado de 60 p/ 120 CI - 03/01/2005
   End If
End Function

Public Function DESMASCARA_CHAPA(Chapa As String)
   DESMASCARA_CHAPA = Replace((Mid(Chapa, 1, 2)) & (Mid(Chapa, 4, 3)) & (Mid(Chapa, 8, 1)), "_", "")
End Function

Public Function TMASCDATA(vData) 'Tira máscara da data
   TMASCDATA = Mid(vData, 7, 4) & Mid(vData, 4, 2) & Mid(vData, 1, 2)
End Function
   
Public Function MASCDATA(vData) 'Põe máscara da data
   MASCDATA = Mid(vData, 7, 2) & "/" & Mid(vData, 5, 2) & "/" & Mid(vData, 1, 4)
End Function
   
Function VERIF_DATA(DataDig As String) As Boolean
   If IsDate(DataDig) = False Then
      VERIF_DATA = False
   ElseIf Mid(DataDig, 4, 2) > 12 Then
      VERIF_DATA = False
   Else
      VERIF_DATA = True
   End If
End Function

Function VerData(DataDig As String) As Boolean
   If DataDig = "__/__/____" Then
      VerData = False
   Else
      If IsDate(DataDig) = False Then
         VerData = False
      ElseIf Mid(DataDig, 4, 2) > 12 Then
         VerData = False
      Else
         VerData = True
      End If
   End If
End Function

Function EData(DataDigitada As String) As Boolean
   If Not IsDate(DataDigitada) Then
      EData = False
   ElseIf Val(Mid(DataDigitada, 7)) <= 1900 Then
      EData = False
   ElseIf Mid(DataDigitada, 4, 2) > 12 Then
      EData = False
   Else
      EData = True
   End If
End Function

Public Function DataNumerica(Data As String) As Long
   If Data = "__/__/____" Then
      DataNumerica = 0
   Else
      DataNumerica = CLng(Format(CDate(Data), "YYYYMMDD"))
   End If
End Function

Public Function DataDiaMesAno(Data As String) As String
   If Trim(Data) <> "0" Then
      DataDiaMesAno = Format(CDate(Format(Data, "0000/00/00")), "DD/MM/YYYY")
   End If
End Function

Public Function RedimensionamentoFlex(Gr As MSFlexGrid, TamanhoAnterior As Long)
   Dim I As Long
   For I = 0 To Gr.Cols - 1
      Gr.ColWidth(I) = ((Gr.Width / 100) * Gr.ColWidth(I) * 100 / TamanhoAnterior)
   Next
End Function

Function TruncaArredonda(Nro As Double, _
                         CasasDecimais As Integer, _
                Optional Arredonda As Boolean) As Variant
   Dim VlrFator As Double
   Dim Mascara As String
   
   If Not Arredonda Then
      VlrFator = CDbl("0," & String(CasasDecimais, "0") & "5")
   Else
      VlrFator = 0
   End If
   
   If CasasDecimais > 0 Then
      Mascara = "0." & String(CasasDecimais, "0")
   Else
      Mascara = "0"
   End If
   
   If Nro <> 0 Then
      If Nro > 0 Then
         TruncaArredonda = Format(Nro - VlrFator, Mascara)
      Else
         TruncaArredonda = Format(Nro + VlrFator, Mascara)
      End If
   Else
      TruncaArredonda = Nro
   End If
End Function

Public Function Tira_Caracter(Texto As String) As String
   Dim L       As Integer
   'Dim Tam     As Integer
   Dim Letra   As String
   Dim StrAux  As String
   
   Tam = Len(Texto)
   StrAux = Texto
   
   For L = 1 To Tam
      Letra = Mid(Texto, L, 1)
      Select Case Asc(UCase(Letra))
      Case 48 To 57  ' 0 a 9
      Case 65 To 90  ' A a Z
      Case Else
         StrAux = Replace(StrAux, Letra, "")
      End Select
   Next
   
   Tira_Caracter = StrAux
End Function

Public Function Tira_Traco(Texto As String, _
                           QualParteFica As Integer)
   Dim Posicao As Integer
   Posicao = InStr(1, Texto, "-", vbBinaryCompare)
   
   If Posicao > 0 Then
      If QualParteFica = 1 Then
         Tira_Traco = Trim(Mid(Texto, 1, Posicao - 1))
      Else
         Tira_Traco = Trim(Mid(Texto, Posicao + 1))
      End If
   Else
      Tira_Traco = Texto
   End If
End Function

Public Function TIRA_VIRGULA(vpreco) As String
   Dim wpos As Long
   wpos = InStr(1, vpreco, ",", vbBinaryCompare)
   
   If wpos > 0 Then
      'TIRA_VIRGULA = CDbl(Replace(vpreco, ",", "."))
      TIRA_VIRGULA = Mid(vpreco, 1, wpos - 1) & "." & _
                     Mid(vpreco, wpos + 1, Len(vpreco) - wpos)
   Else
      TIRA_VIRGULA = vpreco
   End If
End Function

Public Function UltimoDia(Mes As Integer, _
                          Ano As Integer) As Integer
   Dim Aux As Integer
   Aux = Ano Mod 4
   
   Select Case Mes
   Case 1, 3, 5, 7, 8, 10, 12
      UltimoDia = 31
   Case 4, 6, 9, 11
      UltimoDia = 30
   Case 2
      If Aux = 0 Then
         UltimoDia = 29
      Else
         UltimoDia = 28
      End If
   End Select
End Function

Function Cdblx(ByVal Texto As Variant) As Double
   On Error Resume Next
   
   If IsNumeric(Texto) Then
      If InStr(Texto, ",") = 0 Then Texto = Replace(Texto, ".", ",")
      
      Cdblx = CDbl(Trim(Texto))
   Else
      Cdblx = 0
   End If
End Function

Public Sub Maiusculo(Formu As Form)
   If TypeOf Formu.ActiveControl Is TextBox Then
      Formu.ActiveControl = UCase(Formu.ActiveControl)
      Formu.ActiveControl.SelStart = Len(Formu.ActiveControl)
   End If
End Sub

Public Function Formata_Cep(Cep As String) As String
   Formata_Cep = IIf(IsNumeric(Cep), Format(Format(Cep, "00000000"), "&&&&&-&&&"), "")
End Function

Public Function Formata_Cnpj(cnpj As String) As String
   Formata_Cnpj = IIf(IsNumeric(cnpj), Format(Format(cnpj, "00000000000000"), "&&.&&&.&&&/&&&&-&&"), "")
End Function

Public Function Check_Cep(UF As String, _
                          Cep As Long) As Boolean
                          
   Dim CepFormatado As String * 8
   
   CepFormatado = Format(Cep, "00000000")
   
   Select Case Trim(UF)
   Case "AC"
      Select Case Mid(CepFormatado, 1, 1)
      Case 6
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "AL"
      Select Case Mid(CepFormatado, 1, 1)
      Case 5
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "AM"
      Select Case Mid(CepFormatado, 1, 1)
      Case 6
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "AP"
      Select Case Mid(CepFormatado, 1, 1)
      Case 6
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "BA"
      Select Case Mid(CepFormatado, 1, 1)
      Case 4
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "CE"
      Select Case Mid(CepFormatado, 1, 1)
      Case 6
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "DF"
      Select Case Mid(CepFormatado, 1, 1)
      Case 7
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "ES"
      Select Case Mid(CepFormatado, 1, 1)
      Case 2
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "FN"
      Select Case Mid(CepFormatado, 1, 1)
      Case 5
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "GO"
      Select Case Mid(CepFormatado, 1, 1)
      Case 7
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "MA"
      Select Case Mid(CepFormatado, 1, 1)
      Case 6
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "MG"
      Select Case Mid(CepFormatado, 1, 1)
      Case 3
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "MS"
      Select Case Mid(CepFormatado, 1, 1)
      Case 7
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "MT"
      Select Case Mid(CepFormatado, 1, 1)
      Case 7
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "PA"
      Select Case Mid(CepFormatado, 1, 1)
      Case 6
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "PB"
      Select Case Mid(CepFormatado, 1, 1)
      Case 5
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "PE"
      Select Case Mid(CepFormatado, 1, 1)
      Case 5
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "PI"
      Select Case Mid(CepFormatado, 1, 1)
      Case 6
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "PR"
      Select Case Mid(CepFormatado, 1, 1)
      Case 8
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "RJ"
      Select Case Mid(CepFormatado, 1, 1)
      Case 2
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "RN"
      Select Case Mid(CepFormatado, 1, 1)
      Case 5
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "RO"
      Select Case Mid(CepFormatado, 1, 1)
      Case 7
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "RR"
      Select Case Mid(CepFormatado, 1, 1)
      Case 7
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "RS"
      Select Case Mid(CepFormatado, 1, 1)
      Case 9
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "SC"
      Select Case Mid(CepFormatado, 1, 1)
      Case 8
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "SE"
      Select Case Mid(CepFormatado, 1, 1)
      Case 4
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "SP"
      Select Case Mid(CepFormatado, 1, 1)
      Case 0, 1   ' Grande São Paulo, Interior
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "TO"
      Select Case Mid(CepFormatado, 1, 1)
      Case 7
         Check_Cep = True
      Case Else
         Check_Cep = False
      End Select
   Case "EX"
      Check_Cep = IIf(Cep > 0, False, True)
   Case Else
      Check_Cep = False
   End Select
End Function

Public Function CONVERTE_ASPAS(Texto As String) As String
   Dim Posic As Integer
   Posic = InStr(1, Texto, Chr(39), 1)
   
   If Posic = 0 Then
      CONVERTE_ASPAS = Chr(39) & Trim(Texto) & Chr(39)
   Else
      CONVERTE_ASPAS = Chr(34) & Trim(Texto) & Chr(34)
   End If
End Function

Function PoeAspas(ByVal Texto As String) As String
   Dim Pos As Integer
      
   If Trim(Texto) <> "" Then
      Pos = InStr(1, Texto, Chr(39), vbBinaryCompare)
      
      If Pos > 0 Then
         PoeAspas = Chr(34) & Trim(Texto) & Chr(34)
      Else
         Pos = InStr(1, Texto, Chr(34), vbBinaryCompare)
         
         If Pos > 0 Then
            PoeAspas = Chr(39) & Trim(Texto) & Chr(39)
         Else
            PoeAspas = Chr(34) & Trim(Texto) & Chr(34)
         End If
      End If
   Else
      PoeAspas = Chr(39) & " " & Chr(39)
   End If
End Function

Public Function MASCCGC(vcgc) 'Põe máscara no cnpj
   MASCCGC = Mid(vcgc, 1, 2) & "." & Mid(vcgc, 3, 3) & "." _
           & Mid(vcgc, 6, 3) & "/" & Mid(vcgc, 9, 4) & "-" _
           & Mid(vcgc, 13, 2)
End Function

Public Function MascCnpj(cnpj As String)
   MascCnpj = IIf(IsNumeric(cnpj), Format(cnpj, "&&.&&&.&&&/&&&&-&&"), cnpj)
End Function

Public Function NKCD(Vlr) As Double
   NKCD = Cdblx(IIf(IsNull(Trim(Vlr)), 0, Vlr))
End Function

Public Function NKCS(Vlr) As String
   NKCS = IIf(Trim(Vlr) = "", " ", Vlr)
End Function

Function NFormat(Vlr As Variant, _
                 Mascara As String) As String
   Dim TamMasc As Integer
   Dim TamCpo  As Integer
   Dim PosVirg As Integer
   Dim Idx     As Integer
   Dim MascAux As String
   
   MascAux = Replace(Mascara, ",", "@")
   MascAux = Replace(MascAux, ".", ",")
   MascAux = Replace(MascAux, "@", ".")
   
   MascAux = Replace(MascAux, "Z", "#")
   MascAux = Replace(MascAux, "B", "#")
   MascAux = Replace(MascAux, "9", "#")
   MascAux = Replace(MascAux, "0", "#")
   
   TamMasc = Len(MascAux)
   PosVirg = InStr(MascAux, ".")
   
   If Vlr = 0 And Mid(Mascara, 1, 1) = "B" Then
      NFormat = Space(Len(Mascara))
      Exit Function
   End If
   
   If PosVirg > 0 Then
      Mid(MascAux, PosVirg - 1, 1) = "0"
      For Idx = PosVirg + 1 To Len(MascAux)
         Mid(MascAux, Idx, 1) = "0"
      Next
   Else
      Mid(MascAux, Len(MascAux), 1) = "0"
   End If
   
   TamCpo = Len(Format(Trim(Vlr), MascAux))
   
   If TamMasc >= TamCpo Then
      NFormat = Space(TamMasc - TamCpo) & Format(Vlr, MascAux)
   Else
      NFormat = String(TamMasc, "*")
   End If
End Function

Public Function LeituraParametro() As Boolean
   VDataServ = Date
  ' SC_MAQNAME = Cnn.MaquinaLocal
   LeituraParametro = True
End Function

Public Function FirstDay(dtData As Date) As Date
    FirstDay = CDate("01/" & Month(dtData) & "/" & Year(dtData))
End Function

Public Function LastDay(dtData As Date) As Date
    On Error Resume Next
    
    Dim strUltDia As String
    Dim iDia As Integer
    
    iDia = 32
    Do
        iDia = iDia - 1
        strUltDia = CStr(iDia) & "/" & CStr(Month(dtData)) & "/" & CStr(Year(dtData))
    Loop While Not IsDate(strUltDia)
    
    LastDay = CDate(strUltDia)
End Function

Public Sub Sendkeys(Text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(Text), wait
      Set WshShell = Nothing
End Sub

Public Function Crypt(Text As String) As String

   Dim strTempChar As String

   For I = 1 To Len(Text)
   
      If Asc(Mid$(Text, I, 1)) < 128 Then
         strTempChar = Asc(Mid$(Text, I, 1)) + 128
      ElseIf Asc(Mid$(Text, I, 1)) > 128 Then
         strTempChar = Asc(Mid$(Text, I, 1)) - 128
      End If
      
      Mid$(Text, I, 1) = Chr(strTempChar)
   
   Next I
   
   Crypt = Text

End Function

Public Function CARREGA_COMBOS(PROJETO As String, _
                               FORMULARIO As String, _
                               Optional Tipo As String)
                               
   strCombo = "SELECT * FROM TIPOCOMBO" & vbCr
   strCombo = strCombo & "WHERE PROJETO    = " & PoeAspas(PROJETO) & vbCr
   strCombo = strCombo & "  AND FORMULARIO = " & PoeAspas(FORMULARIO) & vbCr
   If Trim(Tipo) <> "" Then strCombo = strCombo & "  AND TIPO       = " & PoeAspas(Tipo) & vbCr
   strCombo = strCombo & "ORDER BY PROJETO,FORMULARIO,TIPO,DESCRICAO"
   CMySql.Consulta strCombo, RsTipoCombo
                               
End Function

Public Function VARIAVEIS_SISTEMA(USUARIO As String)

   CmdSql = "SELECT * FROM ACESSO" & vbCr
   CmdSql = CmdSql & "WHERE USUARIO = " & PoeAspas(USUARIO)
   CMySql.Consulta CmdSql, RsConsulta
   
   If Not RsConsulta.EOF Then
      sysSistema = App.EXEName
      sysUsuario = Trim(RsConsulta("USUARIO"))
      sysPerfil = Trim(RsConsulta("PERFIL"))
   Else
      sysSistema = App.EXEName
      sysUsuario = Trim(Rs("USUARIO"))
      sysPerfil = Trim(Rs("PERFIL"))
   End If
End Function

Public Function IsValidEmail(ByVal email As String) As Boolean
    Dim emailRegex As Object
    Set emailRegex = CreateObject("VBScript.RegExp")
    
    ' Define a expressão regular para validar o e-mail '
    emailRegex.Pattern = "^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$"
    emailRegex.IgnoreCase = True
    
    ' Testa se o e-mail passado como parâmetro é válido '
    IsValidEmail = emailRegex.Test(email)
End Function

Public Function IsValidCPF(ByVal cpf As String) As Boolean
    Dim I As Integer
    Dim sum1 As Integer
    Dim sum2 As Integer
    Dim digit1 As Integer
    Dim digit2 As Integer
    
    ' Remove os caracteres de formatação do CPF '
    cpf = Replace(Replace(Replace(Replace(cpf, ".", ""), "-", ""), "/", ""), " ", "")
    
    ' Verifica se o CPF tem 11 dígitos '
    If Len(cpf) <> 11 Then
        IsValidCPF = False
        Exit Function
    End If
    
    ' Verifica se todos os dígitos do CPF são iguais '
    If cpf = String(11, Mid(cpf, 1, 1)) Then
        IsValidCPF = False
        Exit Function
    End If
    
    ' Calcula o primeiro dígito verificador '
    sum1 = 0
    For I = 1 To 9
        sum1 = sum1 + (Val(Mid(cpf, I, 1)) * (11 - I))
    Next
    digit1 = (11 - (sum1 Mod 11)) Mod 10
    
    ' Calcula o segundo dígito verificador '
    sum2 = 0
    For I = 1 To 9
        sum2 = sum2 + (Val(Mid(cpf, I, 1)) * (12 - I))
    Next
    sum2 = sum2 + (digit1 * 2)
    digit2 = (11 - (sum2 Mod 11)) Mod 10
    
    ' Verifica se os dígitos verificadores estão corretos '
    If (Val(Mid(cpf, 10, 1)) = digit1) And (Val(Mid(cpf, 11, 1)) = digit2) Then
        IsValidCPF = True
    Else
        IsValidCPF = False
    End If
End Function

Public Function IsValidCNPJ(ByVal cnpj As String) As Boolean
    Dim I As Integer
    Dim sum1 As Integer
    Dim sum2 As Integer
    Dim digit1 As Integer
    Dim digit2 As Integer
    
    ' Remove os caracteres de formatação do CNPJ '
    cnpj = Replace(Replace(Replace(Replace(cnpj, ".", ""), "-", ""), "/", ""), " ", "")
    
    ' Verifica se o CNPJ tem 14 dígitos '
    If Len(cnpj) <> 14 Then
        IsValidCNPJ = False
        Exit Function
    End If
    
    ' Verifica se todos os dígitos do CNPJ são iguais '
    If cnpj = String(14, Mid(cnpj, 1, 1)) Then
        IsValidCNPJ = False
        Exit Function
    End If
    
    ' Calcula o primeiro dígito verificador '
    sum1 = 0
    For I = 1 To 12
        sum1 = sum1 + (Val(Mid(cnpj, I, 1)) * (15 - I))
    Next
    digit1 = IIf(sum1 Mod 11 < 2, 0, 11 - (sum1 Mod 11))
    
    ' Calcula o segundo dígito verificador '
    sum2 = 0
    For I = 1 To 13
        sum2 = sum2 + (Val(Mid(cnpj, I, 1)) * (16 - I))
    Next
    digit2 = IIf(sum2 Mod 11 < 2, 0, 11 - (sum2 Mod 11))
    
    ' Verifica se os dígitos verificadores estão corretos '
    If (Val(Mid(cnpj, 13, 1)) = digit1) And (Val(Mid(cnpj, 14, 1)) = digit2) Then
        IsValidCNPJ = True
    Else
        IsValidCNPJ = False
    End If
End Function

Public Function SoNumeros(Obj As Object, Keyasc As Integer)
   If Not ((Keyasc >= Asc(0) And Keyasc <= Asc(9)) Or Keyasc = 8 Or Keyasc = Asc(".") Or Keyasc = Asc(",") Or Keyasc = Asc("-") Or Keyasc = Asc("/")) Then
      Keyasc = 0
      Exit Function
   End If
End Function
