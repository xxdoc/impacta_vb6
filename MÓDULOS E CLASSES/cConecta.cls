VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
   Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
   Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
   Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
   Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
   Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
   Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
   Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
   
   Const SW_SHOWNORMAL = 1
   
   Private CnMy               As New cConMysql
   Private CnSqlSv            As New CConSql
   
   Private Conn               As Object
   Private sBanco             As String
   Private sRet               As Long
   Private sWinDir            As String
   Private sTempPath          As String
   Private sUsuario           As String
   Private sComputador        As String
   Private sPathLog           As String
   Private sPathLogErro       As String
   Private sPathConfINI       As String
   Private sPathLocalLog      As String
   Private EmTransacao        As Boolean
   Private nRegistrosAfetados As Double
   Private ConfConect         As RegConect
   Private sCaminhoDoSistema  As String
   Private sNomeServidor      As String
   Private sCodEmpresa        As String
   Private sUltimoSQL         As String
   Private vsql               As String
   Private sNomeTabelaLog     As String
   
   Private Type RegConect
      Driver      As String * 100
      Odbc        As String * 5
      Dsn         As String * 30
      Pwd         As String * 20
      DataBase    As String * 20
      Maquina     As String * 25
      Servico     As String * 25
      Servidor    As String * 25
      Protocolo   As String * 15
      USUARIO     As String * 20
   End Type

Public Sub Trans_Begin()
   If Not EmTransacao Then
      EmTransacao = True
      Executa "SET LOCK MODE TO WAIT;BEGIN WORK"
   End If
End Sub

Public Sub Trans_Commit()
   If EmTransacao Then
      EmTransacao = False
      Executa "COMMIT WORK"
   End If
End Sub

Public Sub Trans_RollBack()
   If EmTransacao Then
      EmTransacao = False
      Executa "ROLLBACK WORK"
   End If
End Sub

Public Function Executa(Sql As String, _
               Optional NaoFinalizaOnError As Boolean) As Boolean
   Dim sMsgError As String
   On Error GoTo Erro_Executa
   Conn.Execute Sql, nRegistrosAfetados
   'Log_Execute Sql
   sUltimoSQL = Sql
   Executa = True
   Exit Function

Erro_Executa:
   sMsgError = Err.Description
   
   If EmTransacao Then Trans_RollBack
   
   Log_Erro sMsgError, Sql, NaoFinalizaOnError

   If NaoFinalizaOnError Then
      Err.Clear
   Else
      MsgBox "Esse aplicativo será finalizado!", vbCritical, "Atenção"
      End
   End If
End Function

Public Function Consulta(StringSql As String, _
                         pRst As ADODB.Recordset, _
                Optional Acesso As CursorTypeEnum, _
                Optional ModeLock As LockTypeEnum)
   On Error GoTo Saida
   If ModeLock = 0 Then ModeLock = adLockReadOnly
   If pRst.State = 1 Then pRst.Close
   pRst.Open StringSql, Conn, Acesso, ModeLock
   nRegistrosAfetados = pRst.RecordCount
   Exit Function
Saida:
   Log_Erro Err.Description, StringSql
End Function

Private Function Conecta() As ADODB.Connection
   Dim STrcn      As String
   Dim RsMysql    As New ADODB.Recordset
   Dim Pos        As Integer
   Dim CodEmpresa As Integer
   
   sUsuario = String$(200, Chr$(0))
   GetUserName sUsuario, 200
   sUsuario = Left$(sUsuario, InStr(sUsuario, Chr$(0)) - 1)
   
   sComputador = String(20, Chr(0))
   GetComputerName sComputador, 20
   sComputador = Left$(sComputador, InStr(sComputador, Chr$(0)) - 1)
   
   sWinDir = String(200, Chr(0))
   sRet = GetWindowsDirectory(sWinDir, 200)
   sWinDir = Left$(sWinDir, sRet)
   'sPathConfINI = CMYSQL.NomeServidor & "\RELSCRIPTS\ConfiguraBanco\cfg336.ini"
   sPathLocalLog = sWinDir & "\Cinpal.ini"
   
   sTempPath = String$(200, Chr(0))
   sRet = GetTempPath(200, sTempPath)
   sTempPath = Left$(sTempPath, sRet)
   
   If Dir(sPathLocalLog) = "" Then
      WritePrivateProfileString "Geral", "Banco", "Prod", sPathLocalLog
   End If
      
   sBanco = LerINI("Geral", "Banco", sPathLocalLog)
   sCaminhoDoSistema = App.Path
      
   CodEmpresa = Val(PegarEmpresa(sCaminhoDoSistema, _
                                 InStr(1, Command, "EMPR=")))
   
   If CodEmpresa = 0 Then
      MsgBox "Codigo da empresa não definido", vbCritical, "Classe cConecta"
      End
   End If
   
   vsql = "SELECT * FROM ACESSO,CONEXAOBD,EMPRESA,NOMESERVIDOR WHERE ACESSO.IDEMPRESA =  " & _
      CodEmpresa & " AND ACESSO.IDAMBIENTE = '" & IIf(Trim(sBanco) = "Prod", "P", "D") & _
      "' AND CONEXAOBD.IDCONEXAO = ACESSO.IDCONEXAO AND EMPRESA.IDEMPRESA = ACESSO.IDEMPRESA"
   
   'CnMy.Consulta vSql, RsMysql
   CnSqlSv.Consulta vsql, RsMysql
      
   If RsMysql.EOF Then
      MsgBox "Parâmetros de conexão com banco de dados não encontrado. " & Chr(13) & "Entre em contato com o CPD IMEDIATAMENTE!", vbExclamation, "Atenção"
      End
   End If
      
   sNomeServidor = Trim(RsMysql("SERVIDORAPP"))
   ConfConect.Driver = RsMysql("DRIVER")
   ConfConect.Dsn = RsMysql("DSN")
   ConfConect.USUARIO = RsMysql("USUARIO")
   ConfConect.Pwd = RsMysql("SENHA")
   ConfConect.DataBase = RsMysql("BANCO")
   ConfConect.Maquina = RsMysql("HOST")
   
   ConfConect.Servico = RsMysql("SERVICO")
   ConfConect.Servidor = RsMysql("NOMESERVICO")
   ConfConect.Protocolo = RsMysql("protocolo")
  
   STrcn = "driver=" & Trim(ConfConect.Driver) & _
      ";UID=" & Trim(ConfConect.USUARIO) & _
      ";PWD=" & Trim(ConfConect.Pwd) & _
      ";database=" & Trim(ConfConect.DataBase) & _
      ";host=" & Trim(ConfConect.Maquina) & _
      ";SERV=" & Trim(ConfConect.Servico) & _
      ";SRVR=" & Trim(ConfConect.Servidor) & _
      ";PRO=" & Trim(ConfConect.Protocolo) & _
      ";DB_LOCALE=EN_US.8859-1"

   Conn.Open STrcn

   Conn.CursorLocation = adUseClient

   vsql = "set lock mode to wait"
   Conn.Execute vsql
   'ChecaTabeladeLogs
End Function

Private Sub Class_Initialize()
   Set Conn = CreateObject("ADODB.Connection")
   LogEntrada
End Sub

Private Sub Class_Terminate()
   If Conn.State = 1 Then Conn.Close
   Set Conn = Nothing
End Sub

Public Property Get Banco() As String
   Banco = IIf(sBanco = "Prod", "Produção", "Desenvolvimento")
End Property

Public Property Get Trans_EmTransacao() As Boolean
   Trans_EmTransacao = EmTransacao
End Property

Public Property Get Contexto() As String
   Contexto = IIf(sBanco = "Prod", "Prod", "CPD")
End Property

Public Property Get USUARIO() As String
   USUARIO = sUsuario
End Property

Public Property Get MaquinaLocal() As String
   MaquinaLocal = sComputador
End Property

Public Property Get DirWindows() As String
   DirWindows = sWinDir
End Property

Public Property Get DirTemp() As String
   DirTemp = sTempPath
End Property

Private Sub Log_Execute(Sql As String)
   On Error GoTo SaidaLogExec
        
   vsql = "INSERT INTO " & sNomeTabelaLog & "(TIPO,MODULO,DATAHORA,COMPUTADOR,USUARIO,COMANDO,ERRO) VALUES ('X','"
   vsql = vsql & App.EXEName & "','" & ServerTime & "','" & sComputador & "','" & sUsuario & "','" & Replace(Replace(Sql, ";", "@"), "'", "`") & "',' ')"
   CnMy.Executa vsql, False
   Exit Sub
SaidaLogExec:
   Log_Erro Err.Description, Sql
   Err.Clear
   Resume Next
End Sub

Private Sub Log_Erro(Descricao As String, _
            Optional Sql As String, _
            Optional NaoFinalizaOnError As Boolean)
   On Error GoTo SaidaLogErro
   
   vsql = "INSERT INTO " & sNomeTabelaLog & "(TIPO,MODULO,DATAHORA,COMPUTADOR,USUARIO,COMANDO,ERRO) VALUES ('E','"
   vsql = vsql & App.EXEName & "','" & ServerTime & "','" & sComputador & "','" & sUsuario & "','" & Replace(Replace(Sql, ";", "@"), "'", "`") & "','" & Descricao & "')"
   CnMy.Executa vsql, False, True
   'Exit Sub
   
   MsgBox "Ocorreu um erro de execução. " & vbCrLf & vbCrLf & _
          "Aplicativo: " & App.EXEName & vbCrLf & vbCrLf & _
          "Descrição: " & vbCrLf & Trim(Descricao) & vbCrLf & _
          "SQL:" & vbCrLf & Sql & vbCrLf & vbCrLf & _
          String(35, Chr(151)) & vbCrLf & _
          "Um arquivo de LOG referente a este evento foi criado", vbCritical, "Atenção"
   
   If NaoFinalizaOnError Then
      Err.Clear
   Else
      MsgBox "Esse aplicativo será finalizado!", vbCritical, "Atenção"
      End
   End If
   
   Exit Sub
SaidaLogErro:
   MsgBox "Ocorreu um erro gravando o arquivo 'LOG de ERRO!'. Por favor, contate o Depto. de Tecnologia da Informação" & vbCrLf & _
         " Descricao:" & Err.Description, vbCritical
   MsgBox "Esse aplicativo será finalizado", vbCritical, "Atenção"
   Resume Next
   End
End Sub



Public Property Get RegistrosAfetados() As Double
   RegistrosAfetados = nRegistrosAfetados
End Property

Public Function ProgramaExterno(CaminhoCompleto As String)
    ShellExecute 1045, vbNullString, CaminhoCompleto, vbNullString, "C:\", 3
End Function

Public Property Get CaminhoDoSistema() As String
   CaminhoDoSistema = sCaminhoDoSistema
End Property

Public Property Let CaminhoDoSistema(sCam As String)
   sCaminhoDoSistema = sCam
End Property

Public Property Get NomeServidor() As String
   NomeServidor = sNomeServidor
End Property

Public Property Let NomeServidor(sNomServ As String)
   sNomeServidor = sNomServ
End Property

Public Function LerINI(Programa As String, _
                       Chave As String, _
                       Caminho As String) As String
 Dim nRet As Long, sRet As String * 255
 Dim sLinha As String
   On Error GoTo MsgErroINI
   If Dir(Caminho) = "" Then
      MsgBox "O arquivo: " & Caminho & " não pôde ser encontrado!", vbExclamation, "Atenção"
      LerINI = " "
      End
      Exit Function
   End If
   
   Open Caminho For Input As #1
   If EOF(1) Then: Close #1: Exit Function
      Do
         Line Input #1, sLinha
         If UCase(Trim(sLinha)) = ("[" & UCase(Trim(Programa)) & "]") Then
            Do
               Line Input #1, sLinha
               If UCase(Trim(Mid(sLinha, 1, Len(Chave)))) = UCase(Trim(Chave)) Then
                  LerINI = Trim(Mid(sLinha, InStr(1, sLinha, "=", vbTextCompare) + 1))
                  Close #1
                  Exit Function
               End If
             Loop Until EOF(1)
         End If
      Loop Until EOF(1)
   Close #1
   LerINI = "0"

   Exit Function
MsgErroINI:
   MsgBox "Erro ao ler arquivo *.ini!. Contate o CPD", vbCritical, "Atenção"
   Err.Clear
   Exit Function
End Function

Public Function GravaINI(Programa As String, _
                         Chave As String, _
                         Valor As String, _
                         Caminho As String)
   WritePrivateProfileString Programa, Chave, Valor, Caminho
End Function

Public Property Get UltimoComandoSQL() As String
   UltimoComandoSQL = sUltimoSQL
End Property

Public Function ServerTime(Optional DataServ As Date, _
                           Optional HoraServ As String) As String
   Dim pSrvTime As New ADODB.Recordset
   
   vsql = "SELECT DISTINCT CURRENT YEAR TO SECOND FROM CPAIS"
   pSrvTime.Open vsql, Conn, adOpenForwardOnly, adLockReadOnly
   DataServ = Format(pSrvTime(0), "DD/MM/YYYY")
   HoraServ = Format(pSrvTime(0), "HH:MM:SS")
   ServerTime = pSrvTime(0)
End Function

Private Sub LogEntrada()
   Dim sLogName As String
   On Error GoTo errLogEntrada
      
   Exit Sub
      
errLogEntrada:
   Err.Clear
   Resume Next
End Sub

Private Sub ConectaMySql()
   CnMy.Inicializa "config", "xconexao", "192.168.1.9", "cinpal"
   CnMy.Executa "USE CONFIG;"
End Sub

Private Sub ConectaSqlServer()
   
   CnSqlSv.Executa "USE CONFIG;"
End Sub

Public Function Inicializa(nCdEmpresa As Integer)
End Function

Private Sub ChecaTabeladeLogs()
   sNomeTabelaLog = "LOG" & Year(ServerTime) & Format(Month(ServerTime), "00")
   
   CnMy.Executa "CREATE TABLE IF NOT EXISTS " & sNomeTabelaLog & "(Tipo varchar (01) NOT NULL," & _
                "Modulo     varchar (50) NOT NULL ," & _
                "DataHora   varchar (20) NOT NULL," & _
                "seq int AUTO_INCREMENT," & _
                "Computador varchar (50)," & _
                "Usuario varchar (50)," & _
                "Comando longtext," & _
                "Erro longtext," & _
                "Primary Key(tipo, Modulo, DataHora, seq)" & _
                ") TYPE=MyISAM"
End Sub

Public Sub CursorLocation(AdLocation As ADODB.CursorLocationEnum)
   Conn.CursorLocation = AdLocation
End Sub

Function PegarEmpresa(Caminho As String, _
                      Inicio As Long) As String
   Dim StringParcial As String
   Dim I             As Integer
   
   StringParcial = ""
   
   If InStr(1, UCase(Caminho), "\SC\") > 0 Then
      If Mid(Command, Inicio, 5) <> "EMPR=" Then Exit Function
      
      Inicio = Inicio + 5
      
      For I = Inicio To Len(Command)
         StringParcial = StringParcial & Mid(Command, I, 1)
         If Mid(Command, I, 1) = "," Then
            StringParcial = Left(StringParcial, Len(StringParcial) - 1)
            Exit For
         End If
      Next
      
      PegarEmpresa = StringParcial
   Else
      PegarEmpresa = 1
   End If
End Function
