VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConMysql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
   Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                           ByVal lpKeyName As Any, ByVal lpDefault As String, _
                           ByVal lpReturnedString As String, _
                           ByVal nSize As Long, ByVal lpFileName As String) As Long
                           
   Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                           ByVal lpKeyName As Any, _
                           ByVal lpString As Any, _
                           ByVal lpFileName As String) As Long
                           
   Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                           ByVal lpBuffer As String) As Long
                           
   Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
                           ByVal nSize As Long) As Long
                           
   Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, _
                           nSize As Long) As Long
                           
   Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, _
                           nSize As Long) As Long
                           
   Private Conn         As Object
   Private sRet         As Long
   Private sBanco       As String
   Private sWinDir      As String
   Private sTempPath    As String
   Private sUsuario     As String
   Private sComputador  As String
   Private sPathLog     As String
   Private sPathLogErro As String
   Private sPathConfINI As String
   Private sPathLocalLog As String
   Private sUsuarioMysql As String
   Private sNomeBanco    As String
   Private sPwd          As String
   Private sHst          As String
   Private sServidor     As String
   Private nRegistrosAfetados As Double
   Private EmTransacao   As Boolean
   
   'Private WithEvents conn As ADODB.Connection
   'com withevents voce pode codificar os eventos para a conexao com
   'recordset e tambem pode localizar o objeto na lista de codigo
   Private RsMysql As ADODB.Recordset
Attribute RsMysql.VB_VarHelpID = -1
   'rsMySQL é o recordset que sera usado na conexao
   Dim mblnAddMode As Boolean
   'Private WithEvents Teste As ADODB.Recordset
   

Public Sub Trans_Begin()
   If Not EmTransacao Then
      EmTransacao = True
      Conn.BeginTrans
   End If
End Sub

Public Sub Trans_Commit()
   If EmTransacao Then
      EmTransacao = False
      Conn.CommitTrans
   End If
End Sub

Public Sub Trans_RollBack()
   If EmTransacao Then
      EmTransacao = False
      Conn.RollbackTrans
   End If
End Sub

Public Function Executa(Sql As String, Optional FinalizaOnError As Boolean, Optional SemLog As Boolean) As Boolean
   Dim sMsgError As String
   On Error GoTo Erro_Executa
   Conn.Execute Sql, nRegistrosAfetados
   
   '----------------------------------------------------------------
   '' AS ROTINAS DE GRAVAÇÃO DE LOG DESTA CLASSE FORAM DESUTILIZADAS
   '' SENDO SOMENTE NECESSÁRIA AS DA CLASSE CONECTA
   'If SemLog Then Log_Execute Sql
   '----------------------------------------------------------------
   
   Executa = True
   Exit Function

Erro_Executa:
   sMsgError = Err.Description
   
   If EmTransacao Then Trans_RollBack
   
'   Log_Erro sMsgError, Sql

   If FinalizaOnError Then
      MsgBox "Esse aplicativo será finalizado!", vbCritical, "Atenção"
      End
   Else
      Err.Clear
      Executa = False
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
   'nRegistrosAfetados = pRst.RecordCount
   Exit Function
Saida:
   Log_Erro Err.Description, StringSql
End Function

Private Function Conecta() As ADODB.Connection
   Dim strConnect As String      'ESTÁ É A STRING DE CONEXÃO DEVERÁ CONTER A INFORMAÇÃO SOBRE O PROVEDOR E O CAMINHO DO BD.
   Dim sCaminhoDoSistema As String
   Dim vcontexto As String
   
   strConnect = "driver={MySQL ODBC 3.51 DRIVER};server=" & sServidor & ";uid=" & sUsuarioMysql & ";pwd=" & sPwd & ";database=" & sNomeBanco
   Set Conn = New ADODB.Connection
   Conn.CursorLocation = adUseServer
   Conn.Open strConnect
End Function

'Private Sub Class_Terminate()
'   If Conn.State = 1 Then Conn.Close
'   Set Conn = Nothing
'End Sub

Public Property Get Trans_EmTransacao() As Boolean
   Trans_EmTransacao = EmTransacao
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

Private Sub Log_Erro(Descricao As String, _
                     Optional Sql As String)
   On Error GoTo SaidaLogErro

   MsgBox "Ocorreu um erro de execução. " & vbCrLf & vbCrLf & _
          "Aplicativo: " & App.EXEName & vbCrLf & vbCrLf & _
          "Descrição: " & vbCrLf & Trim(Descricao) & vbCrLf & _
          "SQL:" & vbCrLf & Sql & vbCrLf & vbCrLf & _
          String(35, Chr(151)) & vbCrLf & _
          "cConMySql" & vbCrLf & _
          "Esse aplicativo será finalizado, entre em contato com o CPD.", _
          vbCritical, "Atenção"
   End
   Exit Sub
SaidaLogErro:

End Sub

Public Function Inicializa(NomeBanco As String, _
                           Optional Pwd As String, _
                           Optional IPServidor As String, _
                           Optional NomeUsuario As String)
   sNomeBanco = NomeBanco
   sWinDir = String(200, Chr(0))
   sRet = GetWindowsDirectory(sWinDir, 200)
   sWinDir = Left$(sWinDir, sRet)
   Set Conn = CreateObject("ADODB.Connection")
   sPwd = Pwd
   sUsuarioMysql = IIf(Trim(NomeUsuario) = "", "", NomeUsuario)
   sServidor = IIf(Trim(IPServidor) <> "", IPServidor, "localhost")
   Conecta
End Function

Public Property Get RegistrosAfetados() As Double
   RegistrosAfetados = nRegistrosAfetados
End Property

Public Function GravaINI(Programa As String, _
                         Chave As String, _
                         Valor As String, _
                         Caminho As String)
   WritePrivateProfileString Programa, Chave, Valor, Caminho
End Function

Public Function LerINI(Programa As String, _
                        Chave As String, _
                        Caminho As String) As String
   Dim nRet As Long
   Dim sRet As String * 50
   
   On Error GoTo MsgErroINI
   
   If Dir(Caminho) = "" Then
      MsgBox "O arquivo: " & Caminho & " não pôde ser encontrado!", vbExclamation, "Atenção"
      LerINI = " "
      Exit Function
   End If
   
   nRet = GetPrivateProfileString(Programa, Chave, " ", sRet, 50, Caminho)
   LerINI = Trim(Left$(sRet, nRet))
   Exit Function
MsgErroINI:
   MsgBox "Erro ao ler arquivo *.ini!. Contate o CPD", vbCritical, "Atenção"
   Err.Clear
   Exit Function
End Function



