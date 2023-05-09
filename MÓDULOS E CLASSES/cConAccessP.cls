VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Conn As New ADODB.Connection
Private sBanco As String, sRet As Long, sWinDir As String, sTempPath As String
Private sUsuario As String, sComputador As String
Private sPathLog As String, sPathLogErro As String, sPathConfINI As String
Private sPathLocalLog As String
Private sPathBanco As String
Private EmTransacao As Boolean
Private nRegistrosAfetados As Double
Private sPwd As String

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
      
   Executa = True
   Exit Function

Erro_Executa:
   sMsgError = Err.Description
   
   If EmTransacao Then Trans_RollBack
     If FinalizaOnError Then
      MsgBox "Esse aplicativo será finalizado!", vbCritical, "Atenção"
      End
   Else
      Err.Clear
      Executa = False
   End If
End Function

Public Function Consulta(StringSql As String, pRst As ADODB.Recordset, Optional Acesso As CursorTypeEnum, Optional ModeLock As LockTypeEnum)
   On Error GoTo Saida
   If ModeLock = 0 Then ModeLock = adLockReadOnly
   If pRst.State = 1 Then pRst.Close
   pRst.Open StringSql, Conn, Acesso, ModeLock
   Exit Function
Saida:
   MsgBox "Ocorreu um erro e o programa será fechado!"
   End
End Function

Private Function Conecta() As ADODB.Connection
   sUsuario = String$(200, Chr$(0))
   GetUserName sUsuario, 200
   sUsuario = Left$(sUsuario, InStr(sUsuario, Chr$(0)) - 1)
   
   sComputador = String(20, Chr(0))
   GetComputerName sComputador, 20
   sComputador = Left$(sComputador, InStr(sComputador, Chr$(0)) - 1)
   
   sWinDir = String(200, Chr(0))
   sRet = GetWindowsDirectory(sWinDir, 200)
   sWinDir = Left$(sWinDir, sRet)
   sPathLogErro = App.Path & "\LOGS"
   sPathLog = App.Path & "\LOGS"
   sTempPath = String$(200, Chr(0))
   sRet = GetTempPath(200, sTempPath)
   sTempPath = Left$(sTempPath, sRet)
   
   Conn.Provider = "Microsoft.jet.oledb.4.0"
   Conn.Mode = adModeReadWrite
   Conn.CursorLocation = adUseServer
   Conn.Open sPathBanco    '& Conn.ConnectionString
   
End Function

Private Sub Class_Initialize()
   Set Conn = CreateObject("ADODB.Connection")
End Sub

Private Sub Class_Terminate()
   If Conn.State = 1 Then Conn.Close
   Set Conn = Nothing
End Sub

Public Property Get Trans_EmTransacao() As Boolean
   Trans_EmTransacao = EmTransacao
End Property

Public Property Get Usuario() As String
   Usuario = sUsuario
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

Public Function Inicializa(PathBanco As String, Optional Pwd As String)
   sPathBanco = PathBanco
   Set Conn = CreateObject("ADODB.Connection")
   sPwd = Pwd
   Conecta
End Function

Public Property Get RegistrosAfetados() As Double
   RegistrosAfetados = nRegistrosAfetados
End Property

Public Function GravaINI(Programa As String, Chave As String, Valor As String, Caminho As String)
   WritePrivateProfileString Programa, Chave, Valor, Caminho
End Function

Public Function LerINI(Programa As String, Chave As String, Caminho As String) As String
 Dim nRet As Long, sRet As String * 50
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

