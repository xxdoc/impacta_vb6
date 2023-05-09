Attribute VB_Name = "MDLMsgBox"
Option Explicit

' finaliza os temporizadores
Public Declare Function KillTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long) As Long

' busca o cabeçalho da Msgbox
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

' crea os os timers
Public Declare Function SetTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long
    
' escreve em um caption do Msgbox
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" ( _
    ByVal hwnd As Long, _
    ByVal lpString As String) As Long
    
' cria el Msgbox a partir de um  Handle
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal WMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

'Constante para sendMessage
Const SC_CLOSE = &HF060&
Const WM_SYSCOMMAND = &H112

Public Segundos As Integer ' segundos de duración
Private hMessageBox As Long
Public flag As Boolean


' Funcão para os timers
'''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub TimerProc(ByVal hwnd As Long, _
                     ByVal uMsg As Long, _
                     ByVal idEvent As Long, _
                     ByVal dwTime As Long)
        
   ' se executar a primera vez
   If flag = False Then
      ' pegando o Hwnd do quadro de mensagem
      hMessageBox = FindWindow("#32770", App.Title)
   End If
        
   Select Case idEvent
   ' Timer1 : encerra
   Case 1
      If hMessageBox Then
         Call SendMessage(hMessageBox, WM_SYSCOMMAND, SC_CLOSE, ByVal 0&)
      End If
      ' finalizar  os timers
      KillTimer hwnd, 1
      KillTimer hwnd, 2
      flag = False
   'Timer 2 : Para o tempo
   Case 2
      ' Mostra el tiempo restante
      If hMessageBox Then
         Segundos = Segundos - 1000
          
          ' escreve no titleBar
         SetWindowText hMessageBox, "Esta Mensagem será finalizada em: " & _
                        CStr(Segundos / 1000) & " .. segundos "
         flag = True
      End If
   End Select
End Sub
Public Sub MsgBoxTabum(FORMULARIO As Form, _
                       Mensagem As String, _
                       Optional BotaoYesNo As Boolean)
   If Len(Mensagem) < 100 Then Mensagem = Mensagem & Space(100 - Len(Mensagem))
   ' Quantidade de segundos de duração da mensagem
    Segundos = 7000
    ' inicia do timer
    '''''''''''''''''''''''''''
    ' timer para mostrar o tempo restante no title bar do msgbox
    
    SetTimer FORMULARIO.hwnd, 2, 1000, AddressOf TimerProc
    
    ' timer para encerrar quando finalizar o tempo
    SetTimer FORMULARIO.hwnd, 1, Segundos, AddressOf TimerProc
    
    ' mostra o Msgbox
    MsgBox Mensagem, _
           IIf(Not BotaoYesNo, vbInformation, vbYesNo + vbDefaultButton2), _
           App.Title
    
    ' encerra o timer
    KillTimer FORMULARIO.hwnd, 1
    KillTimer FORMULARIO.hwnd, 2
    'hwnd
    ' seta o flag do começo
    flag = False
End Sub

