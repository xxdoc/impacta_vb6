Attribute VB_Name = "MDL_Seguranca"
Option Explicit
   Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
   
   Public szNomeARrquivo   As String
   Public sBanco           As String
   Public sRet             As Long
   Public FINALIZARSIST    As Double
   
   Public VTEMPO           As Integer
   Public NomeFormulario   As String
   Public VSegTela         As String
   Public RetornoSeguranca As String
   Public NomeProjeto      As Form
   Public ADOTela          As New ADODB.Recordset
   
   Dim ADOSeguranca        As New ADODB.Recordset

Public Function TempoRestante() As Integer

   '-----
   ' verifica quantas telas faltam para serem finalizadas
   ' mostra a tela
   '-----
   FINALIZARSIST = FINALIZARSIST - 1
   DownSist.LblDown.Caption = " O Sistema entrar� em manuten��o " & _
                              "dentro de instantes..." & Chr(13)

   If FINALIZARSIST = 1 Then
      DownSist.LblDown.Caption = DownSist.LblDown.Caption & " Voc� " & _
                                 "ter� mais " & FINALIZARSIST & " aviso."
   ElseIf FINALIZARSIST = 0 Then
      DownSist.LblDown.Caption = DownSist.LblDown.Caption & "O Sistema ser� Finalizado. "
   Else
      DownSist.LblDown.Caption = DownSist.LblDown.Caption & " Voc� ter� " & _
                                 "mais " & FINALIZARSIST & " avisos." & _
               "Desculpe-nos pelos transtornos." & Chr(13) & _
               "Caso haja d�vida, entrar em contato com o CPD." & Chr(13) & _
               "Dentro de instantes o Sistema estar� dipon�vel..."
   End If
   
   TempoRestante = FINALIZARSIST
   
   DownSist.Show vbModal
   DoEvents
End Function
