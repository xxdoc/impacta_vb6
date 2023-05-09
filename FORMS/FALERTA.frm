VERSION 5.00
Begin VB.Form FALERTA 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alerta"
   ClientHeight    =   3450
   ClientLeft      =   1290
   ClientTop       =   2040
   ClientWidth     =   8955
   ControlBox      =   0   'False
   Icon            =   "FALERTA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   105
      Top             =   210
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Encerrar o Sistema"
      Height          =   420
      Left            =   7200
      TabIndex        =   3
      Top             =   2940
      Width           =   1695
   End
   Begin VB.CommandButton CmdFecha 
      Caption         =   "&Fechar esta janela"
      Height          =   420
      Left            =   5490
      TabIndex        =   0
      Top             =   2940
      Width           =   1695
   End
   Begin VB.Label LblTMP 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "          "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   330
      Left            =   2565
      TabIndex        =   5
      Top             =   3000
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Tempo Restante:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   75
      TabIndex        =   4
      Top             =   3000
      Width           =   2370
   End
   Begin VB.Label LblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "\zxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1590
      Left            =   105
      TabIndex        =   2
      Top             =   840
      Width           =   8805
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "A T E N Ç Ã O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   555
      Left            =   2925
      TabIndex        =   1
      Top             =   150
      Width           =   2880
   End
End
Attribute VB_Name = "FALERTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
   Private Topo            As Long
   Private ModoSemEscolha  As Boolean

Private Sub CmdFecha_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   End
End Sub

Private Sub Form_Activate()
   Me.WindowState = 0
   Topo = SetWindowPos(Me.hwnd, -1, 1, 1, 1, 1, 3)
   '' para voltar ao normal assim...
   If Not ModoSemEscolha Then Topo = SetWindowPos(Me.hwnd, -2, 1, 1, 1, 1, 3)
   'Me.Show
End Sub

Private Sub Timer1_Timer()
'   LblTMP = MDIPrincipal.NkFecha.TempoRestante
'   LblTMP.Refresh
End Sub

Public Sub SemEscolha()
   Dim Ini As Long
   Dim Dur As Long
   
   ModoSemEscolha = True
   Ini = Timer
   Dur = 20
   
   Do While Timer < Ini + Dur
      LblTMP = Abs(Val(Timer - (Ini + Dur)))
      If LblTMP = 0 Then End
      'LblTMP.Refresh
      DoEvents
   Loop
   End
End Sub

