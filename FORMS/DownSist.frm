VERSION 5.00
Begin VB.Form DownSist 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6870
   ClientLeft      =   4335
   ClientTop       =   1620
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Begin VB.Timer RolaForm 
      Interval        =   200
      Left            =   90
      Top             =   45
   End
   Begin VB.Label LBLTEMPO 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4380
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3930
      Picture         =   "DownSist.frx":0000
      Top             =   2730
      Width           =   480
   End
   Begin VB.Label LblDown 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LblDown"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   900
      TabIndex        =   0
      Top             =   3390
      Width           =   5445
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   4050
      Picture         =   "DownSist.frx":030A
      Top             =   990
      Width           =   480
   End
   Begin VB.Image IMGFigura 
      Height          =   2550
      Left            =   1005
      Picture         =   "DownSist.frx":0614
      Top             =   720
      Width           =   5295
   End
End
Attribute VB_Name = "DownSist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Sub Form_Load()
   Dim hr&, dl&
   Dim usew&, useh&
   Me.Top = (Screen.Height / 2) - (Me.Height / 2)
   Me.Left = Screen.Width
   
   usew& = Me.Width / Screen.TwipsPerPixelX
   useh& = Me.Height / Screen.TwipsPerPixelY
   hr& = CreateEllipticRgn(0, 0, usew, useh)
   dl& = SetWindowRgn(Me.hwnd, hr, True)
   
   'LBLTEMPO = FINALIZARSIST
   Me.Refresh
   
End Sub

Private Sub RolaForm_Timer()
   
   Me.Left = Me.Left - 200
   If Me.Left <= -Me.Width Then Unload Me
   
End Sub
