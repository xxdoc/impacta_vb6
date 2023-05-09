VERSION 5.00
Begin VB.Form FRMMSG 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   2175
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TFecha 
      Interval        =   5000
      Left            =   5310
      Top             =   90
   End
   Begin VB.Timer TAbre 
      Interval        =   10
      Left            =   2850
      Top             =   150
   End
   Begin VB.PictureBox PCTFigura 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   60
      ScaleHeight     =   2055
      ScaleWidth      =   2625
      TabIndex        =   0
      Top             =   15
      Width           =   2625
   End
   Begin VB.Line Line4 
      X1              =   5790
      X2              =   5790
      Y1              =   2085
      Y2              =   -60
   End
   Begin VB.Line Line3 
      X1              =   45
      X2              =   45
      Y1              =   0
      Y2              =   2160
   End
   Begin VB.Line Line2 
      X1              =   5730
      X2              =   30
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   5775
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Label LBLMSG 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "O seu usuário não tem permissão para utilizar esta área do Sistema. Consulte o CPD."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   2835
      TabIndex        =   2
      Top             =   810
      Width           =   2910
   End
   Begin VB.Label LBLTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "ATENÇÃO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   2985
      TabIndex        =   1
      Top             =   255
      Width           =   2715
   End
End
Attribute VB_Name = "FRMMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim Vh As Double

Private Sub Form_KeyPress(KeyAscii As Integer)
   Unload Me
End Sub

Private Sub Form_Load()
   Dim ValorX As Long
   Dim ValorY As Long

   ValorX = Screen.Width
   ValorY = Screen.Height

   Vh = 441
   Me.Height = Vh
   Me.Top = (ValorY - 500) - Vh '- 8000
   Me.Left = ValorX - Me.Width   '- 6200
   DoEvents

End Sub

Private Sub TAbre_Timer()
   If Vh <= 2205 Then
      Vh = Vh + 55
      Me.Height = Vh
      If Me.Top <= -Me.Width Then Unload Me
      Me.Top = Me.Top - 50
      'Me.Left = ValorX  '6200
      DoEvents
   End If

End Sub

Private Sub TFecha_Timer()
   Unload Me
   End Sub
Public Sub ObterResolucaoTela()

  Dim xTwips%, yTwips%, xPixels#, YPixels#
  xTwips = Screen.TwipsPerPixelX
  yTwips = Screen.TwipsPerPixelY
  YPixels = Screen.Height / yTwips
  xPixels = Screen.Width / xTwips
  MsgBox "A Resolução é: " & Str$(xPixels) + _
              " por " + Str$(YPixels)

End Sub

