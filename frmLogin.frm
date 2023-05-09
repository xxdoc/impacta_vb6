VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login - Tabum"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7020
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSenha 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Logon do Sistema"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2003
      TabIndex        =   8
      Top             =   1995
      Width           =   3015
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1710
         TabIndex        =   5
         Top             =   1545
         Width           =   675
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Sair"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   975
         TabIndex        =   4
         Top             =   1545
         Width           =   675
      End
      Begin VB.TextBox TxtSenha 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   975
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   645
         Width           =   1785
      End
      Begin VB.TextBox txtUsuario 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   975
         MaxLength       =   20
         TabIndex        =   0
         Top             =   345
         Width           =   1785
      End
      Begin VB.TextBox txtRepetirSenha 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   975
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   945
         Width           =   1785
      End
      Begin VB.CheckBox chkMostrarSenha 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mostrar Senha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   975
         TabIndex        =   3
         Top             =   1275
         Width           =   1470
      End
      Begin VB.Label lblSenha 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Senha:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   10
         Top             =   675
         Width           =   615
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuário:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   345
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   23
      TabIndex        =   6
      Top             =   4035
      Width           =   6975
      Begin VB.Label lblCapsLook 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Caps Lock Ligado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   105
         TabIndex        =   7
         Top             =   210
         Width           =   6735
      End
   End
   Begin VB.Label lblProductName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Uma explosão de soluções para sua empresa!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   75
      TabIndex        =   11
      Top             =   795
      Width           =   6870
   End
   Begin VB.Image Image 
      Height          =   900
      Left            =   2130
      Picture         =   "frmLogin.frx":1486
      Top             =   -30
      Width           =   2700
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

   Private Type KeyboardBytes
      kbByte(0 To 255) As Byte
   End Type
   
   Private kbArray As KeyboardBytes
   Private Declare Function GetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
   Private Declare Function SetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long

   Private CONT As Integer
   
Private Sub chkMostrarSenha_Click()
   If chkMostrarSenha.Value = 1 Then
      txtSenha.PasswordChar = ""
      txtRepetirSenha.PasswordChar = ""
   Else
      txtSenha.PasswordChar = "*"
      txtRepetirSenha.PasswordChar = "*"
   End If
End Sub

Private Sub cmdOK_Click()

   If Trim(txtUsuario) = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O USUÁRIO"
      txtUsuario.SetFocus
      Exit Sub
   End If
   
   If Trim(txtSenha) = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER A SENHA"
      txtSenha.SetFocus
      Exit Sub
   End If
   
   If txtRepetirSenha.Visible = True Then
      If Trim(txtRepetirSenha) = "" Then
         MsgBoxTabum Me, "FAVOR REPETIR A SENHA"
         txtRepetirSenha.SetFocus
         Exit Sub
      End If
   End If
   
   If txtRepetirSenha.Visible = True Then
      If Len(txtUsuario) < 4 Then
         MsgBoxTabum Me, "USUÁRIO DEVE TER PELO MENOS 4 CARACTERES"
         txtUsuario.SetFocus
         Exit Sub
      End If
      
      If Trim(txtSenha) <> Trim(txtRepetirSenha) Then
         MsgBoxTabum Me, "AS SENHAS FORNECIDAS, NÃO CORRESPONDEM"
         txtSenha.SetFocus
         Exit Sub
      End If
      
      If Len(Trim(txtSenha)) < 6 Then
         MsgBoxTabum Me, "SENHA DEVE TER PELO MENOS 6 CARACTERES"
         txtSenha.SetFocus
         Exit Sub
      End If
            
      CmdSql = "SELECT * FROM ACESSO" & vbCr
      CmdSql = CmdSql & "WHERE USUARIO = " & PoeAspas(txtUsuario)
      CMySql.Consulta CmdSql, Rs
      
      If Not Rs.EOF Then
         txtUsuario = Trim(Rs("USUARIO"))
         txtSenha = ""
         txtRepetirSenha = ""
         txtRepetirSenha.Visible = False
         chkMostrarSenha.Top = 975
         cmdSair.Top = 1250
         cmdOK.Top = 1250
         fraSenha.Height = 1720
      Else
         CmdSql = "SELECT COUNT(*) QTDE" & vbCr
         CmdSql = CmdSql & "FROM ACESSO"
         CMySql.Consulta CmdSql, RsConsulta
         
         If RsConsulta("QTDE") = 0 Then
            CmdSql = "INSERT INTO ACESSO(PERFIL,USUARIO,SENHA,STATUS)" & vbCr
            CmdSql = CmdSql & "VALUES("
            CmdSql = CmdSql & PoeAspas("ADMINISTRADOR") & ","
            CmdSql = CmdSql & PoeAspas(txtUsuario) & ","
            CmdSql = CmdSql & PoeAspas(Crypt(txtSenha)) & ","
            CmdSql = CmdSql & PoeAspas("ATIVO") & ")"
            CMySql.Executa CmdSql, True
            
            CmdSql = "INSERT INTO TIPOCOMBO(PROJETO,FORMULARIO,TIPO,DESCRICAO)" & vbCr
            CmdSql = CmdSql & "VALUES("
            CmdSql = CmdSql & PoeAspas(UCase(App.EXEName)) & ","
            CmdSql = CmdSql & PoeAspas("CADASTRO DE USUÁRIO") & ","
            CmdSql = CmdSql & PoeAspas("PERFIL") & ","
            CmdSql = CmdSql & PoeAspas("ADMINISTRADOR") & ")"
            CMySql.Executa CmdSql, True
            
            CmdSql = "INSERT INTO TIPOCOMBO(PROJETO,FORMULARIO,TIPO,DESCRICAO)" & vbCr
            CmdSql = CmdSql & "VALUES("
            CmdSql = CmdSql & PoeAspas(UCase(App.EXEName)) & ","
            CmdSql = CmdSql & PoeAspas("CADASTRO DE USUÁRIO") & ","
            CmdSql = CmdSql & PoeAspas("STATUS") & ","
            CmdSql = CmdSql & PoeAspas("ATIVO") & ")"
            CMySql.Executa CmdSql, True
            
            GoTo ACESSO_SISTEMA
         Else
            MsgBoxTabum Me, "USUÁRIO NÃO ENCONTRADO"
            Exit Sub
         End If
         
         txtUsuario = ""
         txtSenha = ""
         txtRepetirSenha = ""
         txtRepetirSenha.Visible = True
         chkMostrarSenha.Visible = True
         chkMostrarSenha.Top = 1245
         cmdSair.Top = 1515
         cmdOK.Top = 1515
         fraSenha.Height = 1980
      End If
           
      GoTo ACESSO_SISTEMA
   Else
      
      CmdSql = "SELECT * FROM ACESSO" & vbCr
      CmdSql = CmdSql & "WHERE USUARIO = " & PoeAspas(txtUsuario)
      CMySql.Consulta CmdSql, Rs
      
      If Not Rs.EOF Then
         If Trim(txtUsuario) = Trim(Rs("USUARIO")) And Trim(txtSenha) = Crypt(Trim(Rs("SENHA"))) Then

ACESSO_SISTEMA:

            VARIAVEIS_SISTEMA (txtUsuario)
            
            Set FormMdi = MDIPrincipal
            MDIPrincipal.Show
            Unload frmLogin
         Else
            CONT = CONT + 1
      
            If CONT = 3 Then
               MsgBoxTabum Me, CONT & "ª TENTATIVA, FINALIZANDO SISTEMA"
               End
            End If
      
            MsgBoxTabum Me, "USUÁRIO OU SENHA INCORRETO, TENTE NOVAMENTE" & vbCr & CONT & "ª TENTATIVA de 3"
            txtSenha.SetFocus
            Marca Me
         End If
      Else
         CONT = CONT + 1
      
         If CONT = 3 Then
            MsgBoxTabum Me, CONT & "ª TENTATIVA, FINALIZANDO SISTEMA"
            End
         End If
            
         MsgBoxTabum Me, "USUÁRIO OU SENHA INCORRETO, TENTE NOVAMENTE" & vbCr & CONT & "ª TENTATIVA de 3"
         txtSenha.SetFocus
         Marca Me
      End If
      
   End If
   
End Sub

Private Sub cmdSair_Click()
   End
End Sub

Private Sub Form_Initialize()
   lblCapsLook = ""

   CmdSql = "SHOW TABLES LIKE 'ACESSO'"
   CMySql.Consulta CmdSql, Rs
   
   If Rs.EOF Then
      CmdSql = "CREATE TABLE ACESSO(" & vbCr
      CmdSql = CmdSql & "        PERFIL          VARCHAR(40)     NOT NULL," & vbCr
      CmdSql = CmdSql & "        USUARIO         VARCHAR(20)     NOT NULL," & vbCr
      CmdSql = CmdSql & "        NOME            VARCHAR(60)     NOT NULL," & vbCr
      CmdSql = CmdSql & "        SENHA           VARCHAR(100)    NOT NULL," & vbCr
      CmdSql = CmdSql & "        STATUS          VARCHAR(40)     NOT NULL," & vbCr
      CmdSql = CmdSql & "        PRIMARY KEY(USUARIO)" & vbCr
      CmdSql = CmdSql & ")TYPE=MYISAM"
      CMySql.Executa CmdSql, True
      
      txtUsuario = ""
      txtSenha = ""
      txtRepetirSenha = ""
      txtRepetirSenha.Visible = True
      chkMostrarSenha.Visible = True
      chkMostrarSenha.Top = 1275
   Else
      CmdSql = "SELECT * FROM ACESSO" & vbCr
      CMySql.Consulta CmdSql, Rs
      
      If Rs.EOF Then
         txtUsuario = ""
         txtSenha = ""
         txtRepetirSenha = ""
         txtRepetirSenha.Visible = True
         chkMostrarSenha.Visible = True
         chkMostrarSenha.Top = 1275
      Else
         txtRepetirSenha = ""
         txtRepetirSenha.Visible = False
         chkMostrarSenha.Top = 980
      End If
   End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Sendkeys "{tab}"
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
   Call GetKeyboardState(kbArray)
   If kbArray.kbByte(&H14) = 1 Then
      lblCapsLook.Visible = True
      lblCapsLook = "Caps Lock Ligado"
   Else
      lblCapsLook.Visible = False
      lblCapsLook = ""
   End If
   Call SetKeyboardState(kbArray)
End Sub

Private Sub txtRepetirSenha_GotFocus()
   Marca Me
   txtRepetirSenha.BackColor = QBColor(14)
End Sub

Private Sub txtRepetirSenha_LostFocus()
   txtRepetirSenha.BackColor = QBColor(15)
End Sub

Private Sub txtSenha_GotFocus()
   Marca Me
   txtSenha.BackColor = QBColor(14)
End Sub

Private Sub txtSenha_LostFocus()
   If txtRepetirSenha.Visible = False Then cmdOK.SetFocus
   txtSenha.BackColor = QBColor(15)
End Sub

Private Sub txtUsuario_GotFocus()
   Marca Me
   txtUsuario.BackColor = QBColor(14)
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtUsuario_LostFocus()
   txtUsuario.BackColor = QBColor(15)
End Sub

