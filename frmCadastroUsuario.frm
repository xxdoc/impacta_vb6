VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCadastroUsuario 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Usuário"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12885
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCadastroUsuario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   12885
   Begin VB.Frame fraBotoes 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC7404&
      Height          =   2010
      Left            =   10830
      TabIndex        =   10
      Top             =   -15
      Width           =   2025
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Limpar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   1005
         Picture         =   "frmCadastroUsuario.frx":1486
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpar Campos"
         Top             =   180
         Width           =   950
      End
      Begin VB.CommandButton cmdExcluir 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Excluir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   45
         Picture         =   "frmCadastroUsuario.frx":2008
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir Registro"
         Top             =   1080
         Width           =   950
      End
      Begin VB.CommandButton cmdSalvar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Salvar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   45
         Picture         =   "frmCadastroUsuario.frx":2B12
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salvar Registro"
         Top             =   180
         Width           =   950
      End
   End
   Begin VB.Frame fraDadosUsuario 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC7404&
      Height          =   2010
      Left            =   45
      TabIndex        =   9
      Top             =   -15
      Width           =   10725
      Begin VB.TextBox txtSenha 
         Height          =   360
         Left            =   7230
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1230
         Width           =   1500
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   360
         Left            =   3405
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1230
         Width           =   3210
      End
      Begin VB.ComboBox cmbPerfil 
         Height          =   360
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1230
         Width           =   2595
      End
      Begin VB.TextBox txtUsuario 
         Height          =   360
         Left            =   225
         MaxLength       =   20
         TabIndex        =   0
         Top             =   540
         Width           =   2595
      End
      Begin VB.TextBox txtNome 
         Height          =   360
         Left            =   3405
         MaxLength       =   60
         TabIndex        =   1
         Top             =   540
         Width           =   5325
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Senha"
         Height          =   240
         Index           =   4
         Left            =   7230
         TabIndex        =   15
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Perfil"
         Height          =   240
         Index           =   2
         Left            =   225
         TabIndex        =   14
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuário"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   13
         Top             =   285
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Status"
         Height          =   240
         Index           =   3
         Left            =   3405
         TabIndex        =   12
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nome"
         Height          =   240
         Index           =   1
         Left            =   3405
         TabIndex        =   11
         Top             =   285
         Width           =   540
      End
   End
   Begin MSComctlLib.ListView lvwConsulta 
      Height          =   5055
      Left            =   0
      TabIndex        =   8
      Top             =   1995
      Width           =   12885
      _ExtentX        =   22728
      _ExtentY        =   8916
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmCadastroUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

   Private RsConsulta               As New ADODB.Recordset
   Private Item                     As ListItem
   Private IL                       As Integer
   
   Private Const colPERFIL          As Integer = 1
   Private Const colUSUARIO         As Integer = 2
   Private Const colNome            As Integer = 3
   Private Const colStatus           As Integer = 4
   Private Const colSENHA           As Integer = 5

Private Sub MontaLvw()
   lvwConsulta.ListItems.Clear
   lvwConsulta.ColumnHeaders.Clear
   lvwConsulta.ColumnHeaders.Add , , "", 0
   lvwConsulta.ColumnHeaders.Add , , "PERFIL", 3000
   lvwConsulta.ColumnHeaders.Add , , "USUÁRIO", 2500
   lvwConsulta.ColumnHeaders.Add , , "NOME", 4800
   lvwConsulta.ColumnHeaders.Add , , "STATUS", 2200
   lvwConsulta.ColumnHeaders.Add , , "SENHA", 0
   lvwConsulta.View = lvwReport
End Sub

Private Sub cmdPerfil_GotFocus()
   cmbPerfil.BackColor = QBColor(14)
End Sub

Private Sub cmbPerfil_GotFocus()
   cmbPerfil.BackColor = QBColor(14)
End Sub

Private Sub cmbPerfil_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub cmbPerfil_LostFocus()
   cmbPerfil.BackColor = QBColor(15)
End Sub

Private Sub cmbStatus_GotFocus()
   cmbStatus.BackColor = QBColor(14)
End Sub

Private Sub cmbStatus_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub cmbStatus_LostFocus()
   cmbStatus.BackColor = QBColor(15)
End Sub

Private Sub cmdCancelar_Click()
   txtUsuario.Enabled = True
   cmbPerfil.Enabled = True
   cmbStatus.Enabled = True
   cmdExcluir.Enabled = True
   
   txtUsuario = ""
   txtNome = ""
   If cmbPerfil.ListCount > 0 Then cmbPerfil.ListIndex = 0
   If cmbStatus.ListCount > 0 Then cmbStatus.ListIndex = 0
   txtSenha = ""
   
   CONSULTA_USUARIO
End Sub

Private Sub cmdSalvar_Click()
   
   If Trim(txtUsuario) = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O USUÁRIO"
      txtUsuario.SetFocus
      Exit Sub
   End If
   
   If Trim(txtNome) = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O NOME DO USUÁRIO"
      txtNome.SetFocus
      Exit Sub
   End If
   
   If cmbPerfil = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O PERFIL DO USUÁRIO"
      cmbPerfil.SetFocus
      Exit Sub
   End If
   
   If cmbStatus = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER STATUS DO USUÁRIO"
      cmbStatus.SetFocus
      Exit Sub
   End If
   
   cmdSalvar.Enabled = False

   If txtUsuario.Enabled = True Then
      CmdSql = "SELECT * FROM ACESSO" & vbCr
      CmdSql = CmdSql & "WHERE USUARIO = " & PoeAspas(txtUsuario)
      CMySql.Consulta CmdSql, RsConsulta
      
      If Not RsConsulta.EOF Then
         MsgBoxTabum Me, "USUÁRIO EXISTENTE"
         txtUsuario.SetFocus
         cmdSalvar.Enabled = True
         Exit Sub
      Else
         CmdSql = "INSERT INTO ACESSO(PERFIL,USUARIO,NOME,SENHA,STATUS)" & vbCr
         CmdSql = CmdSql & "VALUES("
         CmdSql = CmdSql & PoeAspas(UCase(cmbPerfil)) & ","
         CmdSql = CmdSql & PoeAspas(UCase(txtUsuario)) & ","
         CmdSql = CmdSql & PoeAspas(UCase(txtNome)) & ","
         
         If Trim(txtSenha) = "" Then
            CmdSql = CmdSql & PoeAspas(Crypt("MUDAR@123")) & ","
         Else
            CmdSql = CmdSql & PoeAspas(Crypt(Trim(txtSenha))) & ","
         End If
         
         CmdSql = CmdSql & PoeAspas(UCase(cmbStatus)) & ")"
         CMySql.Executa CmdSql, True
      
         If Trim(txtSenha) <> "" Then
            MsgBoxTabum Me, "REGISTRO INCLUÍDO COM SUCESSO" & vbCr & "SENHA: " & txtSenha
         Else
            MsgBoxTabum Me, "REGISTRO INCLUÍDO COM SUCESSO" & vbCr & "SENHA PADRÃO: 'MUDAR@123'"
         End If
      End If
   Else
      CmdSql = "UPDATE ACESSO SET PERFIL = " & PoeAspas(UCase(cmbPerfil)) & "," & vbCr
      CmdSql = CmdSql & "                  NOME   = " & PoeAspas(UCase(txtNome)) & "," & vbCr
      
      If Trim(txtSenha) <> "" Then CmdSql = CmdSql & "                  SENHA  = " & PoeAspas(Crypt(txtSenha)) & "," & vbCr
            
      CmdSql = CmdSql & "                  STATUS = " & PoeAspas(UCase(cmbStatus)) & vbCr
      CmdSql = CmdSql & "WHERE USUARIO = " & PoeAspas(txtUsuario)
      CMySql.Executa CmdSql, True
   End If
   
   cmbPerfil.ListIndex = 0
   txtUsuario = ""
   txtNome = ""
   cmbStatus.ListIndex = 0
   txtSenha = ""
               
   cmdSalvar.Enabled = True
   txtUsuario.Enabled = True
   cmbPerfil.Enabled = True
   cmbStatus.Enabled = True
   cmdExcluir.Enabled = True
   
   CONSULTA_USUARIO
   txtUsuario.SetFocus

End Sub

Private Sub Form_Load()
   
   CARREGA_COMBOS App.EXEName, "CADASTRO DE USUÁRIO"
   
   cmbPerfil.Clear
   cmbStatus.Clear
   cmbPerfil.AddItem ""
   cmbStatus.AddItem ""
   
   Do While Not RsTipoCombo.EOF
      Select Case UCase(RsTipoCombo("TIPO"))
      Case "PERFIL"
         cmbPerfil.AddItem Trim(RsTipoCombo("DESCRICAO"))
      Case "STATUS"
         cmbStatus.AddItem Trim(RsTipoCombo("DESCRICAO"))
      End Select
   RsTipoCombo.MoveNext
   Loop
   
   If cmbPerfil.ListCount > 0 Then cmbPerfil.ListIndex = 0
   If cmbStatus.ListCount > 0 Then cmbStatus.ListIndex = 0
   
   CONSULTA_USUARIO
End Sub

Private Sub lvwConsulta_DblClick()
   
   txtSenha.Enabled = True
   txtUsuario = ""
   txtNome = ""
   If cmbPerfil.ListCount > 0 Then cmbPerfil.ListIndex = 0
   If cmbStatus.ListCount > 0 Then cmbStatus.ListIndex = 0
   txtSenha = ""
   
   cmbPerfil = lvwConsulta.SelectedItem.SubItems(colPERFIL)
   txtUsuario = lvwConsulta.SelectedItem.SubItems(colUSUARIO)
   txtNome = lvwConsulta.SelectedItem.SubItems(colNome)
   cmbStatus = lvwConsulta.SelectedItem.SubItems(colStatus)
   
   If lvwConsulta.SelectedItem.SubItems(colUSUARIO) = sysUsuario Then
      txtSenha = Crypt(Trim(lvwConsulta.SelectedItem.SubItems(colSENHA)))
   Else
      txtSenha.Enabled = False
   End If
      
   txtUsuario.Enabled = False
   
   If UCase(sysPerfil) <> "ADMINISTRADOR" Then
      cmbPerfil.Enabled = False
      cmbStatus.Enabled = False
      cmdExcluir.Enabled = False
   End If
End Sub

Private Sub txtNome_GotFocus()
   Marca Me
   txtNome.BackColor = QBColor(14)
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub txtNome_LostFocus()
   txtNome.BackColor = QBColor(15)
End Sub

Private Sub txtSenha_GotFocus()
   Marca Me
   txtSenha.BackColor = QBColor(14)
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Sendkeys "{tab}"
   End If
End Sub

Private Sub txtSenha_LostFocus()
   txtSenha.BackColor = QBColor(15)
End Sub

Private Sub txtUsuario_GotFocus()
   Marca Me
   txtUsuario.BackColor = QBColor(14)
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub txtUsuario_LostFocus()
   txtUsuario.BackColor = QBColor(15)
End Sub

Private Sub CONSULTA_USUARIO(Optional USUARIO As String)
   MontaLvw
   
   CmdSql = "SELECT *" & vbCr
   CmdSql = CmdSql & "FROM ACESSO" & vbCr
   CmdSql = CmdSql & "WHERE USUARIO <> ''" & vbCr
   If cmbPerfil <> "" Then CmdSql = CmdSql & "  AND PERFIL  = " & PoeAspas(cmbPerfil) & vbCr
   If Trim(txtUsuario) <> "" Then CmdSql = CmdSql & "  AND USUARIO LIKE ('%" & Trim(txtUsuario) & "%')" & vbCr
   If Trim(txtNome) <> "" Then CmdSql = CmdSql & "  AND NOME    LIKE ('%" & Trim(txtNome) & "%')" & vbCr
   If cmbStatus <> "" Then CmdSql = CmdSql & "  AND STATUS  = " & PoeAspas(cmbStatus) & vbCr
   
   If UCase(sysPerfil) <> "ADMINISTRADOR" Then
      CmdSql = CmdSql & "  AND USUARIO = " & PoeAspas(sysUsuario) & vbCr
   End If
   
   CmdSql = CmdSql & "ORDER BY PERFIL,USUARIO"
   CMySql.Consulta CmdSql, RsConsulta
      
   Do While Not RsConsulta.EOF
      Set Item = lvwConsulta.ListItems.Add(, , "")
      Item.SubItems(colPERFIL) = Trim(RsConsulta("PERFIL"))
      Item.SubItems(colUSUARIO) = Trim(RsConsulta("USUARIO"))
      Item.SubItems(colNome) = Trim(RsConsulta("NOME"))
      Item.SubItems(colStatus) = Trim(RsConsulta("STATUS"))
      Item.SubItems(colSENHA) = Trim(RsConsulta("SENHA"))
   RsConsulta.MoveNext
   Loop

End Sub
