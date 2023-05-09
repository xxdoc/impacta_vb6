VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCadastroCliente 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Cliente"
   ClientHeight    =   10590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   20250
   Icon            =   "frmCadastroCliente.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10590
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
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
      Left            =   18150
      TabIndex        =   23
      Top             =   -30
      Width           =   2025
      Begin VB.CommandButton cmdConsultar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consultar"
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
         Picture         =   "frmCadastroCliente.frx":1486
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Excluir Registro"
         Top             =   180
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
         Picture         =   "frmCadastroCliente.frx":1F90
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Salvar Registro"
         Top             =   1080
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
         Left            =   1005
         Picture         =   "frmCadastroCliente.frx":2A9A
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Excluir Registro"
         Top             =   1080
         Width           =   950
      End
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
         Picture         =   "frmCadastroCliente.frx":35A4
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Limpar Campos"
         Top             =   180
         Width           =   950
      End
   End
   Begin VB.Frame fraContato 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contato"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   30
      TabIndex        =   18
      Top             =   1635
      Width           =   14490
      Begin VB.TextBox txtContato 
         Height          =   360
         Left            =   180
         MaxLength       =   60
         TabIndex        =   6
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtObs 
         Height          =   360
         Left            =   5175
         MaxLength       =   100
         TabIndex        =   8
         Top             =   480
         Width           =   3915
      End
      Begin VB.TextBox txtEmail 
         Height          =   360
         Left            =   180
         MaxLength       =   60
         TabIndex        =   9
         Top             =   1155
         Width           =   3915
      End
      Begin MSMask.MaskEdBox mskTelefone 
         Height          =   360
         Left            =   2580
         TabIndex        =   7
         Top             =   480
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "## # ####-####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nome"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   22
         Top             =   255
         Width           =   420
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Whatsapp"
         Height          =   195
         Index           =   10
         Left            =   2580
         TabIndex        =   21
         Top             =   255
         Width           =   735
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Observações"
         Height          =   195
         Index           =   9
         Left            =   5160
         TabIndex        =   20
         Top             =   255
         Width           =   945
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "E-mail"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   19
         Top             =   915
         Width           =   420
      End
   End
   Begin VB.Frame RF 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Geral"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   30
      TabIndex        =   11
      Top             =   -30
      Width           =   14490
      Begin VB.TextBox txtCidade 
         Height          =   360
         Left            =   7530
         MaxLength       =   60
         TabIndex        =   5
         Top             =   1185
         Width           =   3525
      End
      Begin VB.TextBox txtBairro 
         Height          =   360
         Left            =   4245
         MaxLength       =   40
         TabIndex        =   4
         Top             =   1185
         Width           =   3075
      End
      Begin VB.TextBox txtEndereco 
         Height          =   360
         Left            =   120
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1185
         Width           =   3915
      End
      Begin MSMask.MaskEdBox mskCPFouCNPJ 
         Height          =   360
         Left            =   2505
         TabIndex        =   1
         Top             =   510
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtNome 
         Height          =   360
         Left            =   5100
         MaxLength       =   60
         TabIndex        =   2
         Top             =   510
         Width           =   5955
      End
      Begin VB.ComboBox cmbTipo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   510
         Width           =   2175
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cidade"
         Height          =   195
         Index           =   5
         Left            =   7500
         TabIndex        =   17
         Top             =   945
         Width           =   495
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bairro"
         Height          =   195
         Index           =   4
         Left            =   4245
         TabIndex        =   16
         Top             =   945
         Width           =   405
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Endereço"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   945
         Width           =   690
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nome"
         Height          =   195
         Index           =   2
         Left            =   5100
         TabIndex        =   14
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "CPF"
         Height          =   195
         Left            =   2505
         TabIndex        =   13
         Top             =   270
         Width           =   300
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   315
      End
   End
   Begin MSComctlLib.ListView lvwConsulta 
      Height          =   7245
      Left            =   0
      TabIndex        =   10
      Top             =   3330
      Width           =   20235
      _ExtentX        =   35692
      _ExtentY        =   12779
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmCadastroCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Private RsConsulta               As New ADODB.Recordset
   Private Item                     As ListItem
   Private IL                       As Integer
   
   Private Const colTipo            As Integer = 1
   Private Const colCodigo          As Integer = 2
   Private Const colNome            As Integer = 3
   Private Const colEndereco        As Integer = 4
   Private Const colBairro          As Integer = 5
   Private Const colCidade          As Integer = 6
   Private Const colContato         As Integer = 7
   Private Const colTelefone        As Integer = 8
   Private Const colObs             As Integer = 9
   Private Const colEmail           As Integer = 10
   Private Const colStatus          As Integer = 11
      
Private Sub MontaLvw()
   lvwConsulta.ListItems.Clear
   lvwConsulta.ColumnHeaders.Clear
   lvwConsulta.ColumnHeaders.Add , , "", 0
   lvwConsulta.ColumnHeaders.Add , , "Tipo", 800
   lvwConsulta.ColumnHeaders.Add , , "Código", 1700
   lvwConsulta.ColumnHeaders.Add , , "Nome", 2000
   lvwConsulta.ColumnHeaders.Add , , "Endereço", 3000
   lvwConsulta.ColumnHeaders.Add , , "Bairro", 0
   lvwConsulta.ColumnHeaders.Add , , "Cidade", 2000
   lvwConsulta.ColumnHeaders.Add , , "Contato", 1700
   lvwConsulta.ColumnHeaders.Add , , "Telefone", 1500
   lvwConsulta.ColumnHeaders.Add , , "Obs", 1500
   lvwConsulta.ColumnHeaders.Add , , "E-mail", 5000
   lvwConsulta.ColumnHeaders.Add , , "Status", 800, lvwColumnCenter
   lvwConsulta.View = lvwReport
End Sub

Private Sub cmbTipo_Click()
   If cmbTipo = "CPF" Then
      mskCPFouCNPJ.Mask = "###.###.###-##"
      lblCodigo = "CPF"
   Else
      mskCPFouCNPJ.Mask = "##.###.###/####-##"
      lblCodigo = "CNPJ"
   End If
End Sub

Private Sub cmbTipo_GotFocus()
   cmbTipo.BackColor = QBColor(14)
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Sendkeys "{tab}"
   End If
End Sub

Private Sub cmbTipo_LostFocus()
   CONSULTAR_REGISTROS
   cmbTipo.BackColor = QBColor(15)
End Sub

Private Sub cmdCancelar_Click()

   cmbTipo.Enabled = True
   mskCPFouCNPJ.Enabled = True
   
   If cmbTipo.ListCount > 0 Then cmbTipo.ListIndex = 0
   If cmbTipo = "CPF" Then mskCPFouCNPJ.Mask = "###.###.###-##"
   If cmbTipo = "CNPJ" Then mskCPFouCNPJ.Mask = "##.###.###/####-##"
   mskCPFouCNPJ = "___.___.___-__"
      
   lblCodigo = "CPF"
   txtNome = ""
   txtEndereco = ""
   txtBairro = ""
   txtCidade = ""
   txtContato = ""
   mskTelefone = "__ _ ____-____"
   txtObs = ""
   txtEmail = ""
      
   CONSULTAR_REGISTROS
   cmbTipo.SetFocus
   
End Sub

Private Sub cmdConsultar_Click()
   CONSULTAR_REGISTROS
End Sub

Private Sub cmdExcluir_Click()
   
   If lvwConsulta.ListItems.Count = 0 Then Exit Sub
   
   If MsgBox("Deseja Realmente Excluir Dado Selecionado?", vbYesNo + vbQuestion + vbDefaultButton2, "Exclusão") = vbNo Then
      Exit Sub
   End If
   
   cmdExcluir.Enabled = False
    
   With lvwConsulta
      For IL = 1 To .ListItems.Count
         Select Case .ListItems(IL).SubItems(colStatus)
         Case "D"
            CmdSql = "DELETE FROM CLIENTE" & vbCr
            CmdSql = CmdSql & "WHERE PROJETO    = " & PoeAspas(UCase(App.EXEName)) & vbCr
            CmdSql = CmdSql & "  AND CODIGO     = " & PoeAspas(.ListItems(IL).SubItems(colCodigo))
            CMySql.Executa CmdSql, True
         End Select
      Next IL
   End With
   
Cancelar:
   cmdExcluir.Enabled = True
   
   Form_Load
   cmbTipo.SetFocus

End Sub

Private Sub cmdSalvar_Click()
   
   If mskCPFouCNPJ.ClipText = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O CÓDIGO"
      mskCPFouCNPJ.SetFocus
      Exit Sub
   End If

   If Trim(txtNome) = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O NOME"
      txtNome.SetFocus
      Exit Sub
   End If
   
   If Trim(txtNome) = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O NOME"
      txtNome.SetFocus
      Exit Sub
   End If
            
   If Trim(txtEndereco) = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O ENDEREÇO"
      txtEndereco.SetFocus
      Exit Sub
   End If
            
   If Trim(txtBairro) = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O BAIRRO"
      txtBairro.SetFocus
      Exit Sub
   End If
            
   If Trim(txtCidade) = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER A CIDADE"
      txtCidade.SetFocus
      Exit Sub
   End If
            
   cmdSalvar.Enabled = False
            
   CmdSql = "SELECT * FROM CLIENTE" & vbCr
   CmdSql = CmdSql & "WHERE PROJETO    = " & PoeAspas(UCase(App.EXEName)) & vbCr
   CmdSql = CmdSql & "  AND CODIGO     = " & PoeAspas(mskCPFouCNPJ.ClipText) & vbCr
   CmdSql = CmdSql & "ORDER BY PROJETO,CODIGO"
   CMySql.Consulta CmdSql, RsConsulta
   
   If RsConsulta.EOF Then
      CmdSql = "INSERT INTO CLIENTE(PROJETO,TIPO,CODIGO,NOME,ENDERECO,BAIRRO,CIDADE,CONTATO,TELEFONE,OBS,EMAIL)" & vbCr
      CmdSql = CmdSql & "VALUES("
      CmdSql = CmdSql & PoeAspas(UCase(App.EXEName)) & ","
      CmdSql = CmdSql & PoeAspas(UCase(cmbTipo)) & ","
      CmdSql = CmdSql & PoeAspas(UCase(mskCPFouCNPJ.ClipText)) & ","
      CmdSql = CmdSql & PoeAspas(txtNome) & ","
      CmdSql = CmdSql & PoeAspas(txtEndereco) & ","
      CmdSql = CmdSql & PoeAspas(txtBairro) & ","
      CmdSql = CmdSql & PoeAspas(txtCidade) & ","
      CmdSql = CmdSql & PoeAspas(txtContato) & ","
      CmdSql = CmdSql & PoeAspas(mskTelefone.ClipText) & ","
      CmdSql = CmdSql & PoeAspas(txtObs) & ","
      CmdSql = CmdSql & PoeAspas(txtEmail) & ")"
      CMySql.Executa CmdSql, True
   
      MsgBoxTabum Me, "REGISTRO INCLUÍDO COM SUCESSO"
   Else
      CmdSql = "UPDATE CLIENTE SET TIPO     = " & PoeAspas(UCase(cmbTipo)) & "," & vbCr
      CmdSql = CmdSql & "                   CODIGO   = " & PoeAspas(UCase(mskCPFouCNPJ.ClipText)) & "," & vbCr
      CmdSql = CmdSql & "                   NOME     = " & PoeAspas(txtNome) & "," & vbCr
      CmdSql = CmdSql & "                   ENDERECO = " & PoeAspas(txtEndereco) & "," & vbCr
      CmdSql = CmdSql & "                   BAIRRO   = " & PoeAspas(txtBairro) & "," & vbCr
      CmdSql = CmdSql & "                   CIDADE   = " & PoeAspas(txtCidade) & "," & vbCr
      CmdSql = CmdSql & "                   CONTATO  = " & PoeAspas(txtContato) & "," & vbCr
      CmdSql = CmdSql & "                   TELEFONE = " & PoeAspas(mskTelefone.ClipText) & "," & vbCr
      CmdSql = CmdSql & "                   OBS      = " & PoeAspas(txtObs) & "," & vbCr
      CmdSql = CmdSql & "                   EMAIL    = " & PoeAspas(txtEmail) & vbCr
      CmdSql = CmdSql & "WHERE PROJETO = " & PoeAspas(UCase(App.EXEName)) & vbCr
      CmdSql = CmdSql & "  AND CODIGO  = " & PoeAspas(UCase(mskCPFouCNPJ.ClipText))
      CMySql.Executa CmdSql, True
      
      MsgBoxTabum Me, "REGISTRO ALTERADO COM SUCESSO"
   End If
         
   cmdCancelar_Click
   cmdSalvar.Enabled = True
   CONSULTAR_REGISTROS
   cmbTipo.SetFocus
   
End Sub

Private Sub Form_Load()
   MontaLvw
   
   cmbTipo.Clear
   cmbTipo.AddItem "CPF"
   cmbTipo.AddItem "CNPJ"
   cmbTipo.ListIndex = 0
   
   CONSULTAR_REGISTROS
End Sub

Private Sub lvwConsulta_DblClick()

   If lvwConsulta.ListItems.Count = 0 Then Exit Sub

   If UCase(lvwConsulta.SelectedItem.SubItems(colTipo)) = "CPF" Then
      mskCPFouCNPJ.Mask = "###.###.###-##"
      mskCPFouCNPJ = "___.___.___-__"
      lblCodigo = "CPF"
   End If
   
   If UCase(lvwConsulta.SelectedItem.SubItems(colTipo)) = "CNPJ" Then
      mskCPFouCNPJ.Mask = "##.###.###/####-##"
      mskCPFouCNPJ = "__.___.___/____-__"
      lblCodigo = "CNPJ"
   End If
      
   txtNome = ""
   txtEndereco = ""
   txtBairro = ""
   txtCidade = ""
   txtContato = ""
   mskTelefone = "__ _ ____-____"
   txtObs = ""
   txtEmail = ""
      
   cmbTipo = UCase(lvwConsulta.SelectedItem.SubItems(colTipo))
      
   If UCase(lvwConsulta.SelectedItem.SubItems(colTipo)) = "CPF" Then
      mskCPFouCNPJ.Mask = "###.###.###-##"
      lblCodigo = "CPF"
      
      mskCPFouCNPJ = Format(lvwConsulta.SelectedItem.SubItems(colCodigo), "&&&.&&&.&&&-&&")
   Else
      mskCPFouCNPJ.Mask = "##.###.###/####-##"
      lblCodigo = "CNPJ"
      
      mskCPFouCNPJ = Format(lvwConsulta.SelectedItem.SubItems(colCodigo), "&&.&&&.&&&/&&&&-&&")
   End If
   
   cmbTipo.Enabled = False
   mskCPFouCNPJ.Enabled = False
         
   txtNome = lvwConsulta.SelectedItem.SubItems(colNome)
   txtEndereco = lvwConsulta.SelectedItem.SubItems(colEndereco)
   txtBairro = lvwConsulta.SelectedItem.SubItems(colBairro)
   txtCidade = lvwConsulta.SelectedItem.SubItems(colCidade)
   txtContato = lvwConsulta.SelectedItem.SubItems(colContato)
   
   If Len(lvwConsulta.SelectedItem.SubItems(colTelefone)) = 11 Then
      mskTelefone = Format(lvwConsulta.SelectedItem.SubItems(colTelefone), "## # ####-####")
   End If
   
   txtObs = lvwConsulta.SelectedItem.SubItems(colObs)
   txtEmail = lvwConsulta.SelectedItem.SubItems(colEmail)
   
   txtNome.SetFocus
End Sub

Private Sub lvwConsulta_KeyDown(KeyCode As Integer, Shift As Integer)

   If lvwConsulta.ListItems.Count = 0 Then Exit Sub
   
   Select Case KeyCode
   Case "46"
      Select Case lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(colStatus)
      Case ""
         lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(colStatus) = "D"
      Case "D"
         lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(colStatus) = ""
      End Select
   End Select
End Sub

Private Sub CONSULTAR_REGISTROS()

   MontaLvw
   
   cmbTipo.Enabled = True
   mskCPFouCNPJ.Enabled = True
   
   CmdSql = "SELECT *" & vbCr
   CmdSql = CmdSql & "FROM CLIENTE" & vbCr
   CmdSql = CmdSql & "WHERE PROJETO = " & PoeAspas(UCase(App.EXEName)) & vbCr
   If mskCPFouCNPJ.ClipText <> "" Then CmdSql = CmdSql & "  AND CODIGO = " & PoeAspas(mskCPFouCNPJ.ClipText) & vbCr
   CmdSql = CmdSql & "ORDER BY PROJETO,TIPO,CODIGO"
   CMySql.Consulta CmdSql, RsConsulta
      
   Do While Not RsConsulta.EOF
      Set Item = lvwConsulta.ListItems.Add(, , "")
      Item.SubItems(colTipo) = Trim(RsConsulta("TIPO"))
      Item.SubItems(colCodigo) = Trim(RsConsulta("CODIGO"))
      Item.SubItems(colNome) = Trim(RsConsulta("NOME"))
      Item.SubItems(colEndereco) = Trim(RsConsulta("ENDERECO"))
      Item.SubItems(colBairro) = Trim(RsConsulta("BAIRRO"))
      Item.SubItems(colCidade) = Trim(RsConsulta("CIDADE"))
      Item.SubItems(colContato) = Trim(RsConsulta("CONTATO"))
      Item.SubItems(colTelefone) = Trim(RsConsulta("TELEFONE"))
      Item.SubItems(colObs) = Trim(RsConsulta("OBS"))
      Item.SubItems(colEmail) = Trim(RsConsulta("EMAIL"))
      Item.SubItems(colStatus) = ""
   RsConsulta.MoveNext
   Loop

End Sub

Private Sub mskCPFouCNPJ_GotFocus()
   Marca Me
   mskCPFouCNPJ.BackColor = QBColor(14)
End Sub

Private Sub mskCPFouCNPJ_LostFocus()
   mskCPFouCNPJ.BackColor = QBColor(15)
End Sub

Private Sub mskCPFouCNPJ_Validate(Cancel As Boolean)

   If mskCPFouCNPJ.ClipText <> "" Then
      If cmbTipo = "CPF" Then
         If IsValidCPF(mskCPFouCNPJ.ClipText) = False Then
            MsgBoxTabum Me, "CPF INVÁLIDO"
            Marca Me
            mskCPFouCNPJ.SetFocus
            Cancel = True
         End If
      Else
         If Len(mskCPFouCNPJ.ClipText) <> 14 Then
            MsgBoxTabum Me, "CNPJ INVÁLIDO"
            Marca Me
            mskCPFouCNPJ.SetFocus
            Cancel = True
         End If
      End If
   End If
   
End Sub

Private Sub mskTelefone_GotFocus()
   Marca Me
   mskTelefone.BackColor = QBColor(14)
End Sub

Private Sub mskTelefone_LostFocus()
   mskTelefone.BackColor = QBColor(15)
End Sub

Private Sub mskTelefone_Validate(Cancel As Boolean)
   If mskTelefone.ClipText <> "" Then
      If Len(mskTelefone.ClipText) <> 11 Then
         MsgBoxTabum Me, "Favor Preencher o campo de Telefone Corretamente"
         Marca Me
         Cancel = True
         mskTelefone.SetFocus
      End If
   End If
End Sub

Private Sub txtBairro_GotFocus()
   Marca Me
   txtBairro.BackColor = QBColor(14)
End Sub

Private Sub txtBairro_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub txtBairro_LostFocus()
   txtBairro.BackColor = QBColor(15)
End Sub

Private Sub txtCidade_GotFocus()
   Marca Me
   txtCidade.BackColor = QBColor(14)
End Sub

Private Sub txtCidade_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub txtCidade_LostFocus()
   txtCidade.BackColor = QBColor(15)
End Sub

Private Sub txtContato_GotFocus()
   Marca Me
   txtContato.BackColor = QBColor(14)
End Sub

Private Sub txtContato_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub txtContato_LostFocus()
   txtContato.BackColor = QBColor(15)
End Sub

Private Sub txtEmail_GotFocus()
   Marca Me
   txtEmail.BackColor = QBColor(14)
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub txtEmail_LostFocus()
   txtEmail.BackColor = QBColor(15)
End Sub

Private Sub txtEmail_Validate(Cancel As Boolean)
   If Trim(txtEmail) <> "" Then
      If IsValidEmail(txtEmail) = False Then
         MsgBoxTabum Me, "Favor Preencher E-mail corretamente"
         Marca Me
         Cancel = True
         txtEmail.SetFocus
      End If
   End If
End Sub

Private Sub txtEndereco_GotFocus()
   Marca Me
   txtEndereco.BackColor = QBColor(14)
End Sub

Private Sub txtEndereco_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub txtEndereco_LostFocus()
   txtEndereco.BackColor = QBColor(15)
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

Private Sub txtObs_GotFocus()
   Marca Me
   txtObs.BackColor = QBColor(14)
End Sub

Private Sub txtObs_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub txtObs_LostFocus()
   txtObs.BackColor = QBColor(15)
End Sub
