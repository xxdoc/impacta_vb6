VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAtendimentoCliente 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atendimento Site"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13335
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAtendimentoCliente.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   13335
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
      Left            =   11295
      TabIndex        =   20
      Top             =   -90
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
         Picture         =   "frmAtendimentoCliente.frx":1486
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   1005
         Picture         =   "frmAtendimentoCliente.frx":2008
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmAtendimentoCliente.frx":2B12
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salvar Registro"
         Top             =   1080
         Width           =   950
      End
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
         Picture         =   "frmAtendimentoCliente.frx":361C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir Registro"
         Top             =   180
         Width           =   950
      End
   End
   Begin VB.Frame RF 
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
      Height          =   1995
      Left            =   30
      TabIndex        =   12
      Top             =   -90
      Width           =   11220
      Begin VB.ComboBox cmbStatus 
         Height          =   360
         Left            =   8820
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   405
         Width           =   2175
      End
      Begin VB.TextBox txtCliente 
         Height          =   360
         Left            =   4620
         MaxLength       =   60
         TabIndex        =   2
         Top             =   405
         Width           =   4035
      End
      Begin VB.TextBox txtPedido 
         Height          =   360
         Left            =   2295
         MaxLength       =   20
         TabIndex        =   1
         Top             =   405
         Width           =   2145
      End
      Begin VB.TextBox txtDescricaoProblema 
         Height          =   360
         Left            =   120
         MaxLength       =   255
         TabIndex        =   4
         Top             =   1035
         Width           =   6915
      End
      Begin VB.TextBox txtTicket 
         Height          =   360
         Left            =   120
         MaxLength       =   15
         TabIndex        =   0
         Top             =   405
         Width           =   1980
      End
      Begin MSMask.MaskEdBox mskAbertura 
         Height          =   360
         Left            =   7155
         TabIndex        =   5
         Top             =   1035
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskEncerramento 
         Height          =   360
         Left            =   8820
         TabIndex        =   6
         Top             =   1035
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Encerramento"
         Height          =   240
         Index           =   7
         Left            =   8805
         TabIndex        =   19
         Top             =   795
         Width           =   1365
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Abertura"
         Height          =   240
         Index           =   10
         Left            =   7170
         TabIndex        =   18
         Top             =   795
         Width           =   855
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Status do Pedido"
         Height          =   240
         Index           =   1
         Left            =   8820
         TabIndex        =   17
         Top             =   150
         Width           =   1695
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cliente"
         Height          =   240
         Index           =   5
         Left            =   4620
         TabIndex        =   16
         Top             =   150
         Width           =   675
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pedido"
         Height          =   240
         Index           =   4
         Left            =   2310
         TabIndex        =   15
         Top             =   150
         Width           =   645
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Descrição do Problema"
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   795
         Width           =   2235
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Número do Ticket"
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   150
         Width           =   1740
      End
   End
   Begin MSComctlLib.ListView lvwConsulta 
      Height          =   7245
      Left            =   15
      TabIndex        =   11
      Top             =   1920
      Width           =   13305
      _ExtentX        =   23469
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
Attribute VB_Name = "frmAtendimentoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

   Private RsConsulta               As New ADODB.Recordset
   Private Item                     As ListItem
   Private IL                       As Integer
   
   Private Const colTicket          As Integer = 1
   Private Const colPedido          As Integer = 2
   Private Const colCliente         As Integer = 3
   Private Const colStatusPed       As Integer = 4
   Private Const colDescricao       As Integer = 5
   Private Const colAbertura        As Integer = 6
   Private Const colEncerramento    As Integer = 7
   Private Const colStatus          As Integer = 8
      
Private Sub MontaLvw()
   lvwConsulta.ListItems.Clear
   lvwConsulta.ColumnHeaders.Clear
   lvwConsulta.ColumnHeaders.Add , , "", 0
   lvwConsulta.ColumnHeaders.Add , , "Ticket", 800
   lvwConsulta.ColumnHeaders.Add , , "Pedido", 1900
   lvwConsulta.ColumnHeaders.Add , , "Cliente", 2000
   lvwConsulta.ColumnHeaders.Add , , "Status Pedido", 2500
   lvwConsulta.ColumnHeaders.Add , , "Descrição", 0
   lvwConsulta.ColumnHeaders.Add , , "Abertura", 2000
   lvwConsulta.ColumnHeaders.Add , , "Encerramento", 1700
   lvwConsulta.ColumnHeaders.Add , , "Status", 800, lvwColumnCenter
   lvwConsulta.View = lvwReport
End Sub

Private Sub cmbStatus_GotFocus()
   cmbStatus.BackColor = QBColor(14)
End Sub

Private Sub cmbStatus_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub cmbStatus_LostFocus()
   cmbStatus.BackColor = QBColor(15)
End Sub

Private Sub cmdCancelar_Click()
   txtTicket = ""
   txtPedido = ""
   txtCliente = ""
   cmbStatus.ListIndex = 0
   txtDescricaoProblema = ""
   mskAbertura = "__/__/____"
   mskEncerramento = "__/__/____"
              
   CONSULTAR_REGISTROS
   
   CmdSql = "SELECT IFNULL(MAX(TICKET), 0) + 1 TICKET" & vbCr
   CmdSql = CmdSql & "FROM ATENDIMENTO"
   CMySql.Consulta CmdSql, RsConsulta
      
   If Not RsConsulta.EOF Then
      txtTicket = Format(Val(RsConsulta("TICKET")), "0000")
   End If
   
   txtTicket.SetFocus
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
            CmdSql = "DELETE FROM ATENDIMENTO" & vbCr
            CmdSql = CmdSql & "WHERE PROJETO    = " & PoeAspas(UCase(App.EXEName)) & vbCr
            CmdSql = CmdSql & "  AND TICKET     = " & PoeAspas(.ListItems(IL).SubItems(colTicket))
            CMySql.Executa CmdSql, True
         End Select
      Next IL
   End With
   
Cancelar:
   cmdExcluir.Enabled = True
   
   txtTicket.Enabled = True
   txtTicket = ""
   
   Form_Load
   txtTicket.Enabled = True
   txtTicket = ""
   txtTicket.SetFocus
   
End Sub

Private Sub cmdSalvar_Click()

  If Trim(txtTicket) = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O TICKET"
      txtTicket.SetFocus
      Exit Sub
   End If

   If Trim(txtPedido) = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O PEDIDO"
      txtPedido.SetFocus
      Exit Sub
   End If
   
   If Trim(txtCliente) = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O CLIENTE"
      txtCliente.SetFocus
      Exit Sub
   End If
            
   If cmbStatus = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O STATUS"
      cmbStatus.SetFocus
      Exit Sub
   End If
            
   If Trim(txtDescricaoProblema) = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER A DESCRIÇÃO DO PROBLEMA"
      txtDescricaoProblema.SetFocus
      Exit Sub
   End If
            
   If mskAbertura.ClipText = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER A DATA DA ABERTURA"
      mskAbertura.SetFocus
      Exit Sub
   End If
            
   cmdSalvar.Enabled = False
            
   CmdSql = "SELECT * FROM ATENDIMENTO" & vbCr
   CmdSql = CmdSql & "WHERE PROJETO    = " & PoeAspas(UCase(App.EXEName)) & vbCr
   CmdSql = CmdSql & "  AND TICKET     = " & PoeAspas(txtTicket) & vbCr
   CmdSql = CmdSql & "ORDER BY PROJETO,TICKET"
   CMySql.Consulta CmdSql, RsConsulta
      
   If RsConsulta.EOF Then
      CmdSql = "INSERT INTO ATENDIMENTO(PROJETO,TICKET,PEDIDO,CLIENTE,STATUS,DESCRICAO,ABERTURA,ENCERRAMENTO)" & vbCr
      CmdSql = CmdSql & "VALUES("
      CmdSql = CmdSql & PoeAspas(UCase(App.EXEName)) & ","
      CmdSql = CmdSql & PoeAspas(UCase(txtTicket)) & ","
      CmdSql = CmdSql & PoeAspas(UCase(txtPedido)) & ","
      CmdSql = CmdSql & PoeAspas(UCase(txtCliente)) & ","
      CmdSql = CmdSql & PoeAspas(UCase(cmbStatus)) & ","
      CmdSql = CmdSql & PoeAspas(UCase(txtDescricaoProblema)) & ","
      CmdSql = CmdSql & Format(mskAbertura, "YYYYMMDD") & ","
      
      If mskEncerramento.ClipText = "" Then
         CmdSql = CmdSql & 19000101 & ")"
      Else
         CmdSql = CmdSql & Format(mskEncerramento, "YYYYMMDD") & ")"
      End If
      
      CMySql.Executa CmdSql, True
   
      MsgBoxTabum Me, "REGISTRO INCLUÍDO COM SUCESSO"
   Else
      CmdSql = "UPDATE ATENDIMENTO SET PEDIDO       = " & PoeAspas(UCase(txtPedido)) & "," & vbCr
      CmdSql = CmdSql & "                       CLIENTE      = " & PoeAspas(UCase(txtCliente)) & "," & vbCr
      CmdSql = CmdSql & "                       STATUS       = " & PoeAspas(cmbStatus) & "," & vbCr
      CmdSql = CmdSql & "                       DESCRICAO    = " & PoeAspas(txtDescricaoProblema) & "," & vbCr
      CmdSql = CmdSql & "                       ABERTURA     = " & Format(mskAbertura, "YYYYMMDD")
      
      If mskEncerramento.ClipText <> "" Then
         CmdSql = CmdSql & "," & vbCr
         CmdSql = CmdSql & "                       ENCERRAMENTO = " & Format(mskEncerramento, "YYYYMMDD") & vbCr
      Else
         CmdSql = CmdSql & vbCr
      End If
      
      CmdSql = CmdSql & "WHERE PROJETO = " & PoeAspas(UCase(App.EXEName)) & vbCr
      CmdSql = CmdSql & "  AND TICKET  = " & PoeAspas(UCase(txtTicket))
      CMySql.Executa CmdSql, True
      
      MsgBoxTabum Me, "REGISTRO ALTERADO COM SUCESSO"
   End If
         
   cmdCancelar_Click
   cmdSalvar.Enabled = True
   txtTicket.Enabled = True
   txtTicket.SetFocus

End Sub

Private Sub Form_Load()
   MontaLvw
   
   cmbStatus.Clear
   cmbStatus.AddItem ""
   
   CmdSql = "SELECT * FROM TIPOCOMBO" & vbCr
   CmdSql = CmdSql & "WHERE PROJETO    = " & PoeAspas(UCase(App.EXEName)) & vbCr
   CmdSql = CmdSql & "  AND FORMULARIO = 'ATENDIMENTO'" & vbCr
   CmdSql = CmdSql & "  AND TIPO       = 'STATUS'" & vbCr
   CmdSql = CmdSql & "ORDER BY PROJETO,FORMULARIO,TIPO"
   CMySql.Consulta CmdSql, RsConsulta
   
   Do While Not RsConsulta.EOF
      cmbStatus.AddItem Trim(RsConsulta("DESCRICAO"))
   RsConsulta.MoveNext
   Loop
      
   If cmbStatus.ListCount > 0 Then cmbStatus.ListIndex = 0
   
   CONSULTAR_REGISTROS
   
   CmdSql = "SELECT IFNULL(MAX(TICKET), 0) + 1 TICKET" & vbCr
   CmdSql = CmdSql & "FROM ATENDIMENTO"
   CMySql.Consulta CmdSql, RsConsulta
      
   If Not RsConsulta.EOF Then
      txtTicket = Format(Val(RsConsulta("TICKET")), "0000")
   End If
   
End Sub

Private Sub lvwConsulta_DblClick()
   
   If lvwConsulta.ListItems.Count = 0 Then Exit Sub

   txtTicket = ""
   txtPedido = ""
   txtCliente = ""
   If cmbStatus.ListIndex > 0 Then cmbStatus.ListIndex = 0
   txtDescricaoProblema = ""
   mskAbertura = "__/__/____"
   mskEncerramento = "__/__/____"
         
   txtTicket = UCase(lvwConsulta.SelectedItem.SubItems(colTicket))
   txtTicket.Enabled = False
         
   txtPedido = lvwConsulta.SelectedItem.SubItems(colPedido)
   txtCliente = lvwConsulta.SelectedItem.SubItems(colCliente)
   cmbStatus = lvwConsulta.SelectedItem.SubItems(colStatusPed)
   txtDescricaoProblema = lvwConsulta.SelectedItem.SubItems(colDescricao)
   mskAbertura = lvwConsulta.SelectedItem.SubItems(colAbertura)
   
   If lvwConsulta.SelectedItem.SubItems(colEncerramento) <> "" Then
      mskEncerramento = lvwConsulta.SelectedItem.SubItems(colEncerramento)
   End If
   
   txtPedido.SetFocus

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

Private Sub mskAbertura_GotFocus()
   Marca Me
   mskAbertura.BackColor = QBColor(14)
End Sub

Private Sub mskAbertura_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub mskAbertura_LostFocus()
   mskAbertura.BackColor = QBColor(15)
End Sub

Private Sub mskEncerramento_GotFocus()
   Marca Me
   mskEncerramento.BackColor = QBColor(14)
End Sub

Private Sub mskEncerramento_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub mskEncerramento_LostFocus()
   mskEncerramento.BackColor = QBColor(15)
End Sub

Private Sub txtCliente_GotFocus()
   Marca Me
   txtCliente.BackColor = QBColor(14)
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub txtCliente_LostFocus()
   txtCliente.BackColor = QBColor(15)
End Sub

Private Sub txtDescricaoProblema_GotFocus()
   Marca Me
   txtDescricaoProblema.BackColor = QBColor(14)
End Sub

Private Sub txtDescricaoProblema_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub txtDescricaoProblema_LostFocus()
   txtDescricaoProblema.BackColor = QBColor(15)
End Sub

Private Sub txtPedido_GotFocus()
   Marca Me
   txtPedido.BackColor = QBColor(14)
End Sub

Private Sub txtPedido_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub txtPedido_LostFocus()
   txtPedido.BackColor = QBColor(15)
End Sub

Private Sub txtTicket_GotFocus()
   Marca Me
   txtTicket.BackColor = QBColor(14)
End Sub

Private Sub txtTicket_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub txtTicket_LostFocus()
   txtTicket.BackColor = QBColor(15)
End Sub

Private Sub CONSULTAR_REGISTROS()

   MontaLvw
   txtTicket.Enabled = True
     
   CmdSql = "SELECT *" & vbCr
   CmdSql = CmdSql & "FROM ATENDIMENTO" & vbCr
   CmdSql = CmdSql & "WHERE PROJETO = " & PoeAspas(UCase(App.EXEName)) & vbCr
   If Trim(txtTicket) <> "" Then CmdSql = CmdSql & "  AND TICKET = " & PoeAspas(txtTicket) & vbCr
   CmdSql = CmdSql & "ORDER BY PROJETO,TICKET DESC"
   CMySql.Consulta CmdSql, RsConsulta
      
   Do While Not RsConsulta.EOF
      Set Item = lvwConsulta.ListItems.Add(, , "")
      Item.SubItems(colTicket) = Trim(RsConsulta("TICKET"))
      Item.SubItems(colPedido) = Trim(RsConsulta("PEDIDO"))
      Item.SubItems(colCliente) = Trim(RsConsulta("CLIENTE"))
      Item.SubItems(colStatusPed) = Trim(RsConsulta("STATUS"))
      Item.SubItems(colDescricao) = Trim(RsConsulta("DESCRICAO"))
      Item.SubItems(colAbertura) = Trim(RsConsulta("ABERTURA"))
      Item.SubItems(colEncerramento) = IIf(Trim(RsConsulta("ENCERRAMENTO")) = "01/01/1900", "", Trim(RsConsulta("ENCERRAMENTO")))
      Item.SubItems(colStatus) = ""
   RsConsulta.MoveNext
   Loop

End Sub
