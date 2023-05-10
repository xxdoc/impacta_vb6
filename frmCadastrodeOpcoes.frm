VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCadastrodeOpcoes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Opções"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13515
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCadastrodeOpcoes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   13515
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraBotoes 
      BackColor       =   &H00FFFFFF&
      Height          =   1110
      Left            =   10680
      TabIndex        =   12
      Top             =   -60
      Width           =   2805
      Begin VB.CommandButton cmdLimpar 
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
         Height          =   900
         Left            =   1845
         Picture         =   "frmCadastrodeOpcoes.frx":1486
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpar Formulário"
         Top             =   150
         Width           =   900
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
         Height          =   900
         Left            =   930
         Picture         =   "frmCadastrodeOpcoes.frx":2008
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Excluir Registro"
         Top             =   150
         Width           =   900
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
         Height          =   900
         Left            =   15
         Picture         =   "frmCadastrodeOpcoes.frx":2B12
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salvar Registro"
         Top             =   150
         Width           =   900
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFFFF&
      Height          =   1110
      Left            =   45
      TabIndex        =   7
      Top             =   -60
      Width           =   10590
      Begin VB.ComboBox cmbFormulario 
         Height          =   360
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   465
         Width           =   3360
      End
      Begin VB.ComboBox cmbTipo 
         Height          =   360
         Left            =   3495
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   465
         Width           =   2745
      End
      Begin VB.TextBox txtDescricao 
         Height          =   360
         Left            =   6285
         MaxLength       =   40
         TabIndex        =   2
         Top             =   465
         Width           =   3915
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Formulário"
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   13
         Top             =   195
         Width           =   1005
      End
      Begin VB.Label lblDescricao 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "DESCRIÇÃO"
         Height          =   240
         Left            =   8805
         TabIndex        =   11
         Top             =   180
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "TIPO"
         Height          =   240
         Left            =   7680
         TabIndex        =   10
         Top             =   180
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipo"
         Height          =   240
         Index           =   0
         Left            =   3495
         TabIndex        =   9
         Top             =   195
         Width           =   420
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Descrição"
         Height          =   240
         Index           =   1
         Left            =   6270
         TabIndex        =   8
         Top             =   195
         Width           =   960
      End
   End
   Begin MSComctlLib.ListView lvwConsulta 
      Height          =   7230
      Left            =   0
      TabIndex        =   6
      Top             =   1065
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   12753
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
Attribute VB_Name = "frmCadastrodeOpcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Private RsConsulta               As New ADODB.Recordset
   Private Item                     As ListItem
   Private IL                       As Integer
   
   Private Const colProjeto         As Integer = 1
   Private Const colFormulario      As Integer = 2
   Private Const colTipo            As Integer = 3
   Private Const colDescricao       As Integer = 4
   Private Const colStatus          As Integer = 5
   
Private Sub MontaLvw()
   lvwConsulta.ListItems.Clear
   lvwConsulta.ColumnHeaders.Clear
   lvwConsulta.ColumnHeaders.Add , , "", 0
   lvwConsulta.ColumnHeaders.Add , , "Projeto", 1800
   lvwConsulta.ColumnHeaders.Add , , "Formulário", 4500
   lvwConsulta.ColumnHeaders.Add , , "Tipo", 2600
   lvwConsulta.ColumnHeaders.Add , , "Descrição", 3300
   lvwConsulta.ColumnHeaders.Add , , "Status", 1000
   lvwConsulta.View = lvwReport
End Sub

Private Sub cmbFormulario_GotFocus()
   cmbFormulario.BackColor = QBColor(14)
End Sub

Private Sub cmbFormulario_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Sendkeys "{tab}"
   End If
End Sub

Private Sub cmbFormulario_LostFocus()
   CONSULTAR_REGISTROS
   cmbFormulario.BackColor = QBColor(15)
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

Private Sub cmdExcluir_Click()
   
   If VERIFICAR_DADOS_BASE = False Then Exit Sub

   If MsgBox("Deseja Realmente Excluir Dado Selecionado?", vbYesNo + vbQuestion + vbDefaultButton2, "Exclusão") = vbNo Then
      Exit Sub
   End If
   
   cmdExcluir.Enabled = False
    
   With lvwConsulta
      For IL = 1 To .ListItems.Count
         Select Case .ListItems(IL).SubItems(colStatus)
         Case "D"
            CmdSql = "DELETE FROM TIPOCOMBO" & vbCr
            CmdSql = CmdSql & "WHERE PROJETO    = " & PoeAspas(UCase(App.EXEName)) & vbCr
            CmdSql = CmdSql & "  AND FORMULARIO = " & PoeAspas(.ListItems(IL).SubItems(colFormulario)) & vbCr
            CmdSql = CmdSql & "  AND TIPO       = " & PoeAspas(.ListItems(IL).SubItems(colTipo)) & vbCr
            CmdSql = CmdSql & "  AND DESCRICAO  = " & PoeAspas(.ListItems(IL).SubItems(colDescricao))
            CMySql.Executa CmdSql, True
         End Select
      Next IL
   End With
   
Cancelar:
   cmdExcluir.Enabled = True
   
   Form_Load
   cmbTipo.SetFocus

End Sub

Private Sub cmdLimpar_Click()
   lblTipo = ""
   lblDescricao = ""
   
   If cmbFormulario.ListCount > 0 Then cmbFormulario.ListIndex = 0
   If cmbTipo.ListCount > 0 Then cmbTipo.ListIndex = 0
   txtDescricao = ""
   
   CONSULTAR_REGISTROS
   cmbFormulario.SetFocus
End Sub

Private Sub cmdSalvar_Click()
   
   If cmbFormulario = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O FORMULÁRIO"
      cmbFormulario.SetFocus
      Exit Sub
   End If

   If cmbTipo = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O TIPO DE INFORMAÇÃO"
      cmbTipo.SetFocus
      Exit Sub
   End If
            
   If Trim(txtDescricao) = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER A DESCRIÇÃO"
      txtDescricao.SetFocus
      Exit Sub
   End If
   
   cmdSalvar.Enabled = False
            
   CmdSql = "SELECT * FROM TIPOCOMBO" & vbCr
   CmdSql = CmdSql & "WHERE PROJETO    = " & PoeAspas(UCase(App.EXEName)) & vbCr
   CmdSql = CmdSql & "  AND FORMULARIO = " & PoeAspas(cmbFormulario) & vbCr
   CmdSql = CmdSql & "  AND TIPO       = " & PoeAspas(cmbTipo) & vbCr
   CmdSql = CmdSql & "  AND DESCRICAO  = " & PoeAspas(txtDescricao) & vbCr
   CmdSql = CmdSql & "ORDER BY PROJETO,FORMULARIO,TIPO,DESCRICAO"
   CMySql.Consulta CmdSql, RsConsulta
   
   If RsConsulta.EOF Then
      CmdSql = "INSERT INTO TIPOCOMBO(PROJETO,FORMULARIO,TIPO,DESCRICAO)" & vbCr
      CmdSql = CmdSql & "VALUES("
      CmdSql = CmdSql & PoeAspas(UCase(App.EXEName)) & ","
      CmdSql = CmdSql & PoeAspas(UCase(cmbFormulario)) & ","
      CmdSql = CmdSql & PoeAspas(UCase(cmbTipo)) & ","
      CmdSql = CmdSql & PoeAspas(txtDescricao) & ")"
      CMySql.Executa CmdSql, True
   
      MsgBoxTabum Me, "REGISTRO INCLUÍDO COM SUCESSO"
   Else
      MsgBoxTabum Me, "REGISTRO JÁ EXISTENTE"
   End If
      
   lblTipo = ""
   lblDescricao = ""
   txtDescricao = ""
      
   cmdSalvar.Enabled = True
   CONSULTAR_REGISTROS
   txtDescricao.SetFocus
   
End Sub

Private Sub Form_Load()
   MontaLvw
   
   cmbFormulario.Clear
   cmbFormulario.AddItem ""
   cmbFormulario.AddItem "CADASTRO DE USUÁRIO"
   cmbFormulario.AddItem "ATENDIMENTO"
   cmbFormulario.AddItem "CUSTO DO PRODUTO"
   
   cmbTipo.Clear
   cmbTipo.AddItem ""
   cmbTipo.AddItem "PERFIL"
   cmbTipo.AddItem "STATUS"
   
'  CUSTO DO PRODUTO
   cmbTipo.AddItem "GATEWAY"
   cmbTipo.AddItem "IOF"
   cmbTipo.AddItem "PARCELAMENTO"
   cmbTipo.AddItem "OUTROS CUSTOS"
   cmbTipo.AddItem "IMPOSTO"
   cmbTipo.AddItem "MARKETING"
   
   CONSULTAR_REGISTROS
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

Private Sub txtDescricao_GotFocus()
   Marca Me
   txtDescricao.BackColor = QBColor(14)
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
      Sendkeys "{tab}"
   End If
End Sub

Private Sub txtDescricao_LostFocus()
   txtDescricao.BackColor = QBColor(15)
End Sub

Private Sub CONSULTAR_REGISTROS()

   MontaLvw
   
   CmdSql = "SELECT *" & vbCr
   CmdSql = CmdSql & "FROM TIPOCOMBO" & vbCr
   CmdSql = CmdSql & "WHERE PROJETO = " & PoeAspas(UCase(App.EXEName)) & vbCr
   If cmbFormulario <> "" Then CmdSql = CmdSql & "  AND FORMULARIO = " & PoeAspas(cmbFormulario) & vbCr
   If cmbTipo <> "" Then CmdSql = CmdSql & "  AND TIPO       = " & PoeAspas(cmbTipo) & vbCr
   CmdSql = CmdSql & "ORDER BY PROJETO,FORMULARIO,TIPO,DESCRICAO"
   CMySql.Consulta CmdSql, RsConsulta
      
   Do While Not RsConsulta.EOF
      Set Item = lvwConsulta.ListItems.Add(, , "")
      Item.SubItems(colProjeto) = Trim(RsConsulta("PROJETO"))
      Item.SubItems(colFormulario) = Trim(RsConsulta("FORMULARIO"))
      Item.SubItems(colTipo) = Trim(RsConsulta("TIPO"))
      Item.SubItems(colDescricao) = Trim(RsConsulta("DESCRICAO"))
   RsConsulta.MoveNext
   Loop

End Sub

Private Function VERIFICAR_DADOS_BASE() As Boolean
   VERIFICAR_DADOS_BASE = True
   
   CmdSql = ""
   
   With lvwConsulta
      For IL = 1 To .ListItems.Count
         Select Case .ListItems(IL).SubItems(colStatus)
         Case "D"
            Select Case .ListItems(IL).SubItems(colTipo)
            Case "PERFIL"
               CmdSql = "SELECT * FROM ACESSO" & vbCr
               CmdSql = CmdSql & "WHERE PERFIL = " & PoeAspas(.ListItems(IL).SubItems(colDescricao))
            Case "STATUS"
               CmdSql = "SELECT * FROM ACESSO" & vbCr
               CmdSql = CmdSql & "WHERE STATUS = " & PoeAspas(.ListItems(IL).SubItems(colDescricao))
            End Select
            
            If CmdSql <> "" Then
               CMySql.Consulta CmdSql, RsConsulta
               
               If Not RsConsulta.EOF Then
                  MsgBoxTabum Me, "HÁ REGISTRO COM ESTE DADO, NÃO PODE SER EXCLUÍDO" & vbCr & _
                                  "TIPO: " & .ListItems(IL).SubItems(colTipo) & vbCr & _
                                  "DESCRIÇÃO: " & .ListItems(IL).SubItems(colDescricao)
                  
                  .ListItems(IL).SubItems(colStatus) = ""
                  VERIFICAR_DADOS_BASE = False
               End If
            End If
         End Select
      Next IL
   End With
  
End Function
