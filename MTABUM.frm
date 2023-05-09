VERSION 5.00
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Tabum - Uma explosão de soluções para sua empresa."
   ClientHeight    =   11940
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   16590
   Icon            =   "MTABUM.frx":0000
   Picture         =   "MTABUM.frx":1486
   WindowState     =   2  'Maximized
   Begin VB.Menu MnuCadastro 
      Caption         =   "Cadastros"
      Begin VB.Menu MnuCadastrodeOpcoes 
         Caption         =   "AC1 - Cadastro de Opções"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MnuCadastroUsuario 
         Caption         =   "AC1 - Cadastro de Usuário"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MnuCadastroCliente 
         Caption         =   "AC2 - Cadastro de Cliente"
         Shortcut        =   {F3}
      End
      Begin VB.Menu MnuTraco01 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSair 
         Caption         =   "Sair do Sistema"
      End
   End
   Begin VB.Menu MnuAtividades 
      Caption         =   "Atividades"
      Begin VB.Menu MnuAtendimentoSite 
         Caption         =   "AC2 - Atendimento ao Cliente"
         Shortcut        =   {F4}
      End
      Begin VB.Menu MnuControleFinanceiro 
         Caption         =   "AC3 - Controle Financeiro"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MnuCustodoProduto 
         Caption         =   "AC3 - Custo do Produto"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu MnuAjuda 
      Caption         =   "Ajuda"
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
   Me.Caption = Me.Caption & " / Usuário: " & sysUsuario & " / Perfil: " & sysPerfil
End Sub

Private Sub MnuAjuda_Click()
   ShowForm frmAjuda
End Sub

Private Sub MnuAtendimentoSite_Click()
   ShowForm frmAtendimentoCliente
End Sub

Private Sub MnuCadastroCliente_Click()
   ShowForm frmCadastroCliente
End Sub

Private Sub MnuCadastrodeOpcoes_Click()
   If UCase(sysPerfil) = "ADMINISTRADOR" Then
      ShowForm frmCadastrodeOpcoes
   Else
      MsgBox "ACESSO NÃO AUTORIZADO", 64, "SEM ACESSO"
      Exit Sub
   End If
End Sub

Private Sub MnuCadastroUsuario_Click()
   ShowForm frmCadastroUsuario
End Sub

Private Sub MnuControleFinanceiro_Click()
   ShowForm frmControleFinanceiro
End Sub

Private Sub MnuCustodoProduto_Click()
   ShowForm frmCustoProduto
End Sub

Private Sub MnuSair_Click()
   End
End Sub


