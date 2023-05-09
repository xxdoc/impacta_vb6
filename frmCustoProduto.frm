VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustoProduto 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custo do Produto"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustoProduto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   10185
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
      Height          =   3795
      Left            =   9060
      TabIndex        =   43
      Top             =   0
      Width           =   1065
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
         Picture         =   "frmCustoProduto.frx":1486
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Picture         =   "frmCustoProduto.frx":1F90
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Salvar Registro"
         Top             =   1950
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
         Picture         =   "frmCustoProduto.frx":2A9A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Excluir Registro"
         Top             =   2835
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
         Left            =   45
         Picture         =   "frmCustoProduto.frx":35A4
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Limpar Campos"
         Top             =   1065
         Width           =   950
      End
   End
   Begin VB.Frame fraTotalCustos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total (Produto + Taxas + Marketing)"
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
      Height          =   750
      Left            =   4035
      TabIndex        =   41
      Top             =   4185
      Width           =   4950
      Begin VB.TextBox txtTotalCustoOperacional 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2940
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   19
         Text            =   "0,00"
         Top             =   285
         Width           =   1725
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total Custo Operacional"
         Height          =   240
         Index           =   11
         Left            =   165
         TabIndex        =   42
         Top             =   285
         Width           =   2385
      End
   End
   Begin VB.Frame fraMarketing 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Marketing"
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
      Height          =   750
      Left            =   75
      TabIndex        =   38
      Top             =   4185
      Width           =   3855
      Begin VB.TextBox txtPorcentagemMarketingAPagar 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2445
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   18
         Text            =   "0,00"
         Top             =   285
         Width           =   1290
      End
      Begin VB.TextBox txtPorcentagemMarketing 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1410
         MaxLength       =   6
         TabIndex        =   17
         Text            =   "0,00"
         Top             =   285
         Width           =   960
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "% Marketing"
         Height          =   240
         Index           =   14
         Left            =   105
         TabIndex        =   39
         Top             =   285
         Width           =   1230
      End
   End
   Begin VB.Frame fraTaxa 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Custos e Taxas Fixas"
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
      Height          =   3120
      Left            =   4470
      TabIndex        =   33
      Top             =   1050
      Width           =   4515
      Begin VB.TextBox txtOutrosCustosAPagar 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2910
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   2031
         Width           =   1455
      End
      Begin VB.TextBox txtParcelamentoAPagar 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2910
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   14
         Text            =   "0,00"
         Top             =   1454
         Width           =   1455
      End
      Begin VB.TextBox txtIOFAPagar 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2910
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   13
         Text            =   "0,00"
         Top             =   877
         Width           =   1455
      End
      Begin VB.TextBox txtGatewayAPagar 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2910
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   12
         Text            =   "0,00"
         Top             =   300
         Width           =   1455
      End
      Begin VB.TextBox txtImpostoAPagar 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2895
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   16
         Text            =   "0,00"
         Top             =   2610
         Width           =   1455
      End
      Begin VB.TextBox txtImposto 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1695
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   11
         Text            =   "0,00"
         Top             =   2610
         Width           =   1065
      End
      Begin VB.TextBox txtGateway 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1695
         MaxLength       =   12
         TabIndex        =   7
         Text            =   "0,00"
         Top             =   300
         Width           =   1065
      End
      Begin VB.TextBox txtIOF 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1695
         MaxLength       =   12
         TabIndex        =   8
         Text            =   "0,00"
         Top             =   877
         Width           =   1065
      End
      Begin VB.TextBox txtParcelamento 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1695
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   9
         Text            =   "0,00"
         Top             =   1454
         Width           =   1065
      End
      Begin VB.TextBox txtOutrosCustos 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1695
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   10
         Text            =   "0,00"
         Top             =   2031
         Width           =   1065
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gateway"
         Height          =   240
         Index           =   9
         Left            =   120
         TabIndex        =   40
         Top             =   300
         Width           =   870
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Imposto"
         Height          =   240
         Index           =   10
         Left            =   120
         TabIndex        =   37
         Top             =   2610
         Width           =   795
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "IOF"
         Height          =   240
         Index           =   8
         Left            =   120
         TabIndex        =   36
         Top             =   870
         Width           =   345
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Parcelamento"
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   35
         Top             =   1455
         Width           =   1335
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Outros Custos"
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   34
         Top             =   2025
         Width           =   1410
      End
   End
   Begin VB.Frame fraProduto 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dados do Produto"
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
      Height          =   1050
      Left            =   75
      TabIndex        =   30
      Top             =   0
      Width           =   8910
      Begin VB.TextBox txtProduto 
         Height          =   360
         Left            =   165
         MaxLength       =   200
         TabIndex        =   0
         Top             =   525
         Width           =   6555
      End
      Begin VB.TextBox txtPrecoPraticado 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   6870
         MaxLength       =   12
         TabIndex        =   1
         Text            =   "0,00"
         Top             =   525
         Width           =   1860
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Produto"
         Height          =   240
         Index           =   7
         Left            =   165
         TabIndex        =   32
         Top             =   255
         Width           =   765
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Preço Praticado"
         Height          =   240
         Index           =   6
         Left            =   6870
         TabIndex        =   31
         Top             =   255
         Width           =   1560
      End
   End
   Begin VB.Frame fraCusto 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Custo do Produto"
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
      Height          =   3120
      Left            =   75
      TabIndex        =   26
      Top             =   1050
      Width           =   4305
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "0,00"
         Top             =   2625
         Width           =   1440
      End
      Begin VB.TextBox txtPrecoFornecedor 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "0,00"
         Top             =   285
         Width           =   1440
      End
      Begin VB.TextBox txtFreteFornecedor 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "0,00"
         Top             =   870
         Width           =   1440
      End
      Begin VB.TextBox txtTotalCustoFornecedor 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "0,00"
         Top             =   1455
         Width           =   1440
      End
      Begin VB.TextBox txtPrecoRecomendado 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "0,00"
         Top             =   2040
         Width           =   1440
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "VER QUAL CAMPO"
         Height          =   240
         Index           =   12
         Left            =   135
         TabIndex        =   45
         Top             =   2625
         Width           =   1755
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Preço Fornecedor"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   44
         Top             =   285
         Width           =   1740
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Preço Recomendado"
         Height          =   240
         Index           =   3
         Left            =   150
         TabIndex        =   29
         Top             =   2040
         Width           =   1995
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total Custo Fornecedor"
         Height          =   240
         Index           =   2
         Left            =   165
         TabIndex        =   28
         Top             =   1455
         Width           =   2340
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Frete Fornecedor"
         Height          =   240
         Index           =   1
         Left            =   165
         TabIndex        =   27
         Top             =   870
         Width           =   1710
      End
   End
   Begin MSComctlLib.ListView lvwLucro 
      Height          =   1905
      Left            =   15
      TabIndex        =   24
      Top             =   4935
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   3360
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwConsulta 
      Height          =   1920
      Left            =   0
      TabIndex        =   25
      Top             =   6855
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   3387
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmCustoProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

   Private RsConsulta               As New ADODB.Recordset
   Private Item                     As ListItem
   Private IL                       As Integer
   
   Private Const colProjeto         As Integer = 1
   Private Const colProduto         As Integer = 2
   Private Const colPrecoPraticado  As Integer = 3
   Private Const colPrecoFornecedor As Integer = 4
   Private Const colFreteFornecedor As Integer = 5
   Private Const colTotalCusto      As Integer = 6
   Private Const colPrecoRecomend   As Integer = 7
   Private Const colStatus          As Integer = 8

Private Sub MontaLvwUnidades()
   lvwLucro.ListItems.Clear
   lvwLucro.ColumnHeaders.Clear
   lvwLucro.ColumnHeaders.Add , , "", 0
   lvwLucro.ColumnHeaders.Add , , "Qtde Unidades Vendidas", 3000, lvwColumnRight
   lvwLucro.ColumnHeaders.Add , , "Lucro", 1500, lvwColumnRight
   lvwLucro.View = lvwReport
End Sub

Private Sub MontaLvw()
   lvwConsulta.ListItems.Clear
   lvwConsulta.ColumnHeaders.Clear
   lvwConsulta.ColumnHeaders.Add , , "", 0
   lvwConsulta.ColumnHeaders.Add , , "Projeto", 0
   lvwConsulta.ColumnHeaders.Add , , "Produto", 3000
   lvwConsulta.ColumnHeaders.Add , , "Preço Praticado", 2000
   lvwConsulta.ColumnHeaders.Add , , "Preço Fornecedor", 2000
   lvwConsulta.ColumnHeaders.Add , , "Frete Fornecedor", 2000
   lvwConsulta.ColumnHeaders.Add , , "Total Custo", 2000
   lvwConsulta.ColumnHeaders.Add , , "Preço Recomendado", 2000
   lvwConsulta.ColumnHeaders.Add , , "Status", 800
   lvwConsulta.View = lvwReport
End Sub

Private Sub cmdCancelar_Click()

   MontaLvw
   MontaLvwUnidades

   txtProduto = ""
   txtPrecoPraticado = "0,00"
   
   txtPrecoFornecedor = "0,00"
   txtFreteFornecedor = "0,00"
   txtTotalCustoFornecedor = "0,00"
   txtPrecoRecomendado = "0,00"
   
   txtGateway = "0,00"
   txtGatewayAPagar = "0,00"
   txtIOF = "0,00"
   txtIOFAPagar = "0,00"
   txtParcelamento = "0,00"
   txtParcelamentoAPagar = "0,00"
   txtOutrosCustos = "0,00"
   txtOutrosCustosAPagar = "0,00"
   txtImposto = "0,00"
   txtImpostoAPagar = "0,00"
   
   txtPorcentagemMarketing = "0,00"
   txtPorcentagemMarketingAPagar = "0,00"
   
   txtTotalCustoOperacional = "0,00"
   
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
            CmdSql = CmdSql & "WHERE PROJETO = " & PoeAspas(UCase(App.EXEName)) & vbCr
            CmdSql = CmdSql & "  AND PRODUTO = " & PoeAspas(.ListItems(IL).SubItems(colProduto))
            CMySql.Executa CmdSql, True
         End Select
      Next IL
   End With
   
Cancelar:
   cmdExcluir.Enabled = True
      
   Form_Load
   
   txtProduto = ""
   txtProduto.SetFocus

End Sub

Private Sub cmdSalvar_Click()

   If Trim(txtProduto) = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O PRODUTO"
      txtProduto.SetFocus
      Exit Sub
   End If

   If Val(txtPrecoPraticado) = 0 Then
      MsgBoxTabum Me, "FAVOR PREENCHER O PREÇO PRATICADO"
      txtPrecoPraticado.SetFocus
      Exit Sub
   End If
   
   If Val(txtPrecoFornecedor) = 0 Then
      MsgBoxTabum Me, "FAVOR PREENCHER O PREÇO DO FORNECEDOR"
      txtPrecoFornecedor.SetFocus
      Exit Sub
   End If
   
   If Val(txtFreteFornecedor) = 0 Then
      MsgBoxTabum Me, "FAVOR PREENCHER O PREÇO DO FRETE"
      txtFreteFornecedor.SetFocus
      Exit Sub
   End If
               
   cmdSalvar.Enabled = False
           
   CmdSql = "SELECT * FROM CUSTO_PRODUTO" & vbCr
   CmdSql = CmdSql & "WHERE PROJETO = " & PoeAspas(UCase(App.EXEName)) & vbCr
   CmdSql = CmdSql & "  AND PRODUTO = " & PoeAspas(txtProduto) & vbCr
   CmdSql = CmdSql & "ORDER BY PROJETO,PRODUTO"
   CMySql.Consulta CmdSql, RsConsulta
      
   If RsConsulta.EOF Then
      CmdSql = "INSERT INTO CUSTO_PRODUTO(PROJETO,PRODUTO,PRECO_PRATICADO,PRECO_FORNECEDOR,FRETE_FORNECEDOR,"
      CmdSql = CmdSql & "GATEWAY,IOF,PARCELAMENTO,OUTROS_CUSTOS,IMPOSTO,MARKETING)" & vbCr
      CmdSql = CmdSql & "VALUES("
      CmdSql = CmdSql & PoeAspas(UCase(App.EXEName)) & ","
      CmdSql = CmdSql & PoeAspas(UCase(txtProduto)) & ","
      CmdSql = CmdSql & Str(txtPrecoPraticado) & ","
      CmdSql = CmdSql & Str(txtPrecoFornecedor) & ","
      CmdSql = CmdSql & Str(txtFreteFornecedor) & ","
      CmdSql = CmdSql & Str(txtGateway) & ","
      CmdSql = CmdSql & Str(txtIOF) & ","
      CmdSql = CmdSql & Str(txtParcelamento) & ","
      CmdSql = CmdSql & Str(txtOutrosCustos) & ","
      CmdSql = CmdSql & Str(txtImposto) & ","
      CmdSql = CmdSql & Str(txtPorcentagemMarketing) & ")"
      CMySql.Executa CmdSql, True
   
      MsgBoxTabum Me, "REGISTRO INCLUÍDO COM SUCESSO"
   Else
      
      CmdSql = "UPDATE SET PRECO_PRATICADO  = " & Str(txtPrecoPraticado) & "," & vbCr
      CmdSql = CmdSql & "           PRECO_FORNECEDOR = " & Str(txtPrecoFornecedor) & "," & vbCr
      CmdSql = CmdSql & "           FRETE_FORNECEDOR = " & Str(txtFreteFornecedor) & "," & vbCr
      CmdSql = CmdSql & "           GATEWAY          = " & Str(txtGateway) & "," & vbCr
      CmdSql = CmdSql & "           IOF              = " & Str(txtIOF) & "," & vbCr
      CmdSql = CmdSql & "           PARCELAMENTO     = " & Str(txtParcelamento) & "," & vbCr
      CmdSql = CmdSql & "           OUTROS_CUSTOS    = " & Str(txtOutrosCustos) & "," & vbCr
      CmdSql = CmdSql & "           IMPOSTO          = " & Str(txtImposto) & "," & vbCr
      CmdSql = CmdSql & "           MARKETING        = " & Str(txtPorcentagemMarketing) & vbCr
      CmdSql = CmdSql & "WHERE PROJETO = " & PoeAspas(UCase(App.EXEName)) & vbCr
      CmdSql = CmdSql & "  AND PRODUTO = " & PoeAspas(UCase(txtProduto)) & ","
      CMySql.Executa CmdSql, True
      
      MsgBoxTabum Me, "REGISTRO ALTERADO COM SUCESSO"
   End If
         
   cmdCancelar_Click
   cmdSalvar.Enabled = True
   txtProduto.Enabled = True
   txtProduto.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = vbKeyReturn Then
         Sendkeys "{tab}"
   End If
End Sub

Private Sub Form_Load()
   CONSULTAR_REGISTROS
End Sub

Private Sub lvwConsulta_DblClick()
   
   If lvwConsulta.ListItems.Count = 0 Then Exit Sub
   
   txtProduto = ""
   txtPrecoPraticado = "0,00"
   
   txtPrecoFornecedor = "0,00"
   txtFreteFornecedor = "0,00"
   txtTotalCustoFornecedor = "0,00"
   txtPrecoRecomendado = "0,00"
   
   txtGateway = "0,00"
   txtGatewayAPagar = "0,00"
   txtIOF = "0,00"
   txtIOFAPagar = "0,00"
   txtParcelamento = "0,00"
   txtParcelamentoAPagar = "0,00"
   txtOutrosCustos = "0,00"
   txtOutrosCustosAPagar = "0,00"
   txtImposto = "0,00"
   txtImpostoAPagar = "0,00"
   
   txtPorcentagemMarketing = "0,00"
   txtPorcentagemMarketingAPagar = "0,00"
   
   txtTotalCustoOperacional = "0,00"
         
   txtProduto = UCase(lvwConsulta.SelectedItem.SubItems(colProduto))
   txtProduto.Enabled = False
         
   txtPrecoPraticado = lvwConsulta.SelectedItem.SubItems(colPrecoPraticado)
   txtPrecoFornecedor = lvwConsulta.SelectedItem.SubItems(colPrecoFornecedor)
   txtFreteFornecedor = lvwConsulta.SelectedItem.SubItems(colFreteFornecedor)
   txtTotalCustoFornecedor = lvwConsulta.SelectedItem.SubItems(colTotalCusto)
   txtPrecoRecomendado = lvwConsulta.SelectedItem.SubItems(colPrecoRecomend)
      
   txtPrecoPraticado.SetFocus

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

Private Sub txtFreteFornecedor_GotFocus()
   Marca Me
End Sub

Private Sub txtFreteFornecedor_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtFreteFornecedor, KeyAscii
End Sub

Private Sub txtFreteFornecedor_LostFocus()
   If Trim(txtFreteFornecedor) = "" Then txtFreteFornecedor = "0,00"
End Sub

Private Sub txtGateway_GotFocus()
   Marca Me
End Sub

Private Sub txtGateway_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtGateway, KeyAscii
End Sub

Private Sub txtGateway_LostFocus()
   If Trim(txtGateway) = "" Then txtGateway = "0,00"
End Sub

Private Sub txtImposto_GotFocus()
   Marca Me
End Sub

Private Sub txtImposto_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtImposto, KeyAscii
End Sub

Private Sub txtImposto_LostFocus()
   If Trim(txtImposto) = "" Then txtImposto = "0,00"
End Sub

Private Sub txtIOF_GotFocus()
   Marca Me
End Sub

Private Sub txtIOF_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtIOF, KeyAscii
End Sub

Private Sub txtIOF_LostFocus()
   If Trim(txtIOF) = "" Then txtIOF = "0,00"
End Sub

Private Sub txtOutrosCustos_GotFocus()
   Marca Me
End Sub

Private Sub txtOutrosCustos_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtOutrosCustos, KeyAscii
End Sub

Private Sub txtOutrosCustos_LostFocus()
   If Trim(txtOutrosCustos) = "" Then txtOutrosCustos = "0,00"
End Sub

Private Sub txtParcelamento_GotFocus()
   Marca Me
End Sub

Private Sub txtParcelamento_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtParcelamento, KeyAscii
End Sub

Private Sub txtParcelamento_LostFocus()
   If Trim(txtParcelamento) = "" Then txtParcelamento = "0,00"
End Sub

Private Sub txtPorcentagemMarketing_GotFocus()
   Marca Me
End Sub

Private Sub txtPorcentagemMarketing_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtPorcentagemMarketing, KeyAscii
End Sub

Private Sub txtPorcentagemMarketing_LostFocus()
   If Trim(txtPorcentagemMarketing) = "" Then txtPorcentagemMarketing = "0,00"
End Sub

Private Sub txtPrecoFornecedor_GotFocus()
   Marca Me
End Sub

Private Sub txtPrecoFornecedor_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtPrecoFornecedor, KeyAscii
End Sub

Private Sub txtPrecoFornecedor_LostFocus()
   If Trim(txtPrecoFornecedor) = "" Then txtPrecoFornecedor = "0,00"
End Sub

Private Sub txtPrecoPraticado_GotFocus()
   Marca Me
End Sub

Private Sub txtPrecoPraticado_KeyPress(KeyAscii As Integer)
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii < (0) Then KeyAscii = 0
   If KeyAscii = 46 Then KeyAscii = 44
   SoNumeros txtPrecoPraticado, KeyAscii
End Sub


Private Sub txtPrecoPraticado_LostFocus()
   If Not IsNumeric(txtPrecoPraticado.Text) Then txtPrecoPraticado.Text = 0
   txtPrecoPraticado.Text = Format((txtPrecoPraticado.Text), "###,##0.00")
End Sub

Private Sub txtProduto_GotFocus()
   Marca Me
End Sub

Private Sub CONSULTAR_REGISTROS()

   MontaLvw
   MontaLvwUnidades
   
   txtProduto.Enabled = True
       
   CmdSql = "SELECT *" & vbCr
   CmdSql = CmdSql & "FROM CUSTO_PRODUTO" & vbCr
   CmdSql = CmdSql & "WHERE PROJETO = " & PoeAspas(UCase(App.EXEName)) & vbCr
   If Trim(txtProduto) <> "" Then CmdSql = CmdSql & "  AND PRODUTO LIKE " & PoeAspas("%" & Trim(txtProduto) & "%") & vbCr
   CmdSql = CmdSql & "ORDER BY PROJETO,PRODUTO"
   CMySql.Consulta CmdSql, RsConsulta
      
   Do While Not RsConsulta.EOF
      Set Item = lvwConsulta.ListItems.Add(, , "")
      Item.SubItems(colProjeto) = Trim(RsConsulta("PROJETO"))
      Item.SubItems(colProduto) = Trim(RsConsulta("PRODUTO"))
      Item.SubItems(colPrecoPraticado) = Trim(RsConsulta("PRECO_PRATICADO"))
      Item.SubItems(colPrecoFornecedor) = Trim(RsConsulta("PRECO_FORNECEDOR"))
      Item.SubItems(colFreteFornecedor) = Trim(RsConsulta("FRETE_FORNECEDOR"))
      Item.SubItems(colTotalCusto) = "CÁLCULO"
      Item.SubItems(colPrecoRecomend) = "CÁLCULO"
      Item.SubItems(colStatus) = ""
   RsConsulta.MoveNext
   Loop

End Sub

Private Function CALCULA_CUSTOS()
   txtTotalCustoFornecedor = txtPrecoFornecedor + txtFreteFornecedor
   txtPrecoRecomendado = txtTotalCustoFornecedor * 3
   
   
End Function
