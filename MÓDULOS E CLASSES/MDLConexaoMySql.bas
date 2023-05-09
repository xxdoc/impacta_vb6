Attribute VB_Name = "MDLConexaoMySql"
Option Explicit
   Private RsMysql   As ADODB.Recordset
   Public CMySql     As New cConMysql
   Dim mblnAddMode   As Boolean
   Dim PathFull      As String
   Dim ArqDir        As String
   Dim RegAux        As String
   Dim ArqMySql()    As String
   Dim NroArq        As Integer
   Dim K             As Integer
   Dim CmdSql        As String
   
Public Function CriaTab_MySql(NomeServidor As String, _
                   ParamArray NomeProjeto() As Variant)
   On Error Resume Next
   Dim I          As Integer
   Dim CmdSql     As String
   Dim NomeProj
   
   For Each NomeProj In NomeProjeto
      ReDim ArqMySql(100)
   
      NroArq = 0
      PathFull = "C:\ARQUIVOS GERAIS\PROGRAMAS\MTABUM\RELSCRIPTS\ScriptsMySql\" & NomeProj & "\"
      
      ArqDir = Dir(PathFull & "*.MYS", vbNormal)
            
         If ArqDir <> "" Then
         Do While ArqDir <> ""
            NroArq = NroArq + 1
            ArqMySql(NroArq) = ArqDir
            ArqDir = Dir
         Loop
         
         For K = 1 To NroArq                             ' - 4  ==> TIRA A EXTENSAO DO ARQUIVO
            If Dir("C:\MYSQL\DATA\PADRAO\" & Mid(ArqMySql(K), 1, Len(Trim(ArqMySql(K))) - 4) & ".*", vbNormal) <> "" Then
               CmdSql = "FLUSH TABLES"
               CMySql.Executa CmdSql, True
               
               CmdSql = "UNLOCK TABLES"
               CMySql.Executa CmdSql, True
               
               Kill "C:\MYSQL\DATA\PADRAO\" & Mid(ArqMySql(K), 1, Len(Trim(ArqMySql(K))) - 4) & ".*"
            End If
            
            Open PathFull & ArqMySql(K) For Input As #1
            
            CmdSql = Empty
            
            Do While Not EOF(1)
               Line Input #1, RegAux
               If InStr(1, RegAux, ";", vbBinaryCompare) > 0 Then
                  CmdSql = CmdSql & Mid(Trim(RegAux), 1, Len(Trim(RegAux)) - 1) & vbCrLf
                  CMySql.Executa CmdSql
                  CmdSql = Empty
               Else
                  CmdSql = CmdSql & Trim(RegAux) & vbCrLf
               End If
            Loop
            
            Close #1
            
            CMySql.Executa CmdSql
         Next K
      End If
   Next
End Function

Public Function CriaTab_MySql_Server(NomeProjeto As String, USUARIO As String)
   ReDim ArqMySql(100)
   
   NroArq = 0
   
   PathFull = CMySql.MaquinaLocal & "\RELSCRIPTS\SCRIPTSMYSQL\" & NomeProjeto & "\"
                           
   ArqDir = Dir(PathFull & "*.MYS", vbNormal)
         
   If ArqDir <> "" Then
      Do While ArqDir <> ""
         NroArq = NroArq + 1
         ArqMySql(NroArq) = ArqDir
         ArqDir = Dir
      Loop
      
      For K = 1 To NroArq                             ' - 4  ==> TIRA A EXTENSAO DO ARQUIVO
         If Dir("C:\MYSQL\DATA\PADRAO\" & Mid(ArqMySql(K), 1, Len(Trim(ArqMySql(K))) - 4) & ".*", vbNormal) <> "" Then
            CmdSql = "DELETE FROM " & Mid(ArqMySql(K), 1, Len(Trim(ArqMySql(K))) - 4) & _
               " WHERE USUARIO = '" & UCase(USUARIO) & "'"
            CMySql.Executa CmdSql
         End If
      Next K
   End If
End Function

Public Sub Main()
   CMySql.Inicializa "Tabum", "", "localhost", ""
   CriaTab_MySql Cnn.NomeServidor, App.EXEName
   TabelasFixas
   
   frmLogin.Show
End Sub

Public Sub TabelasFixas()

'  TABELA REFERENTE AO CADASTRO DE OPÇÕES -  FORM FrmCadastrodeOpcoes
   CmdSql = "CREATE TABLE IF NOT EXISTS TIPOCOMBO(" & vbCr
   CmdSql = CmdSql & "        PROJETO         VARCHAR(20)     NOT NULL," & vbCr
   CmdSql = CmdSql & "        FORMULARIO      VARCHAR(60)     NOT NULL," & vbCr
   CmdSql = CmdSql & "        TIPO            VARCHAR(30)     NOT NULL," & vbCr
   CmdSql = CmdSql & "        DESCRICAO       VARCHAR(40)     NOT NULL," & vbCr
   CmdSql = CmdSql & "        PRIMARY KEY(PROJETO,FORMULARIO,TIPO,DESCRICAO)" & vbCr
   CmdSql = CmdSql & ")TYPE=MYISAM"
   CMySql.Executa CmdSql, True
   
'  TABELA CLIENTES
   CmdSql = "CREATE TABLE IF NOT EXISTS CLIENTE(" & vbCr
   CmdSql = CmdSql & "        PROJETO         VARCHAR(20)     NOT NULL," & vbCr
   CmdSql = CmdSql & "        TIPO            CHAR(04)        NOT NULL," & vbCr
   CmdSql = CmdSql & "        CODIGO          VARCHAR(14)     NOT NULL," & vbCr
   CmdSql = CmdSql & "        NOME            VARCHAR(60)     NOT NULL," & vbCr
   CmdSql = CmdSql & "        ENDERECO        VARCHAR(100)    NOT NULL," & vbCr
   CmdSql = CmdSql & "        BAIRRO          VARCHAR(40)     NOT NULL," & vbCr
   CmdSql = CmdSql & "        CIDADE          VARCHAR(60)     NOT NULL," & vbCr
   CmdSql = CmdSql & "        CONTATO         VARCHAR(60)     DEFAULT ' '," & vbCr
   CmdSql = CmdSql & "        TELEFONE        VARCHAR(20)     DEFAULT ' '," & vbCr
   CmdSql = CmdSql & "        OBS             VARCHAR(100)    DEFAULT ' '," & vbCr
   CmdSql = CmdSql & "        EMAIL           VARCHAR(60)     DEFAULT ' '," & vbCr
   CmdSql = CmdSql & "        PRIMARY KEY(PROJETO,CODIGO)" & vbCr
   CmdSql = CmdSql & ")TYPE=MYISAM"
   CMySql.Executa CmdSql, True

'  TABELA CLIENTES
   CmdSql = "CREATE TABLE IF NOT EXISTS ATENDIMENTO(" & vbCr
   CmdSql = CmdSql & "        PROJETO         VARCHAR(20)     NOT NULL," & vbCr
   CmdSql = CmdSql & "        TICKET          VARCHAR(15)     NOT NULL," & vbCr
   CmdSql = CmdSql & "        PEDIDO          VARCHAR(20)     NOT NULL," & vbCr
   CmdSql = CmdSql & "        CLIENTE         VARCHAR(60)     NOT NULL," & vbCr
   CmdSql = CmdSql & "        STATUS          VARCHAR(50)     NOT NULL," & vbCr
   CmdSql = CmdSql & "        DESCRICAO       VARCHAR(255)    NOT NULL," & vbCr
   CmdSql = CmdSql & "        ABERTURA        DATE            NOT NULL," & vbCr
   CmdSql = CmdSql & "        ENCERRAMENTO    DATE            DEFAULT '01/01/1900'," & vbCr
   CmdSql = CmdSql & "        PRIMARY KEY(PROJETO,TICKET)" & vbCr
   CmdSql = CmdSql & ")TYPE=MYISAM"
   CMySql.Executa CmdSql, True

'  TABELA CUSTO PRODUTO
   CmdSql = "CREATE TABLE IF NOT EXISTS CUSTO_PRODUTO(" & vbCr
   CmdSql = CmdSql & "        PROJETO                 VARCHAR(20)    NOT NULL," & vbCr
   CmdSql = CmdSql & "        PRODUTO                 VARCHAR(200)   NOT NULL," & vbCr
   CmdSql = CmdSql & "        PRECO_PRATICADO         DECIMAL(12,2)  NOT NULL," & vbCr
   CmdSql = CmdSql & "        PRECO_FORNECEDOR        DECIMAL(12,2)  NOT NULL," & vbCr
   CmdSql = CmdSql & "        FRETE_FORNECEDOR        DECIMAL(12,2)  NOT NULL," & vbCr
   CmdSql = CmdSql & "        GATEWAY                 DECIMAL(12,2)  NOT NULL," & vbCr
   CmdSql = CmdSql & "        IOF                     DECIMAL(12,2)  NOT NULL," & vbCr
   CmdSql = CmdSql & "        PARCELAMENTO            DECIMAL(12,2)  NOT NULL," & vbCr
   CmdSql = CmdSql & "        OUTROS_CUSTOS           DECIMAL(12,2)  NOT NULL," & vbCr
   CmdSql = CmdSql & "        IMPOSTO                 DECIMAL(12,2)  NOT NULL," & vbCr
   CmdSql = CmdSql & "        MARKETING               DECIMAL(12,2)  NOT NULL," & vbCr
   CmdSql = CmdSql & "        PRIMARY KEY(PROJETO,PRODUTO)" & vbCr
   CmdSql = CmdSql & ")TYPE=MYISAM"
   CMySql.Executa CmdSql, True

End Sub

