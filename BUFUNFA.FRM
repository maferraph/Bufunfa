VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Tela_Bufunfa 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bufunfa"
   ClientHeight    =   4092
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6132
   Icon            =   "Bufunfa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Bufunfa"
   MaxButton       =   0   'False
   ScaleHeight     =   4092
   ScaleWidth      =   6132
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox DIR_1 
      Height          =   288
      Left            =   5760
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CommandButton BT_Sair 
      Caption         =   "Sai&r"
      Height          =   852
      Left            =   5160
      Picture         =   "Bufunfa.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Finaliza o programa Bufunfa"
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton BT_Imprimir 
      Caption         =   "&Imprimir"
      Height          =   852
      Left            =   4320
      Picture         =   "Bufunfa.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimi os dados."
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton BT_Cancelar 
      Caption         =   "&Cancelar"
      Height          =   852
      Left            =   3480
      Picture         =   "Bufunfa.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancela operação."
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton BT_Salvar 
      Caption         =   "&Salvar"
      Height          =   852
      Left            =   2640
      Picture         =   "Bufunfa.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salva os dados."
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton BT_Apagar 
      Caption         =   "&Apagar"
      Height          =   852
      Left            =   1800
      Picture         =   "Bufunfa.frx":106A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Apaga dados."
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton BT_Consultar 
      Caption         =   "C&onsultar"
      Height          =   852
      Left            =   960
      Picture         =   "Bufunfa.frx":14AC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Consulta dados."
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton BT_Novo 
      Caption         =   "&Novo"
      Height          =   852
      Left            =   120
      Picture         =   "Bufunfa.frx":17B6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Novo dado"
      Top             =   120
      Width           =   852
   End
   Begin TabDlg.SSTab ST_1 
      Height          =   2892
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   5892
      _ExtentX        =   10393
      _ExtentY        =   5101
      _Version        =   327680
      Tabs            =   4
      TabsPerRow      =   2
      TabHeight       =   420
      TabCaption(0)   =   "Movimentos"
      TabPicture(0)   =   "Bufunfa.frx":1AC0
      Tab(0).ControlCount=   2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FR_Movimentos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FG_1"
      Tab(0).Control(1).Enabled=   0   'False
      TabCaption(1)   =   "Registros Cheques"
      TabPicture(1)   =   "Bufunfa.frx":1ADC
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FR_Registros"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "Cadastros"
      TabPicture(2)   =   "Bufunfa.frx":1AF8
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FR_Cadastros"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "Configurações"
      TabPicture(3)   =   "Bufunfa.frx":1B14
      Tab(3).ControlCount=   0
      Tab(3).ControlEnabled=   0   'False
      Begin MSFlexGridLib.MSFlexGrid FG_1 
         Height          =   2052
         Left            =   120
         TabIndex        =   47
         Top             =   2640
         Width           =   5532
         _ExtentX        =   9758
         _ExtentY        =   3620
         _Version        =   327680
         Rows            =   1
         AllowBigSelection=   0   'False
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin VB.Frame FR_Registros 
         Height          =   2172
         Left            =   -74880
         TabIndex        =   42
         Top             =   600
         Width           =   5532
         Begin VB.TextBox TXT_NumFinal 
            Height          =   288
            Left            =   3840
            TabIndex        =   21
            ToolTipText     =   "Número do último cheque"
            Top             =   1560
            Width           =   1212
         End
         Begin VB.ComboBox CB_ContaCheque 
            Height          =   288
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   18
            ToolTipText     =   "Conta do cheque à ser registrado"
            Top             =   480
            Width           =   1212
         End
         Begin VB.TextBox TXT_NumInicio 
            Height          =   288
            Left            =   1920
            TabIndex        =   20
            ToolTipText     =   "Número do primeiro cheque"
            Top             =   1560
            Width           =   1212
         End
         Begin VB.ListBox LT_Cheques 
            Height          =   1584
            ItemData        =   "Bufunfa.frx":1B30
            Left            =   120
            List            =   "Bufunfa.frx":1B32
            Sorted          =   -1  'True
            TabIndex        =   17
            ToolTipText     =   "Lista dos cheques registrados"
            Top             =   240
            Width           =   1332
         End
         Begin MSMask.MaskEdBox TXT_DataCheque 
            Height          =   288
            Left            =   3840
            TabIndex        =   19
            ToolTipText     =   "Data do registro"
            Top             =   480
            Width           =   1212
            _ExtentX        =   2138
            _ExtentY        =   508
            _Version        =   327680
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   8
            Format          =   "ddddd"
            Mask            =   "99/99/99"
            PromptChar      =   "_"
         End
         Begin VB.Label LB_DataCheque 
            AutoSize        =   -1  'True
            Caption         =   "Data Registro:"
            Height          =   192
            Left            =   3840
            TabIndex        =   46
            Top             =   240
            Width           =   1032
         End
         Begin VB.Label LB_NumFinal 
            AutoSize        =   -1  'True
            Caption         =   "Nº  Final:"
            Height          =   192
            Left            =   3840
            TabIndex        =   45
            Top             =   1320
            Width           =   636
         End
         Begin VB.Label LB_ContaCheque 
            AutoSize        =   -1  'True
            Caption         =   "Con&ta:"
            Height          =   192
            Left            =   1920
            TabIndex        =   44
            Top             =   240
            Width           =   456
         End
         Begin VB.Label LB_NumInicio 
            AutoSize        =   -1  'True
            Caption         =   "Nº  Inicial:"
            Height          =   192
            Left            =   1920
            TabIndex        =   43
            Top             =   1320
            Width           =   696
         End
      End
      Begin VB.Frame FR_Cadastros 
         Height          =   2172
         Left            =   -74880
         TabIndex        =   35
         Top             =   600
         Width           =   5532
         Begin VB.TextBox TXT_NumConta 
            Height          =   288
            Left            =   4080
            MaxLength       =   15
            TabIndex        =   27
            ToolTipText     =   "Número da conta corrente da conta registrada"
            Top             =   1680
            Width           =   1212
         End
         Begin VB.TextBox TXT_NumAgencia 
            Height          =   288
            Left            =   2760
            MaxLength       =   15
            TabIndex        =   26
            ToolTipText     =   "Número da agência da conta registrada"
            Top             =   1680
            Width           =   1212
         End
         Begin VB.TextBox TXT_NomeAgencia 
            Height          =   288
            Left            =   2760
            MaxLength       =   20
            TabIndex        =   25
            ToolTipText     =   "Nome da agência da conta registrada"
            Top             =   1080
            Width           =   2532
         End
         Begin VB.TextBox TXT_Banco 
            Height          =   288
            Left            =   2760
            MaxLength       =   25
            TabIndex        =   24
            ToolTipText     =   "Nome do banco da conta registrada"
            Top             =   480
            Width           =   2532
         End
         Begin VB.TextBox TXT_NomeConta 
            Height          =   288
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   23
            ToolTipText     =   "Nome da conta cadastrada"
            Top             =   1080
            Width           =   1212
         End
         Begin VB.ListBox LT_Cadastros 
            Height          =   1776
            ItemData        =   "Bufunfa.frx":1B34
            Left            =   120
            List            =   "Bufunfa.frx":1B36
            TabIndex        =   22
            ToolTipText     =   "Lista de contas cadastradas"
            Top             =   240
            Width           =   972
         End
         Begin VB.Label LB_NumConta 
            AutoSize        =   -1  'True
            Caption         =   "Nº  C/C:"
            Height          =   192
            Left            =   4080
            TabIndex        =   40
            Top             =   1440
            Width           =   576
         End
         Begin VB.Label LB_NumAgencia 
            AutoSize        =   -1  'True
            Caption         =   "Nº  Agência:"
            Height          =   192
            Left            =   2760
            TabIndex        =   39
            Top             =   1440
            Width           =   888
         End
         Begin VB.Label LB_NomeAgencia 
            AutoSize        =   -1  'True
            Caption         =   "Nome da Agência:"
            Height          =   192
            Left            =   2760
            TabIndex        =   38
            Top             =   840
            Width           =   1344
         End
         Begin VB.Label LB_Banco 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   192
            Left            =   2760
            TabIndex        =   37
            Top             =   240
            Width           =   504
         End
         Begin VB.Label LB_NomeConta 
            AutoSize        =   -1  'True
            Caption         =   "Conta:"
            Height          =   192
            Left            =   1320
            TabIndex        =   36
            Top             =   840
            Width           =   456
         End
      End
      Begin VB.Frame FR_Movimentos 
         Height          =   2172
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   5532
         Begin VB.OptionButton RD_Lancamento 
            Caption         =   "Lançamentos"
            Height          =   192
            Left            =   360
            TabIndex        =   10
            ToolTipText     =   "Movimento do tipo entrada"
            Top             =   840
            Width           =   1212
         End
         Begin VB.OptionButton RD_Entrada 
            Caption         =   "Entrada"
            Height          =   192
            Left            =   360
            TabIndex        =   9
            ToolTipText     =   "Movimento do tipo entrada"
            Top             =   600
            Width           =   852
         End
         Begin VB.OptionButton RD_Saida 
            Caption         =   "Saída"
            Height          =   252
            Left            =   360
            TabIndex        =   8
            ToolTipText     =   "Movimento do tipo saída"
            Top             =   360
            Width           =   732
         End
         Begin MSMask.MaskEdBox TXT_DataEmissao 
            Height          =   288
            Left            =   2280
            TabIndex        =   13
            ToolTipText     =   "Data da emissao do documento"
            Top             =   1080
            Width           =   1212
            _ExtentX        =   2138
            _ExtentY        =   508
            _Version        =   327680
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   8
            Format          =   "ddddd"
            Mask            =   "99/99/99"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TXT_Documento 
            Height          =   288
            Left            =   2280
            TabIndex        =   15
            ToolTipText     =   "Número/Nome do documento"
            Top             =   1680
            Width           =   1212
         End
         Begin VB.ComboBox CB_Conta 
            Height          =   288
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   12
            ToolTipText     =   "Conta do registro do movimento"
            Top             =   1680
            Width           =   1332
         End
         Begin VB.ComboBox CB_Tipo 
            Height          =   288
            ItemData        =   "Bufunfa.frx":1B38
            Left            =   2280
            List            =   "Bufunfa.frx":1B3A
            TabIndex        =   11
            ToolTipText     =   "Histórico do movimento"
            Top             =   480
            Width           =   2892
         End
         Begin MSMask.MaskEdBox TXT_DataPre 
            Height          =   288
            Left            =   3960
            TabIndex        =   14
            ToolTipText     =   "Data pré-datada do documento"
            Top             =   1080
            Width           =   1212
            _ExtentX        =   2138
            _ExtentY        =   508
            _Version        =   327680
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   8
            Format          =   "ddddd"
            Mask            =   "99/99/99"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TXT_Valor 
            Height          =   288
            Left            =   3960
            TabIndex        =   16
            ToolTipText     =   "Valor do documento"
            Top             =   1680
            Width           =   1212
            _ExtentX        =   2138
            _ExtentY        =   508
            _Version        =   327680
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label LB_Documento 
            AutoSize        =   -1  'True
            Caption         =   "Docu&mento:"
            Height          =   192
            Left            =   2280
            TabIndex        =   34
            Top             =   1440
            Width           =   864
         End
         Begin VB.Label LB_Conta 
            AutoSize        =   -1  'True
            Caption         =   "Con&ta:"
            Height          =   192
            Left            =   360
            TabIndex        =   33
            Top             =   1440
            Width           =   456
         End
         Begin VB.Label LB_Valor 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Height          =   192
            Left            =   3960
            TabIndex        =   32
            Top             =   1440
            Width           =   420
         End
         Begin VB.Label LB_DataPre 
            AutoSize        =   -1  'True
            Caption         =   "Data Pré:"
            Height          =   192
            Left            =   3960
            TabIndex        =   31
            Top             =   840
            Width           =   672
         End
         Begin VB.Label LB_DataEmissao 
            AutoSize        =   -1  'True
            Caption         =   "Data Emissão:"
            Height          =   192
            Left            =   2280
            TabIndex        =   30
            Top             =   840
            Width           =   1056
         End
         Begin VB.Label LB_Tipo 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   192
            Left            =   2280
            TabIndex        =   29
            Top             =   240
            Width           =   372
         End
      End
   End
End
Attribute VB_Name = "Tela_Bufunfa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BT_Apagar_Click()
    If ST_1.Tab = 0 Then 'Tab de Movimento
        LimpaCampos (0)
        RD_Saida.SetFocus
    ElseIf ST_1.Tab = 1 Then 'Tab de Registros de Cheques
        LimpaCampos (1)
        CB_ContaCheque.SetFocus
    ElseIf ST_1.Tab = 2 Then 'Tab de Cadastros
        LimpaCampos (2)
        TXT_NomeConta.SetFocus
    End If
End Sub
Private Sub BT_Cancelar_Click()
    FR_Movimentos.Visible = True
    FG_1.Visible = False
    BotoesTelaEmEdicao (False)
    If ST_1.Tab = 0 Then 'Tab de Movimento
        AtivaTabs (-1)
        AtivaTelaMovimentos (False)
        LimpaCampos (0)
        BT_Novo.SetFocus
    ElseIf ST_1.Tab = 1 Then 'Tab de Registros de Cheques
        AtivaTabs (-1)
        AtivaTelaRegistros (False)
        LimpaCampos (1)
        BT_Novo.SetFocus
    ElseIf ST_1.Tab = 2 Then 'Tab de Cadastros
        AtivaTabs (-1)
        AtivaTelaCadastros (False)
        LimpaCampos (2)
        BT_Novo.SetFocus
    End If
    Estado = "N"
End Sub
Private Sub BT_Consultar_Click()
    Estado = "C"
    If ST_1.Tab = 0 Then 'Tab de Movimento
        Tela_Movimento_Dlg.Show vbModal
    ElseIf ST_1.Tab = 1 Then 'Tab de Registros de Cheques
    
    ElseIf ST_1.Tab = 2 Then 'Tab de Cadastros
    
    End If
End Sub
Private Sub BT_Novo_Click()
    Estado = "E"
    BotoesTelaEmEdicao (True)
    If ST_1.Tab = 0 Then 'Tab de Movimento
        If LT_Cadastros.ListCount = 0 Then
            Beep
            MsgBox ("Ainda não existe nenhuma conta cadastrada... inclua uma conta antes de usar o programa.")
            ST_1.Tab = 2
        End If
        AtivaTabs (0)
        AtivaTelaMovimentos (True)
        RD_Saida.SetFocus
    ElseIf ST_1.Tab = 1 Then 'Tab de Registros de Cheques
        AtivaTabs (1)
        AtivaTelaRegistros (True)
        TXT_DataCheque.Text = Date
        CB_ContaCheque.SetFocus
    ElseIf ST_1.Tab = 2 Then 'Tab de Cadastros
        AtivaTabs (2)
        AtivaTelaCadastros (True)
        TXT_NomeConta.SetFocus
    End If
End Sub
Private Sub BT_Sair_Click()
    End
End Sub
Private Sub BT_Salvar_Click()
    Dim Numero As Integer
    Dim Valor As Currency
    
    If ST_1.Tab = 0 Then 'Tab de Movimento
        If VerificaDados(0) = True Then
            BDBUF_TBMOV.AddNew
            BDBUF_TBMOV_CPHIS.Value = CB_Tipo.Text
            BDBUF_TBMOV_CPIMP.Value = True
            BDBUF_TBMOV_CPDEM.Value = TXT_DataEmissao.Text
            BDBUF_TBMOV_CPDOC.Value = TXT_Documento.Text
            BDBUF_TBMOV_CPVAL.Value = TXT_Valor.Text
            If RD_Saida.Value = True Then
                BDBUF_TBMOV_CPTIP.Value = "S"
                If IsDate(BDBUF_TBMOV_CPDPR.Value) Then
                    BDBUF_TBMOV_CPDPR.Value = TXT_DataPre.Text
                End If
                BDBUF_TBMOV_CPCON.Value = CB_Conta.List(CB_Conta.ListIndex)
            ElseIf RD_Entrada.Value = True Then
                BDBUF_TBMOV_CPTIP.Value = "E"
            ElseIf RD_Lancamento.Value = True Then
                BDBUF_TBMOV_CPTIP.Value = "L"
                BDBUF_TBMOV_CPCON.Value = CB_Conta.List(CB_Conta.ListIndex)
            End If
            BDBUF_TBMOV.Update
            If RD_Saida.Value = True Then
                If BDBUF_TBMOV.RecordCount <> 0 Then
                    BDBUF_TBMOV.MoveFirst
                    While Not BDBUF_TBMOV.EOF
                        If CB_Tipo.ListCount = 0 And BDBUF_TBMOV_CPTIP.Value = "S" Then
                            CB_Tipo.AddItem (BDBUF_TBMOV_CPHIS.Value)
                        End If
                        For I = 0 To CB_Tipo.ListCount - 1
                            If Trim(CB_Tipo.List(I)) <> Trim(BDBUF_TBMOV_CPHIS.Value) And _
                                BDBUF_TBMOV_CPTIP.Value = "S" Then
                                CB_Tipo.AddItem (BDBUF_TBMOV_CPHIS.Value)
                                Exit For
                            End If
                        Next I
                        BDBUF_TBMOV.MoveNext
                    Wend
                End If
            ElseIf RD_Entrada.Value = True Then
                If BDBUF_TBMOV.RecordCount <> 0 Then
                    BDBUF_TBMOV.MoveFirst
                    While Not BDBUF_TBMOV.EOF
                        If CB_Tipo.ListCount = 0 And BDBUF_TBMOV_CPTIP.Value = "E" Then
                            CB_Tipo.AddItem (BDBUF_TBMOV_CPHIS.Value)
                        End If
                        For I = 0 To CB_Tipo.ListCount - 1
                            If Trim(CB_Tipo.List(I)) <> Trim(BDBUF_TBMOV_CPHIS.Value) And _
                                BDBUF_TBMOV_CPTIP.Value = "E" Then
                                CB_Tipo.AddItem (BDBUF_TBMOV_CPHIS.Value)
                                Exit For
                            End If
                        Next I
                        BDBUF_TBMOV.MoveNext
                    Wend
                End If
            End If
            LimpaCampos (0)
            CB_Tipo.SetFocus
        End If
    ElseIf ST_1.Tab = 1 Then 'Tab de Registros de Cheques
        If VerificaDados(1) = True Then
            BDBUF_TBREG.MoveLast
            Numero = BDBUF_TBREG_CPNRG.Value + 1
                        
            BDBUF_TBREG.AddNew
            BDBUF_TBREG_CPNRG = Numero
            BDBUF_TBREG_CPCON.Value = CB_ContaCheque.List(CB_ContaCheque.ListIndex)
            BDBUF_TBREG_CPDAT.Value = TXT_DataCheque.Text
            BDBUF_TBREG_CPNIN.Value = TXT_NumInicio.Text
            BDBUF_TBREG_CPNFN.Value = TXT_NumFinal.Text
            BDBUF_TBREG.Update
            LT_Cheques.AddItem (Numero)
            LimpaCampos (1)
            AtivaTabs (-1)
            AtivaTelaRegistros (False)
            BT_Cancelar.Value = True
        End If
    ElseIf ST_1.Tab = 2 Then 'Tab de Cadastros
        If VerificaDados(2) = True Then
            BDBUF_TBCAD.Seek "=", Trim(TXT_NomeConta.Text)
            If BDBUF_TBCAD.NoMatch Then
                BDBUF_TBCAD.AddNew
                BDBUF_TBCAD_CPCON = TXT_NomeConta.Text
                BDBUF_TBCAD_CPNMB = TXT_Banco.Text
                BDBUF_TBCAD_CPNMA = TXT_NomeAgencia.Text
                BDBUF_TBCAD_CPNUA = TXT_NumAgencia.Text
                BDBUF_TBCAD_CPNUC = TXT_NumConta.Text
                BDBUF_TBCAD_CPATU = False
                BDBUF_TBCAD.Update
            Else
                Beep
                MsgBox ("Já existe uma conta com esse nome ou ocorreu um erro durante a gravação dos dados... verifique os dados digitados e tente mais tarde.")
                TXT_NomeConta.SetFocus
                Exit Sub
            End If
            CB_Conta.AddItem (TXT_NomeConta.Text)
            LT_Cadastros.AddItem (TXT_NomeConta.Text)
            LimpaCampos (2)
            AtivaTabs (-1)
            AtivaTelaCadastros (False)
            BT_Cancelar.Value = True
        End If
    End If
End Sub
Private Sub CB_Conta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TXT_DataEmissao.SetFocus
    End If
End Sub
Private Sub CB_ContaCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TXT_DataCheque.SetFocus
    End If
End Sub
Private Sub CB_Tipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And RD_Saida.Value = True Then
        CB_Conta.SetFocus
    ElseIf KeyAscii = 13 And RD_Entrada.Value = True Then
        TXT_DataEmissao.SetFocus
    ElseIf KeyAscii = 13 And RD_Lancamento.Value = True Then
        CB_Conta.SetFocus
    End If
End Sub
Private Sub Form_Load()
    BotoesTelaEmEdicao (False)
    DIR_1.Path = "C:\Meus Documentos\Programas\Bufunfa\Bufunfa"
        
    'Abrindo Workspaces
    Set WKSP = DBEngine.Workspaces(0)
    
    'Abrindo Banco de Dados
    Set BDBUF = WKSP.OpenDatabase(DIR_1.Path & "\Bufunfa.mdb")
    
    ' Abre Tabelas
    Set BDBUF_TBMOV = BDBUF.OpenRecordset("Movimentos")
    Set BDBUF_TBCAD = BDBUF.OpenRecordset("Cadastros")
    Set BDBUF_TBREG = BDBUF.OpenRecordset("Registros")
                
    ' Abre Campos da Tabela Movimentos
    Set BDBUF_TBMOV_CPHIS = BDBUF_TBMOV.Fields("Histórico")
    Set BDBUF_TBMOV_CPTIP = BDBUF_TBMOV.Fields("Tipo")
    Set BDBUF_TBMOV_CPIMP = BDBUF_TBMOV.Fields("Imprimir")
    Set BDBUF_TBMOV_CPCON = BDBUF_TBMOV.Fields("Conta")
    Set BDBUF_TBMOV_CPDEM = BDBUF_TBMOV.Fields("Data Emissão")
    Set BDBUF_TBMOV_CPDPR = BDBUF_TBMOV.Fields("Data Pré")
    Set BDBUF_TBMOV_CPDOC = BDBUF_TBMOV.Fields("Documento")
    Set BDBUF_TBMOV_CPVAL = BDBUF_TBMOV.Fields("Valor")
    
    ' Abre Campos da Tabela Registros Cheques
    Set BDBUF_TBREG_CPNRG = BDBUF_TBREG.Fields("Número do Registro")
    Set BDBUF_TBREG_CPCON = BDBUF_TBREG.Fields("Conta")
    Set BDBUF_TBREG_CPDAT = BDBUF_TBREG.Fields("Data do Registro")
    Set BDBUF_TBREG_CPNIN = BDBUF_TBREG.Fields("Número Inicial")
    Set BDBUF_TBREG_CPNFN = BDBUF_TBREG.Fields("Número Final")

    ' Abre Campos da Tabela Cadastros
    Set BDBUF_TBCAD_CPCON = BDBUF_TBCAD.Fields("Conta")
    Set BDBUF_TBCAD_CPNMB = BDBUF_TBCAD.Fields("Nome do Banco")
    Set BDBUF_TBCAD_CPNMA = BDBUF_TBCAD.Fields("Nome da Agência")
    Set BDBUF_TBCAD_CPNUA = BDBUF_TBCAD.Fields("Número da Agência")
    Set BDBUF_TBCAD_CPNUC = BDBUF_TBCAD.Fields("Número da Conta Corrente")
    Set BDBUF_TBCAD_CPATU = BDBUF_TBCAD.Fields("Atual")

    BDBUF_TBCAD.Index = "Conta"
    'BDBUF_TBMOV.Index = "ProcMov"
    BDBUF_TBREG.Index = "Número do Registro"
    
    'Tab de Movimento
    FG_1.Visible = False
    AtivaTelaMovimentos (False)
    If BDBUF_TBMOV.RecordCount <> 0 Then
        BDBUF_TBMOV.MoveFirst
        While Not BDBUF_TBMOV.EOF
            If CB_Tipo.ListCount = 0 Then
                CB_Tipo.AddItem (BDBUF_TBMOV_CPHIS.Value)
            End If
            For I = 0 To CB_Tipo.ListCount - 1
                If Trim(CB_Tipo.List(I)) <> Trim(BDBUF_TBMOV_CPHIS.Value) And _
                    BDBUF_TBMOV_CPTIP.Value = "S" Then
                    CB_Tipo.AddItem (BDBUF_TBMOV_CPHIS.Value)
                    Exit For
                End If
            Next I
            BDBUF_TBMOV.MoveNext
        Wend
    End If
    Dim NomeBan As String
    NomeBan = ""
    If BDBUF_TBCAD.RecordCount <> 0 Then
        BDBUF_TBCAD.MoveFirst
        While Not BDBUF_TBCAD.EOF
            CB_Conta.AddItem (BDBUF_TBCAD_CPCON.Value)
            CB_ContaCheque.AddItem (BDBUF_TBCAD_CPCON.Value)
            If BDBUF_TBCAD_CPATU.Value = True Then
                NomeBan = BDBUF_TBCAD_CPCON.Value
            End If
            BDBUF_TBCAD.MoveNext
        Wend
    End If
    If NomeBan <> "" Then
        For I = 0 To CB_Conta.ListCount
            If CB_Conta.List(I) = NomeBan Then
                CB_Conta.ListIndex = I
                CB_ContaCheque.ListIndex = I
                Exit For
            End If
        Next I
    End If
    
    'Tab de Registro de Cheques
    AtivaTelaRegistros (False)
    LT_Cheques.Enabled = False
    If BDBUF_TBREG.RecordCount <> 0 Then
        BDBUF_TBREG.MoveFirst
        While Not BDBUF_TBREG.EOF
            LT_Cheques.AddItem (BDBUF_TBREG_CPNRG.Value)
            BDBUF_TBREG.MoveNext
        Wend
    End If
    
    
    'Tab de Cadastros
    AtivaTelaCadastros (False)
    LT_Cadastros.Enabled = False
    If BDBUF_TBCAD.RecordCount <> 0 Then
        BDBUF_TBCAD.MoveFirst
        While Not BDBUF_TBCAD.EOF
            LT_Cadastros.AddItem (BDBUF_TBCAD_CPCON.Value)
            BDBUF_TBCAD.MoveNext
        Wend
    End If
    If LT_Cadastros.ListCount = 0 Then
        Beep
        MsgBox ("Ainda não existe nenhuma conta cadastrada... inclua uma conta antes de usar o programa.")
    End If
    
    Estado = "N"
    ST_1.Tab = 0
End Sub
Private Sub Menu_Sair_Click()
    Unload Tela_Bufunfa
End Sub
Private Sub Form_Terminate()
    Unload Tela_Bufunfa
End Sub
Private Sub Form_Unload(Cancel As Integer)
    BDBUF_TBMOV.Close
    BDBUF_TBCAD.Close
    BDBUF.Close
    End
End Sub
Private Sub RD_Entrada_Click()
    If RD_Entrada.Value = True Then
        LB_DataPre.Visible = False
        TXT_DataPre.Visible = False
        LB_Conta.Visible = False
        CB_Conta.Visible = False
        CB_Tipo.Clear
        If BDBUF_TBMOV.RecordCount <> 0 Then
            BDBUF_TBMOV.MoveFirst
            While Not BDBUF_TBMOV.EOF
                If CB_Tipo.ListCount = 0 And BDBUF_TBMOV_CPTIP.Value = "E" Then
                    CB_Tipo.AddItem (BDBUF_TBMOV_CPHIS.Value)
                End If
                For I = 0 To CB_Tipo.ListCount - 1
                    If Trim(CB_Tipo.List(I)) <> Trim(BDBUF_TBMOV_CPHIS.Value) And _
                        BDBUF_TBMOV_CPTIP.Value = "E" Then
                        CB_Tipo.AddItem (BDBUF_TBMOV_CPHIS.Value)
                        Exit For
                    End If
                Next I
                BDBUF_TBMOV.MoveNext
            Wend
        End If
    End If
End Sub
Private Sub RD_Entrada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CB_Tipo.SetFocus
    End If
End Sub
Private Sub RD_Lancamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CB_Tipo.SetFocus
    End If
End Sub
Private Sub RD_Saida_Click()
    If RD_Saida.Value = True Then
        LB_DataPre.Visible = True
        TXT_DataPre.Visible = True
        LB_Conta.Visible = True
        CB_Conta.Visible = True
        CB_Tipo.Clear
        If BDBUF_TBMOV.RecordCount <> 0 Then
            BDBUF_TBMOV.MoveFirst
            While Not BDBUF_TBMOV.EOF
                If CB_Tipo.ListCount = 0 And BDBUF_TBMOV_CPTIP.Value = "S" Then
                    CB_Tipo.AddItem (BDBUF_TBMOV_CPHIS.Value)
                End If
                For I = 0 To CB_Tipo.ListCount - 1
                    If Trim(CB_Tipo.List(I)) <> Trim(BDBUF_TBMOV_CPHIS.Value) And _
                        BDBUF_TBMOV_CPTIP.Value = "S" Then
                        CB_Tipo.AddItem (BDBUF_TBMOV_CPHIS.Value)
                        Exit For
                    End If
                Next I
                BDBUF_TBMOV.MoveNext
            Wend
        End If
    End If
End Sub
Private Sub RD_Saida_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CB_Tipo.SetFocus
    End If
End Sub
Private Sub RD_Lancamento_Click()
    If RD_Lancamento.Value = True Then
        LB_DataPre.Visible = True
        TXT_DataPre.Visible = True
        LB_Conta.Visible = True
        CB_Conta.Visible = True
        CB_Tipo.Clear
        If BDBUF_TBMOV.RecordCount <> 0 Then
            BDBUF_TBMOV.MoveFirst
            While Not BDBUF_TBMOV.EOF
                If CB_Tipo.ListCount = 0 And BDBUF_TBMOV_CPTIP.Value = "L" Then
                    CB_Tipo.AddItem (BDBUF_TBMOV_CPHIS.Value)
                End If
                For I = 0 To CB_Tipo.ListCount - 1
                    If Trim(CB_Tipo.List(I)) <> Trim(BDBUF_TBMOV_CPHIS.Value) And _
                        BDBUF_TBMOV_CPTIP.Value = "L" Then
                        CB_Tipo.AddItem (BDBUF_TBMOV_CPHIS.Value)
                        Exit For
                    End If
                Next I
                BDBUF_TBMOV.MoveNext
            Wend
        End If
    End If
End Sub
Private Sub TXT_Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TXT_NomeAgencia.SetFocus
    End If
End Sub
Private Sub TXT_DataCheque_GotFocus()
    TXT_DataCheque.SelLength = Len(TXT_DataCheque.Text)
End Sub
Private Sub TXT_DataCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TXT_NumInicio.SetFocus
    End If
End Sub
Private Sub TXT_DataEmissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And RD_Saida.Value = True Then
        TXT_DataPre.SetFocus
    End If
End Sub
Private Sub TXT_DataPre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TXT_Documento.SetFocus
    End If
End Sub
Private Sub TXT_Documento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TXT_Valor.SetFocus
    End If
End Sub
Private Sub TXT_NomeAgencia_Change()
    If KeyAscii = 13 Then
        TXT_NumAgencia.SetFocus
    End If
End Sub
Private Sub TXT_NomeConta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TXT_Banco.SetFocus
    End If
End Sub
Private Sub TXT_NumAgencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TXT_NumConta.SetFocus
    End If
End Sub
Private Sub TXT_NumConta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BT_Salvar.SetFocus
    End If
End Sub
Private Sub TXT_NumFinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BT_Salvar.SetFocus
    End If
End Sub
Private Sub TXT_NumInicio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TXT_NumFinal.SetFocus
    End If
End Sub
Private Sub TXT_Saldo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BT_Salvar.SetFocus
    End If
End Sub
Private Sub TXT_Valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BT_Salvar.SetFocus
    End If
End Sub

'*************************************************
'                  F U N Ç Õ E S
'*************************************************

Public Static Sub AtivaTelaMovimentos(Valor As Boolean)
    FR_Movimentos.Enabled = Valor
    RD_Saida.Enabled = Valor
    RD_Entrada.Enabled = Valor
    RD_Lancamento.Enabled = Valor
    CB_Tipo.Enabled = Valor
    CB_Conta.Enabled = Valor
    TXT_DataEmissao.Enabled = Valor
    TXT_DataPre.Enabled = Valor
    TXT_Documento.Enabled = Valor
    TXT_Valor.Enabled = Valor
    LB_Tipo.Enabled = Valor
    LB_Conta.Enabled = Valor
    LB_DataEmissao.Enabled = Valor
    LB_DataPre.Enabled = Valor
    LB_Documento.Enabled = Valor
    LB_Valor.Enabled = Valor
End Sub
Public Static Sub AtivaTelaRegistros(Valor As Boolean)
    FR_Registros.Enabled = Valor
    CB_ContaCheque.Enabled = Valor
    TXT_DataCheque.Enabled = Valor
    TXT_NumInicio.Enabled = Valor
    TXT_NumFinal.Enabled = Valor
    LB_ContaCheque.Enabled = Valor
    LB_DataCheque.Enabled = Valor
    LB_NumInicio.Enabled = Valor
    LB_NumFinal.Enabled = Valor
End Sub
Private Static Sub AtivaTelaCadastros(Valor As Boolean)
    FR_Cadastros.Enabled = Valor
    TXT_NomeConta.Enabled = Valor
    TXT_Banco.Enabled = Valor
    TXT_NomeAgencia.Enabled = Valor
    TXT_NumAgencia.Enabled = Valor
    TXT_NumConta.Enabled = Valor
    LB_NomeConta.Enabled = Valor
    LB_Banco.Enabled = Valor
    LB_NomeAgencia.Enabled = Valor
    LB_NumAgencia.Enabled = Valor
    LB_NumConta.Enabled = Valor
End Sub
Public Static Sub AtivaTabs(Valor As Integer)
    If Valor = -1 Then
        ST_1.TabEnabled(0) = True
        ST_1.TabEnabled(1) = True
        ST_1.TabEnabled(2) = True
        ST_1.TabEnabled(3) = True
    ElseIf Valor = 0 Then
        ST_1.TabEnabled(0) = True
        ST_1.TabEnabled(1) = False
        ST_1.TabEnabled(2) = False
        ST_1.TabEnabled(3) = False
    ElseIf Valor = 1 Then
        ST_1.TabEnabled(0) = False
        ST_1.TabEnabled(1) = True
        ST_1.TabEnabled(2) = False
        ST_1.TabEnabled(3) = False
    ElseIf Valor = 2 Then
        ST_1.TabEnabled(0) = False
        ST_1.TabEnabled(1) = False
        ST_1.TabEnabled(2) = True
        ST_1.TabEnabled(3) = False
    ElseIf Valor = 3 Then
        ST_1.TabEnabled(0) = False
        ST_1.TabEnabled(1) = False
        ST_1.TabEnabled(2) = False
        ST_1.TabEnabled(3) = True
    End If
End Sub
Public Static Sub BotoesTelaEmEdicao(Valor As Boolean)
    If Valor = True Then
        BT_Novo.Enabled = False
        BT_Consultar.Enabled = False
        BT_Apagar.Enabled = True
        BT_Salvar.Enabled = True
        BT_Cancelar.Enabled = True
        BT_Imprimir.Enabled = False
        BT_Sair.Enabled = False
    Else
        BT_Novo.Enabled = True
        BT_Consultar.Enabled = True
        BT_Apagar.Enabled = False
        BT_Salvar.Enabled = False
        BT_Cancelar.Enabled = False
        BT_Imprimir.Enabled = False
        BT_Sair.Enabled = True
    End If
End Sub
Private Static Sub LimpaCampos(Valor As Integer)
    If Valor = 0 Then  'Tab de Movimento
        CB_Tipo.Text = ""
        TXT_DataEmissao.Text = "__/__/__"
        TXT_DataPre.Text = "__/__/__"
        TXT_Documento.Text = ""
        TXT_Valor.Text = ""
    ElseIf Valor = 1 Then 'Tab de Registros de Cheques
        TXT_DataCheque.Text = "__/__/__"
        TXT_NumInicio.Text = ""
        TXT_NumFinal.Text = ""
    ElseIf Valor = 2 Then 'Tab de Cadastros
        TXT_NomeConta.Text = ""
        TXT_Banco.Text = ""
        TXT_NomeAgencia.Text = ""
        TXT_NumAgencia.Text = ""
        TXT_NumConta.Text = ""
    End If
End Sub
Private Static Function VerificaDados(Valor As Integer) As Boolean
    If Valor = 0 Then  'Tab de Movimento
        If RD_Saida.Value = False And RD_Entrada.Value = False And RD_Lancamento.Value = False Then
            MsgBox ("Campo tipo deve ser preenchido.")
            RD_Saida.SetFocus
            Exit Function
        ElseIf CB_Tipo.Text = "" Then
            MsgBox ("Campo do histórico deve ser preenchido.")
            CB_Tipo.SetFocus
            Exit Function
        ElseIf CB_Conta.Text = "" And RD_Saida.Value = True Then
            MsgBox ("Campo conta deve ser preenchida.")
            CB_Conta.SetFocus
            Exit Function
        ElseIf TXT_DataEmissao.Text = "" Then
            MsgBox ("Campo data emissão do documento deve ser preenchida.")
            TXT_DataEmissao.SetFocus
            Exit Function
        ElseIf TXT_Documento.Text = "" Then
            MsgBox ("Campo documento deve ser preenchido.")
            TXT_Documento.SetFocus
            Exit Function
        ElseIf TXT_Valor.Text = "" Then
            MsgBox ("Campo valor deve ser preenchido.")
            TXT_Valor.SetFocus
            Exit Function
        End If
    ElseIf Valor = 1 Then 'Tab de Registros de Cheques
        If CB_ContaCheque.Text = "" Then
            MsgBox ("Campo conta do registro deve ser preenchida.")
            CB_ContaCheque.SetFocus
            Exit Function
        ElseIf TXT_DataCheque.Text = "" Then
            MsgBox ("Campo data do registro deve ser preenchida.")
            TXT_DataCheque.SetFocus
            Exit Function
        ElseIf TXT_NumInicio.Text = "" Then
            MsgBox ("Campo número inicial do cheque deve ser preenchido.")
            TXT_NumInicio.SetFocus
            Exit Function
        ElseIf TXT_NumFinal.Text = "" Then
            MsgBox ("Campo número final do cheque deve ser preenchido.")
            TXT_NumFinal.SetFocus
            Exit Function
        End If
    ElseIf Valor = 2 Then 'Tab de Cadastros
        If TXT_NomeConta.Text = "" Then
            MsgBox ("Campo nome da conta deve ser preenchida.")
            TXT_NomeConta.SetFocus
            Exit Function
        ElseIf TXT_Banco.Text = "" Then
            MsgBox ("Campo nome do banco deve ser preenchido.")
            TXT_Banco.SetFocus
            Exit Function
        ElseIf TXT_NomeAgencia.Text = "" Then
            MsgBox ("Campo nome da agência deve ser preenchida.")
            TXT_NomeAgencia.SetFocus
            Exit Function
        ElseIf TXT_NumAgencia.Text = "" Then
            MsgBox ("Campo número da agência deve ser preenchida.")
            TXT_NumAgencia.SetFocus
            Exit Function
        ElseIf TXT_NumConta.Text = "" Then
            MsgBox ("Campo número da conta corrente deve ser preenchida.")
            TXT_NumConta.SetFocus
            Exit Function
        End If
    End If
    VerificaDados = True
End Function
Public Static Sub BotoesTelaEmConsulta(Valor As Boolean)
    If Valor = True Then
        BT_Novo.Enabled = False
        BT_Consultar.Enabled = False
        BT_Apagar.Enabled = False
        BT_Salvar.Enabled = False
        BT_Cancelar.Enabled = True
        BT_Imprimir.Enabled = True
        BT_Sair.Enabled = False
    Else
        BT_Novo.Enabled = True
        BT_Consultar.Enabled = True
        BT_Apagar.Enabled = False
        BT_Salvar.Enabled = False
        BT_Cancelar.Enabled = False
        BT_Imprimir.Enabled = False
        BT_Sair.Enabled = True
    End If
End Sub
Public Static Sub MontaFG()
    Dim Num, Rec As Integer
    Dim DataIni As Date
    Dim Saldo
    
    FG_1.Clear
    FG_1.Left = FR_Movimentos.Left
    FG_1.Top = FR_Movimentos.Top
    FG_1.FixedCols = 0
    FG_1.Cols = 5
    FG_1.TextMatrix(0, 0) = "Data"
    FG_1.ColWidth(0) = 1000
    FG_1.ColAlignment(0) = flexAlignCenterCenter
    FG_1.TextMatrix(0, 1) = "Histórico"
    FG_1.ColWidth(1) = 2000
    FG_1.ColAlignment(1) = flexAlignLeftCenter
    FG_1.TextMatrix(0, 2) = "Crédito"
    FG_1.ColWidth(2) = 750
    FG_1.ColAlignment(2) = flexAlignCenterCenter
    FG_1.TextMatrix(0, 3) = "Débito"
    FG_1.ColWidth(3) = 750
    FG_1.ColAlignment(3) = flexAlignCenterCenter
    FG_1.TextMatrix(0, 4) = "Saldo"
    FG_1.ColWidth(4) = 900
    FG_1.ColAlignment(4) = flexAlignCenterCenter
    FG_1.CellFontBold = True
    
    Num = 1
    Rec = 0
    ' Ínicio do controle em 23 / 10 / 97
    ' Saldo: - R$ 524,31
    Saldo = -576.21
    DataIni = #10/23/97#
    
    Do While Rec <= BDBUF_TBMOV.RecordCount
        BDBUF_TBMOV.MoveFirst
        Do While Not BDBUF_TBMOV.EOF
            If BDBUF_TBMOV_CPDEM.Value = DataIni Then
                If BDBUF_TBMOV_CPTIP.Value = "S" Or BDBUF_TBMOV_CPTIP.Value = "E" Then
                    FG_1.AddItem BDBUF_TBMOV_CPDEM.Value, Num
                    FG_1.TextMatrix(Num, 1) = BDBUF_TBMOV_CPHIS.Value
                    If BDBUF_TBMOV_CPTIP.Value = "E" Then 'Débito
                        FG_1.TextMatrix(Num, 2) = Val(BDBUF_TBMOV_CPVAL.Value)
                        If Num = 1 Then
                            FG_1.TextMatrix(Num, 4) = Val(Saldo) + Val(FG_1.TextMatrix(Num, 2))
                        Else
                            FG_1.TextMatrix(Num, 4) = Val(FG_1.TextMatrix(Num - 1, 4)) + Val(FG_1.TextMatrix(Num, 2))
                        End If
                    ElseIf BDBUF_TBMOV_CPTIP.Value = "S" Then 'Crédito
                        FG_1.TextMatrix(Num, 3) = Val(BDBUF_TBMOV_CPVAL.Value)
                        If Num = 1 Then
                            FG_1.TextMatrix(Num, 4) = Val(Saldo) - Val(FG_1.TextMatrix(Num, 3))
                        Else
                            FG_1.TextMatrix(Num, 4) = Val(FG_1.TextMatrix(Num - 1, 4)) - Val(FG_1.TextMatrix(Num, 3))
                        End If
                    End If
                    Num = Num + 1
                End If
                Rec = Rec + 1
            End If
            BDBUF_TBMOV.MoveNext
        Loop
        DataIni = DateAdd("d", 1, DataIni)
        If DataIni > Date Then Exit Do
    Loop
End Sub

