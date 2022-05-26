VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Tela_Movimento_Dlg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " consulta do movimento"
   ClientHeight    =   2724
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   4224
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2724
   ScaleWidth      =   4224
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FR_Datas 
      Caption         =   "Escolha as datas ( Data inical é 23/10/97 )"
      Height          =   852
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   3732
      Begin MSMask.MaskEdBox TXT_Ate 
         Height          =   288
         Left            =   2760
         TabIndex        =   5
         ToolTipText     =   "Data da emissao do documento"
         Top             =   480
         Width           =   852
         _ExtentX        =   1503
         _ExtentY        =   508
         _Version        =   327680
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   8
         Format          =   "ddddd"
         Mask            =   "99/99/99"
         PromptChar      =   "_"
      End
      Begin VB.OptionButton RD_SelecionarDatas 
         Caption         =   "Selecionar datas"
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1572
      End
      Begin VB.OptionButton RD_TodasDatas 
         Caption         =   "Todos registros"
         Height          =   252
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1452
      End
      Begin MSMask.MaskEdBox TXT_Desde 
         Height          =   288
         Left            =   2760
         TabIndex        =   4
         ToolTipText     =   "Data da emissao do documento"
         Top             =   240
         Width           =   852
         _ExtentX        =   1503
         _ExtentY        =   508
         _Version        =   327680
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   8
         Format          =   "ddddd"
         Mask            =   "99/99/99"
         PromptChar      =   "_"
      End
      Begin VB.Label LB_Ate 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
         Height          =   192
         Left            =   2400
         TabIndex        =   11
         Top             =   480
         Width           =   276
      End
      Begin VB.Label LB_Desde 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   192
         Left            =   2160
         TabIndex        =   10
         Top             =   240
         Width           =   528
      End
   End
   Begin VB.CommandButton BT_Cancelar 
      Caption         =   "&Cancelar"
      Height          =   372
      Left            =   2280
      TabIndex        =   7
      Top             =   2160
      Width           =   1092
   End
   Begin VB.CommandButton BT_OK 
      Caption         =   "&Ok"
      Height          =   372
      Left            =   720
      TabIndex        =   6
      Top             =   2160
      Width           =   1092
   End
   Begin VB.OptionButton RD_Todos 
      Caption         =   "&Exibir todos lançamentos"
      Height          =   252
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   2052
   End
   Begin VB.OptionButton RD_MinhaCC 
      Caption         =   "&Apenas minha C/C"
      Height          =   252
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   1932
   End
   Begin VB.Label LB_Texto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione uma das opeções de consulta do movimento:"
      Height          =   192
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4032
   End
End
Attribute VB_Name = "Tela_Movimento_Dlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BT_Cancelar_Click()
    Tela_Movimento_Dlg.Hide
    Tela_Bufunfa.ST_1.Tab = 0
    Tela_Bufunfa.BT_Consultar.SetFocus
End Sub
Private Sub BT_OK_Click()
    If RD_MinhaCC.Value = True Then
        If RD_TodasDatas.Value = False And RD_SelecionarDatas.Value = False Then
            MsgBox ("Selecione primeiro uma opção.")
            RD_TodasDatas.SetFocus
            Exit Sub
        ElseIf RD_SelecionarDatas.Value = True Then
            If TXT_Desde.Text = "" And TXT_Ate.Text = "" Then
                MsgBox ("Selecione primeiro uma data.")
                TXT_Desde.SetFocus
                Exit Sub
            ElseIf Not IsDate(TXT_Desde.Text) Then
                MsgBox ("Data Inválida.")
                TXT_Desde.SelLength = Len(TXT_Desde.Text)
                TXT_Desde.SetFocus
                Exit Sub
            ElseIf Not IsDate(TXT_Ate.Text) Then
                MsgBox ("Data Inválida.")
                TXT_Ate.SelLength = Len(TXT_Ate.Text)
                TXT_Ate.SetFocus
                Exit Sub
            ElseIf TXT_Desde.Text < #10/23/97# Then
                MsgBox ("Data inicial é menor que a data do sistema - 23/10/97.")
                TXT_Desde.SelLength = Len(TXT_Desde.Text)
                TXT_Desde.SetFocus
                Exit Sub
            ElseIf TXT_Ate.Text > Date Then
                MsgBox ("Data final ainda não existe")
                TXT_Ate.SelLength = Len(TXT_Ate.Text)
                TXT_Ate.SetFocus
                Exit Sub
            ElseIf TXT_Ate.Text < TXT_Desde.Text Then
                MsgBox ("Data final é menor que a data inicial.")
                TXT_Ate.SelLength = Len(TXT_Ate.Text)
                TXT_Ate.SetFocus
                Exit Sub

            End If
        End If
    End If
    
    Tela_Movimento_Dlg.Hide
    Tela_Bufunfa.ST_1.Tab = 0
    Tela_Bufunfa.AtivaTabs (0)
    If RD_MinhaCC.Value = True Then
        Tela_Bufunfa.FR_Movimentos.Visible = False
        Tela_Bufunfa.BotoesTelaEmConsulta (True)
        Tela_Bufunfa.MontaFG
        Tela_Bufunfa.FG_1.Visible = True
    ElseIf RD_Todos.Value = True Then
        Tela_Bufunfa.BotoesTelaEmConsulta (True)
    End If
End Sub
Private Sub RD_MinhaCC_Click()
    FR_Datas.Enabled = True
    RD_TodasDatas.Enabled = True
    RD_SelecionarDatas.Enabled = True
    LB_Desde.Enabled = True
    LB_Ate.Enabled = True
    TXT_Desde.Enabled = True
    TXT_Ate.Enabled = True
    RD_TodasDatas.Value = True
    RD_SelecionarDatas.Value = False
    TXT_Desde.Text = "__/__/__"
    TXT_Ate.Text = "__/__/__"
End Sub
Private Sub RD_MinhaCC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RD_TodasDatas.SetFocus
    End If
End Sub
Private Sub RD_SelecionarDatas_Click()
    LB_Desde.Enabled = True
    LB_Ate.Enabled = True
    TXT_Desde.Enabled = True
    TXT_Ate.Enabled = True
    TXT_Desde.Text = "__/__/__"
    TXT_Ate.Text = "__/__/__"
End Sub
Private Sub RD_SelecionarDatas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TXT_Desde.SetFocus
    End If
End Sub
Private Sub RD_TodasDatas_Click()
    LB_Desde.Enabled = False
    LB_Ate.Enabled = False
    TXT_Desde.Enabled = False
    TXT_Ate.Enabled = False
    TXT_Desde.Text = "__/__/__"
    TXT_Ate.Text = "__/__/__"
End Sub
Private Sub RD_TodasDatas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BT_OK.SetFocus
    End If
End Sub
Private Sub RD_Todos_Click()
    FR_Datas.Enabled = False
    RD_TodasDatas.Enabled = False
    RD_SelecionarDatas.Enabled = False
    LB_Desde.Enabled = False
    LB_Ate.Enabled = False
    TXT_Desde.Enabled = False
    TXT_Ate.Enabled = False
    RD_TodasDatas.Value = False
    RD_SelecionarDatas.Value = False
    TXT_Desde.Text = "__/__/__"
    TXT_Ate.Text = "__/__/__"
End Sub
Private Sub RD_Todos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BT_OK.SetFocus
    End If
End Sub
Private Sub TXT_Ate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BT_OK.SetFocus
    End If
End Sub
Private Sub TXT_Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TXT_Ate.SetFocus
    End If
End Sub
