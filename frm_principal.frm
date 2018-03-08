VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frm_principal 
   BackColor       =   &H00404080&
   Caption         =   "PrestServ"
   ClientHeight    =   7755
   ClientLeft      =   5775
   ClientTop       =   2790
   ClientWidth     =   10995
   DrawStyle       =   1  'Dash
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frm_principal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   10995
   Begin MSComDlg.CommonDialog Comm_busca_bd 
      Left            =   4320
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Conectar outro Bd"
      Filter          =   "*.mdb | clientes.mdb"
      InitDir         =   "app.path"
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   7920
      Picture         =   "frm_principal.frx":08CA
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
   Begin VB.Menu mnu_cli 
      Caption         =   "&Tarefas"
      Begin VB.Menu mnu_adicionar_cli 
         Caption         =   "&Adicionar novo cliente"
      End
      Begin VB.Menu mnuseparetor1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_pesquisa_cli 
         Caption         =   "&Pesquisar clientes"
      End
      Begin VB.Menu mnuseparetor2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_orcamento 
         Caption         =   "&Orçamento"
      End
      Begin VB.Menu mnuseparetor7 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_vencimentos 
         Caption         =   "&Identifcar oportunidade (MD)"
      End
      Begin VB.Menu separetor11 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_grava_emp_tel 
         Caption         =   "&Informar dados do estabelecimento"
      End
      Begin VB.Menu SEPARETOR13 
         Caption         =   "-"
      End
      Begin VB.Menu mnusair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu mnu_rel 
      Caption         =   "&Relatorios"
      Begin VB.Menu mnu_rel_certificado 
         Caption         =   "Imprime &certificado"
      End
      Begin VB.Menu mnuseparetor3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_rel_orcamento 
         Caption         =   "Imprimir &orçamento"
      End
      Begin VB.Menu mnuseparetor4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_rel_recibo 
         Caption         =   "Imprimir &recibo"
      End
   End
   Begin VB.Menu MNU_SEG 
      Caption         =   "C&ontrole de acesso"
      Begin VB.Menu mnu_con_outro_bd 
         Caption         =   "&Conectar outro BD"
      End
      Begin VB.Menu separetor20 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_senha 
         Caption         =   "&Senha de acesso"
      End
   End
   Begin VB.Menu mnu_help1 
      Caption         =   "Ajuda"
      Begin VB.Menu mnu_help 
         Caption         =   "Sobre o &PrestServ"
      End
      Begin VB.Menu mnu_help2 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "frm_principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public arquivo As Database
Public area As Workspace
Private tab_aux, tab_imp As Recordset
Private tot_recibo As Recordset
Private cod As Long
Public con_bd As String

Option Explicit

Private Sub Form_Load()
Dim teste As Recordset
con_bd = App.Path
On Error GoTo testa_base
frm_principal.Caption = "PrestServ" & Space(100 - Len(con_bd & "\clientes.mdb") * 1.5) & con_bd & "\clientes.mdb"
Set area = DBEngine.Workspaces(0)


Set arquivo = area.OpenDatabase(con_bd & "\clientes.mdb", False, False)

Exit Sub

testa_base:
    If Err.Number = 3024 Or Err.Number = 32755 Then
        MsgBox "Nenhum banco de dados 'CLIENTES.mdb' foi localizado.", vbCritical, "Atenção"
        End
    End If


End Sub

Private Sub Form_Resize()
Picture1.Left = ScaleWidth - 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
arquivo.Close
End Sub



Private Sub mnu_adicionar_cli_Click()
 frm_cad_cli.Show
 frm_principal.Enabled = False

End Sub





Private Sub mnu_con_outro_bd_Click()
On Error GoTo testa_base
' troca o bd apontado e refaz a agenda com base nos novos dados  - beto  04/06/2004
    
    frm_principal.Caption = "PrestServ" & Space(100 - Len(con_bd & "\clientes.mdb") * 1.5) & con_bd & "\clientes.mdb"
    Set area = DBEngine.Workspaces(0)
    Set arquivo = area.OpenDatabase(con_bd & "\clientes.mdb", False, False)
        
    Exit Sub

testa_base:
    If Err.Number = 32755 Then
        Exit Sub
    Else
        MsgBox "Ocorreu um erro ao tentar conectar o banco de dados: " & Err.Description, vbCritical, "Atenção"
        End
    End If
End Sub

Private Sub mnu_grava_emp_tel_Click()
frm_cad_emp_tel.Show
frm_principal.Enabled = False
End Sub

Private Sub mnu_help_Click()
frm_principal.Enabled = True
Info_sistema.Show
End Sub

Private Sub mnu_help2_Click()

fmrHelp.Show
frm_principal.Enabled = False
    

End Sub

Private Sub mnu_orcamento_Click()
frm_orcamento.Show
frm_principal.Enabled = False
End Sub

Private Sub mnu_pesquisa_cli_Click()
FRMPESQUISA_CLI.Show
frm_principal.Enabled = False
End Sub

Private Sub mnu_rel_certificado_Click()
frm_certif.Show
frm_principal.Enabled = False
End Sub

Private Sub mnu_rel_orcamento_Click()
frm_imp_orca.Show
frm_principal.Enabled = False
End Sub
Private Sub mnu_rel_recibo_Click()
frm_recibo.Show
frm_principal.Enabled = False

End Sub

Private Sub mnu_senha_Click()
frm_senha.Show
frm_principal.Enabled = False
End Sub


Private Sub mnu_vencimentos_Click()
frm_vencimento.Show
End Sub

Private Sub mnusair_Click()
End
End Sub

Private Sub mnuteste_Click()
Info_sistema.Show
End Sub
