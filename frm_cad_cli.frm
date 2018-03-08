VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_cad_cli 
   BackColor       =   &H00404080&
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   5550
   ClientLeft      =   8100
   ClientTop       =   3945
   ClientWidth     =   8070
   Icon            =   "frm_cad_cli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8070
   Begin VB.TextBox email 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   11
      Top             =   3120
      Width           =   4695
   End
   Begin MSMask.MaskEdBox cep 
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483637
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99999-999"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton bt_pesq 
      Caption         =   "Pesquisar "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      Picture         =   "frm_cad_cli.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton bt_orca 
      Caption         =   "Orçamento"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      Picture         =   "frm_cad_cli.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4440
      Width           =   1335
   End
   Begin MSMask.MaskEdBox servico 
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483637
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox telefone 
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   2640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483637
      MaxLength       =   14
      Mask            =   "(##) #########"
      PromptChar      =   "_"
   End
   Begin VB.TextBox cidade 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   8
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox Bairro 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   7
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox Endereco 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      MaxLength       =   49
      TabIndex        =   6
      Top             =   720
      Width           =   6615
   End
   Begin VB.TextBox nome 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      MaxLength       =   40
      TabIndex        =   5
      Top             =   240
      Width           =   6615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      Picture         =   "frm_cad_cli.frx":114E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gravar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Picture         =   "frm_cad_cli.frx":1590
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00404080&
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H00404080&
      Caption         =   "Cep: "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      Caption         =   "Data do último serviço realizado:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   8280
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label7 
      BackColor       =   &H00404080&
      Caption         =   "Telefone:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00404080&
      Caption         =   "Cidade:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404080&
      Caption         =   "Bairro:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404080&
      Caption         =   "Endereço:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404080&
      Caption         =   "Nome:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frm_cad_cli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cad_cli As Recordset



Private Sub bt_pesq_Click()

bt_pesq.Enabled = False
FRMPESQUISA_CLI.Show

End Sub

Private Sub bt_orca_Click()
bt_orca.Enabled = False
frm_orcamento.Show
End Sub



Private Sub Command1_Click()
On Error GoTo data_erro
Dim x As Integer

cad_cli.AddNew
  'cad_cli("cod") = cod.Caption
  cad_cli("nome") = nome.Text & " "
  cad_cli("Endereco") = Endereco.Text & " "
  cad_cli("bairro") = Bairro.Text & " "
  cad_cli("cidade") = cidade.Text & " "
  cad_cli("telefone") = telefone.Text
  cad_cli("EMAIL") = email.Text & " "
  If servico = "__/__/____" Then
     cad_cli("data_serv") = Format(CDate(Date), "dd/mm/yyyy")
  ElseIf CDate(servico) Then
  
  cad_cli("data_serv") = servico
  End If
  cad_cli("data_mala") = "----------"
  cad_cli("cep") = cep.Text
  
  
     
cad_cli.Update

Call limpa_form_cli

data_erro:
 x = Err.Number
 If Err.Number = 13 Then
      MsgBox "Data informada não existe", vbCritical, "Pesquisa"
     servico.SetFocus
 End If
 nome.SetFocus
End Sub
     
     
 



Private Sub Command2_Click()

Unload frm_cad_cli
End Sub

Private Sub Command3_Click()
frm_orcamento.Show
End Sub





Private Sub Form_Load()

Set cad_cli = frm_principal.arquivo.OpenRecordset("clientes", dbOpenTable)
   cad_cli.Index = "cod"

   
   

End Sub
Private Sub limpa_form_cli()


nome.Text = ""
Endereco.Text = ""
Bairro.Text = ""
cidade.Text = ""
telefone.Mask = ""
telefone.Text = ""
telefone.Mask = "(9999)9999-9999"
servico.Mask = ""
servico.Text = ""
servico.Mask = "99/99/9999"
cep.Mask = ""
cep.Text = ""
cep.Mask = "99999-999"
email.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
cad_cli.Close
frm_principal.Enabled = True

End Sub


Private Sub nome_Change()
If nome.Text <> "" Then
   Command1.Enabled = True
Else
   Command1.Enabled = False
End If
End Sub
