VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form FRMPESQUISA_CLI 
   BackColor       =   &H00404080&
   Caption         =   "Pesquisa"
   ClientHeight    =   6240
   ClientLeft      =   7740
   ClientTop       =   3405
   ClientWidth     =   7470
   Icon            =   "FRMPESQUISA_CLI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7470
   Begin VB.ListBox List1 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      ItemData        =   "FRMPESQUISA_CLI.frx":08CA
      Left            =   0
      List            =   "FRMPESQUISA_CLI.frx":08CC
      TabIndex        =   28
      Top             =   840
      Width           =   7470
   End
   Begin VB.TextBox email 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   3240
      MaxLength       =   20
      MouseIcon       =   "FRMPESQUISA_CLI.frx":08CE
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "Para enviar Email clique duas vezes"
      Top             =   2880
      Width           =   3975
   End
   Begin MSMAPI.MAPISession sessao 
      Left            =   3240
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages mensagem 
      Left            =   3840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.CommandButton Command6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      MaskColor       =   &H000040C0&
      Picture         =   "FRMPESQUISA_CLI.frx":1198
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      MaskColor       =   &H80000006&
      Picture         =   "FRMPESQUISA_CLI.frx":11FA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Sair "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      Picture         =   "FRMPESQUISA_CLI.frx":1556
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5160
      Width           =   1215
   End
   Begin MSMask.MaskEdBox telefone 
      Height          =   300
      Left            =   960
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   -2147483637
      MaxLength       =   15
      Mask            =   "(####)####-####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox cidade 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   960
      MaxLength       =   25
      TabIndex        =   3
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox bairro 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   960
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox ender 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   960
      MaxLength       =   49
      TabIndex        =   2
      Top             =   1920
      Width           =   6375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Nova Pesquisa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      Picture         =   "FRMPESQUISA_CLI.frx":1998
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apagar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      Picture         =   "FRMPESQUISA_CLI.frx":1DDA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Alterar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      Picture         =   "FRMPESQUISA_CLI.frx":221C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox nome 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   960
      MaxLength       =   40
      TabIndex        =   0
      Top             =   960
      Width           =   5895
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00404080&
      Caption         =   "Pesquisa por endereço"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   4680
      TabIndex        =   26
      Top             =   360
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404080&
      Caption         =   "Pesquisa por nome"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   4680
      TabIndex        =   25
      Top             =   120
      Value           =   -1  'True
      Width           =   2175
   End
   Begin MSMask.MaskEdBox cep 
      Height          =   300
      Left            =   4920
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   -2147483637
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99999-999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label10 
      BackColor       =   &H00404080&
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   2640
      TabIndex        =   27
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label9 
      BackColor       =   &H00404080&
      Caption         =   "Cep: "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   4440
      TabIndex        =   24
      Top             =   2520
      Width           =   615
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   7560
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   7560
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label data_m 
      BackColor       =   &H00404080&
      Caption         =   "-------------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   5880
      TabIndex        =   23
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label data_s 
      BackColor       =   &H00404080&
      Caption         =   "-------------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   22
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00404080&
      Caption         =   "Mala  enviada em: "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   4320
      TabIndex        =   21
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00404080&
      Caption         =   "Ultimo serviço realizado em : "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   -1560
      X2              =   7440
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label cod 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   495
      Left            =   1800
      TabIndex        =   19
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00404080&
      Caption         =   "Número do cliente:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404080&
      Caption         =   "Telefone:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404080&
      Caption         =   "Bairro:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404080&
      Caption         =   "Cidade:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      Caption         =   "Endereço:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404080&
      Caption         =   "Nome:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   960
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   7560
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "FRMPESQUISA_CLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cad_cli As Recordset
Private Sub limpa_cad()
cod.Caption = ""
nome.Text = ""
ender.Text = ""
Bairro.Text = ""
cidade.Text = ""
telefone.Mask = ""
telefone.Text = ""
telefone.MaxLength = 1
data_m = "-------------"
data_s = "-------------"
cep.Mask = ""
cep.Text = ""
cep.MaxLength = 9
email.Text = ""

End Sub
Private Sub pesquisa(SQL)

If Option1.Value Then
 If SQL = "" Then
 
    Set cad_cli = frm_principal.arquivo.OpenRecordset("Select * from clientes order by nome", dbOpenDynaset)
 Else
    Set cad_cli = frm_principal.arquivo.OpenRecordset("Select * from clientes where " & SQL & "order by nome", dbOpenDynaset)
    
 End If
 List1.Clear
    Do While Not cad_cli.EOF
        List1.AddItem cad_cli.Fields(1) & String(41 - Len(cad_cli.Fields(1)), ".") & cad_cli.Fields(2)
        cad_cli.MoveNext
   
    Loop
    
ElseIf Option2.Value Then
 If SQL = "" Then
    Set cad_cli = frm_principal.arquivo.OpenRecordset("Select * from clientes order by endereco", dbOpenDynaset)
 Else
    Set cad_cli = frm_principal.arquivo.OpenRecordset("Select * from clientes where " & SQL & "order by endereco", dbOpenDynaset)
 End If
 List1.Clear
    Do While Not cad_cli.EOF
        List1.AddItem cad_cli.Fields(2) & String(52 - Len(cad_cli.Fields(2)), ".") & cad_cli.Fields(1)
        cad_cli.MoveNext
   
    Loop
    
End If

If cad_cli.RecordCount > 0 Then
       cad_cli.MovePrevious
End If

End Sub
Private Sub mostra_cad()
Call limpa_cad
cod.Caption = cad_cli.Fields(0)
nome.Text = cad_cli.Fields(1)
ender.Text = cad_cli.Fields(2)
Bairro.Text = cad_cli.Fields(5)
cidade.Text = cad_cli.Fields(4)
telefone.Mask = ""
telefone.Text = cad_cli.Fields(3)
email.Text = cad_cli("email")
If cad_cli.Fields(6) <> "" Then data_s = cad_cli.Fields(6)
If cad_cli.Fields(7) <> "" Then data_m = cad_cli.Fields(7)
If cad_cli("cep") <> "" Then cep.Text = cad_cli("cep")

End Sub





Private Sub Command1_Click()
If nome.Text = "" Then
    MsgBox "Nome não pode estar em branco", , "Pesquisa"
ElseIf ender.Text = "" Then
     MsgBox "Endereço não pode estar em branco", , "Pesquisa"
ElseIf Bairro.Text = "" Then
     MsgBox "Bairro não pode estar em branco", , "Pesquisa"
ElseIf cidade.Text = "" Then
     MsgBox "Cidade não pode estar em branco", , "Pesquisa"


Else

If telefone.Text = "" Then
    telefone.Mask = "(9999)9999-9999"
End If

cad_cli.Edit
  'cad_cli("cod") = cod.Caption
  cad_cli("nome") = nome.Text
  cad_cli("Endereco") = ender.Text
  cad_cli("bairro") = Bairro.Text
  cad_cli("cidade") = cidade.Text
  cad_cli("telefone") = telefone.Text
  cad_cli("cep") = cep.Text
  cad_cli("email") = email.Text
cad_cli.Update

 MsgBox "Alteração efetuada com sucesso", , "Pesquisa"

End If

End Sub

Private Sub Command2_Click()
Dim orcatemp As Recordset
Dim msg As Integer
Dim x As String

msg = MsgBox("Confirma deleção do registro", vbYesNo, "Pesquisa")

If msg = 6 Then
  x = cad_cli.Fields(0)
  cad_cli.Delete
  cad_cli.MoveNext
 If cad_cli.EOF Then
   cad_cli.MovePrevious
   
   '--- esta seçao apaga todos os orçamentos ligados a este cliente
   Set orcatemp = frm_principal.arquivo.OpenRecordset("select * from orcamento where cod_cli = " & x, dbOpenDynaset)
   
      Do While Not orcatemp.EOF 'apaga orcamentos ligados ao cliente
         orcatemp.Delete
         orcatemp.MoveNext
      Loop
      
   If cad_cli.BOF Then
     Command1.Enabled = False
     Command2.Enabled = False
     Command5.Enabled = False
     Command6.Enabled = False
     Call limpa_cad
   Else
     Call mostra_cad
   End If
 Else
   Call mostra_cad
 End If
End If


End Sub

Private Sub Command3_Click()
Call limpa_cad
Command1.Enabled = False
Command2.Enabled = False
List1.Visible = True
If Option1.Value Then
 
  nome.Text = ""
  nome.SetFocus
  Call pesquisa("")
  
ElseIf Option2.Value Then
  
  
  nome.Enabled = False
  ender.Enabled = True
  ender.Text = ""
  ender.SetFocus
  Call pesquisa("")
End If



End Sub

Private Sub Command4_Click()
frm_cad_cli.bt_pesq.Enabled = True
Unload FRMPESQUISA_CLI

End Sub

Private Sub Command5_Click()

If cad_cli.BOF Then

   cad_cli.MoveNext
   
ElseIf Not cad_cli.BOF Then
   cad_cli.MovePrevious
   If cad_cli.BOF Then
     cad_cli.MoveNext
   Else
     Call mostra_cad
   End If


End If
End Sub

Private Sub Command6_Click()

If cad_cli.EOF Then

  cad_cli.MovePrevious
ElseIf Not cad_cli.EOF Then
   cad_cli.MoveNext
   If cad_cli.EOF Then
     cad_cli.MovePrevious
   Else
   
     Call mostra_cad
   End If

End If
   

   

End Sub






Private Sub email_DblClick()
On Error GoTo sair
sessao.Action = 1
mensagem.SessionID = sessao.SessionID
mensagem.Compose
mensagem.RecipAddress = email.Text
mensagem.AddressResolveUI = True
mensagem.ResolveName
mensagem.Send True
sair:
    Exit Sub

End Sub

Private Sub ender_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
      If ender.Text <> "" Then
        Call pesquisa("endereco like '" & ender.Text & "*'")
        If List1.ListCount = 1 Then
           List1.ListIndex = 0
        End If
      End If
End If
   End Sub

Private Sub Form_Load()
Call pesquisa("")
End Sub

Private Sub Form_Unload(Cancel As Integer)
cad_cli.Close
frm_principal.Enabled = True

End Sub

Private Sub List1_Click()
Dim pesquisa As String
Dim texto As String
Dim carac As String
Dim i As Integer

texto = ""
carac = ""
For i = 1 To Len(List1.Text)
   carac = Mid(List1.Text, i, 1)
    If carac = "." Then
       Exit For
    End If
   texto = texto & carac
Next i


Set cad_cli = frm_principal.arquivo.OpenRecordset("select * from clientes order by nome", dbOpenDynaset)

If Option1.Value Then
    
   pesquisa = "nome = '" & Trim(texto) & "'"
   cad_cli.FindFirst pesquisa
  
   Call mostra_cad
   
ElseIf Option2.Value Then
   
   pesquisa = "endereco = '" & texto & "'"
   cad_cli.FindFirst pesquisa
   Call mostra_cad
End If

 Command1.Enabled = True
 Command2.Enabled = True
 List1.Visible = False
 nome.Enabled = True
 ender.Enabled = True
End Sub



Private Sub nome_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If nome.Text <> "" Then
  

        Call pesquisa("nome like '" & nome.Text & "*'")
        If List1.ListCount = 1 Then
           List1.ListIndex = 0
        End If
End If
End If
End Sub


Private Sub Option1_Click()
Call pesquisa("")
End Sub

Private Sub Option2_Click()
Call pesquisa("")
End Sub
