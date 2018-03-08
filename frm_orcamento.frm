VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form frm_orcamento 
   BackColor       =   &H00404080&
   Caption         =   "Orçamento"
   ClientHeight    =   8145
   ClientLeft      =   6435
   ClientTop       =   2730
   ClientWidth     =   9000
   FillColor       =   &H80000012&
   ForeColor       =   &H80000004&
   Icon            =   "frm_orcamento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   9000
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2505
      ItemData        =   "frm_orcamento.frx":08CA
      Left            =   240
      List            =   "frm_orcamento.frx":08CC
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Inserir layout"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   18
      Top             =   3480
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox texto 
      Height          =   3255
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5741
      _Version        =   393217
      BackColor       =   -2147483644
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frm_orcamento.frx":08CE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Tabulação"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      Top             =   3480
      Width           =   1455
   End
   Begin VB.PictureBox IMPRESSORA 
      Height          =   480
      Left            =   8280
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   21
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Inserir cabeçalho"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      MaskColor       =   &H8000000C&
      TabIndex        =   15
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton LINHA 
      BackColor       =   &H8000000C&
      Caption         =   "Linha horizontal"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H8000000C&
      TabIndex        =   14
      Top             =   3480
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00404080&
      Caption         =   "Buscar p/endereço"
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
      Left            =   120
      TabIndex        =   13
      Top             =   360
      Width           =   1935
   End
   Begin VB.OptionButton option1 
      BackColor       =   &H00404080&
      Caption         =   "Buscar p/nome"
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
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Apaga"
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
      Height          =   735
      Left            =   2520
      Picture         =   "frm_orcamento.frx":094E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7320
      Width           =   1200
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Altera"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Picture         =   "frm_orcamento.frx":0D90
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7320
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1305
      ItemData        =   "frm_orcamento.frx":11D2
      Left            =   2400
      List            =   "frm_orcamento.frx":11D4
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Imprimir"
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
      Height          =   735
      Left            =   5280
      Picture         =   "frm_orcamento.frx":11D6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   1200
   End
   Begin VB.CommandButton Command3 
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
      Height          =   735
      Left            =   6480
      Picture         =   "frm_orcamento.frx":1618
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7320
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   " Nova pesquisa"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      Picture         =   "frm_orcamento.frx":1A5A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7320
      Width           =   1560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gravar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Picture         =   "frm_orcamento.frx":1E9C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7320
      Width           =   1200
   End
   Begin VB.TextBox nome 
      BackColor       =   &H8000000B&
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   8880
      X2              =   9000
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label dados_cli 
      BackColor       =   &H00404080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   975
      Left            =   240
      TabIndex        =   20
      Top             =   840
      Width           =   8175
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000005&
      X1              =   7440
      X2              =   7440
      Y1              =   3360
      Y2              =   3960
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   5760
      X2              =   5760
      Y1              =   3360
      Y2              =   3960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   2160
      X2              =   2160
      Y1              =   720
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   7680
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label numero 
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
      Height          =   375
      Left            =   7560
      TabIndex        =   8
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404080&
      Caption         =   "Orçamento: "
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
      Left            =   7680
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label data 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9000
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      Caption         =   "Nome cliente:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frm_orcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private cod_cli, numero_ant As Long
Private orcamento, cad_cli As Recordset
Private cod_orca As Long
Public nome_pesq As String


Private Sub Command1_Click()

If Val(cod_cli) = 0 Then
     MsgBox "Nenhum cliente foi selecionado.", vbExclamation, "Atenção"
Else
orcamento.AddNew
  'orcamento("cod_orca") = Val(numero.Caption)
  orcamento("cod_cli") = cod_cli
  orcamento("data") = Format(Date, "dd/mm/yy")
  orcamento("texto") = texto.Text
orcamento.Update

texto.Text = ""

'numero.Caption = Val(numero.Caption) + 1
'numero_ant = Numero.Caption
Call mostra_orca
End If
End Sub



Private Sub Command2_Click()
Dim k As Integer

If texto.Text <> "" Then
  k = MsgBox("Atenção: Todas as alterações do ORÇAMENTO serão perdidas, deseja continuar ?", vbYesNo, "Pesquisa")
  If k = 7 Then
    
     texto.SetFocus
     Exit Sub
  End If
End If
List1.Visible = True
texto.Text = ""
Numero.Caption = ""
frm_orcamento.Cls

List2.Visible = False
nome.Enabled = True
Command4.Enabled = False
Command5.Visible = False
Command6.Enabled = False
texto.Enabled = False

End Sub

Private Sub Command3_Click()
frm_cad_cli.bt_orca.Enabled = True
Unload frm_orcamento
End Sub

Private Sub Command4_Click()
On Error GoTo arq_erro


IMPRESSORA.ReplaceSelectionFormula ("{orcamento.cod_orca} = " & cod_orca)

IMPRESSORA.Action = 1

Exit Sub

arq_erro:

  If (Err.Number = 20513) Or (Err.Number = 20526) Then
     MsgBox "Impressora não esta pronta", vbCritical, "Atenção"
  Else
     MsgBox " Houve um erro ao tentar imprimir: " & Err.Description, vbCritical, "Atenção"
  End If
End Sub

Private Sub Command5_Click()

orcamento.Edit
  orcamento("texto") = texto.Text
orcamento.Update

texto.Text = ""
Numero.Caption = "" 'numero_ant
Command5.Visible = False
Command4.Enabled = False

End Sub

Private Sub Command6_Click()
Dim msg As Integer
Dim indice As Integer

msg = MsgBox("Confirma deleção do registro", vbYesNo, "Pesquisa")
If msg = 6 Then
 orcamento.Delete
 texto.Text = ""
 For indice = 0 To List2.ListCount - 1
  If List2.Selected(indice) Then
       List2.RemoveItem (indice)
       Exit For
  End If
 Next indice
 Command6.Enabled = False
 Command5.Visible = False
 Numero.Caption = "" 'numero_ant
 Command4.Enabled = False
End If
End Sub



Private Sub Command7_Click()
If Val(cod_cli) = 0 Then
     MsgBox "Nenhum cliente foi selecionado.", vbExclamation, "Atenção"
     Else

texto.Text = texto.Text & "Descrição                           Valor/Unidade   Quantidade        Total" & (Chr(13) & Chr(10))
texto.SelStart = Len(texto.Text)
texto.SetFocus

End If

End Sub

Private Sub Command8_Click()
If Val(cod_cli) = 0 Then
     MsgBox "Nenhum cliente foi selecionado.", vbExclamation, "Atenção"
     Else
     

If Mid(texto.Text, Len(texto.Text), 1) = " " Then

  texto.Text = Mid(texto.Text, 1, Len(texto.Text) - 1) & "          "

Else
      texto.Text = texto.Text & "          "
End If
texto.SelStart = Len(texto.Text)
texto.SetFocus

End If


End Sub


Private Sub Command9_Click()
If Val(cod_cli) = 0 Then
     MsgBox "Nenhum cliente foi selecionado.", vbExclamation, "Atenção"
     Else


texto.Text = "Modelo: " & (Chr(13) & Chr(10))
texto.Text = texto.Text & " " & (Chr(13) & Chr(10))
texto.Text = texto.Text & "Ano:                     Placa: " & (Chr(13) & Chr(10))
texto.Text = texto.Text & " " & (Chr(13) & Chr(10))
texto.Text = texto.Text & " " & (Chr(13) & Chr(10))
texto.Text = texto.Text & "----------------------------------------------------------------------------------" & (Chr(13) & Chr(10))
texto.Text = texto.Text & "Descrição                            Valor/Unidade      Quantidade          Total" & (Chr(13) & Chr(10))
texto.Text = texto.Text & "----------------------------------------------------------------------------------" & (Chr(13) & Chr(10))
texto.Text = texto.Text & " " & (Chr(13) & Chr(10))
texto.Text = texto.Text & "----------------------------------------------------------------------------------" & (Chr(13) & Chr(10))
texto.Text = texto.Text & " " & (Chr(13) & Chr(10))
texto.Text = texto.Text & " " & (Chr(13) & Chr(10))
texto.Text = texto.Text & "                       Forma de pagto: 3X R$   " & (Chr(13) & Chr(10))
texto.Text = texto.Text & "                                  ou" & (Chr(13) & Chr(10))
texto.Text = texto.Text & "                            à vista: R$ "

End If

End Sub

Private Sub Form_Load()

Set cad_cli = frm_principal.arquivo.OpenRecordset("select * from clientes order by nome", dbOpenDynaset)



Set orcamento = frm_principal.arquivo.OpenRecordset("orcamento", dbOpenTable)
   
    
data.Caption = Date



End Sub



Private Sub Form_Unload(Cancel As Integer)
cad_cli.Close
orcamento.Close
frm_principal.Enabled = True

End Sub





Private Sub LINHA_Click()

If Val(cod_cli) = 0 Then
     MsgBox "Nenhum cliente foi selecionado.", vbExclamation, "Atenção"
     Else

texto.Text = texto.Text & "----------------------------------------------------------------------------" & (Chr(13) & Chr(10))

texto.SelStart = Len(texto.Text)
texto.SetFocus

End If
End Sub



Private Sub List1_Click()
On Error GoTo PesqErro
Dim cod As String

List1.Visible = False
cod = Trim(Mid(List1.Text, 1, 4))

Set cad_cli = frm_principal.arquivo.OpenRecordset("select * from clientes where cod = " & cod, dbOpenDynaset)

dados_cli.Caption = " Nome:         " & cad_cli.Fields(1) & Chr(13)
dados_cli.Caption = dados_cli.Caption & "  Endereço:    " & cad_cli.Fields(2) & Chr(13)
dados_cli.Caption = dados_cli.Caption & "  Bairro:         " & cad_cli.Fields(5) & Chr(13)
dados_cli.Caption = dados_cli.Caption & "  Cidade:        " & cad_cli.Fields(4)

texto.Enabled = True
texto.SetFocus
cod_cli = cad_cli.Fields(0)
nome.Text = ""

Call mostra_orca

Exit Sub
PesqErro:

  MsgBox "Houve um erro na pesquisa." & Err.Description, vbCritical, "Atenção"
  
End Sub


Private Sub List2_Click()
Dim k As String
Dim i As Integer

If texto.Text <> "" Then
 k = MsgBox("Atenção: Todas as alterações do ORÇAMENTO serão perdidas, deseja continuar ?", vbYesNo, "Pesquisa")
  If k = 7 Then
    
     texto.SetFocus
     Exit Sub
  End If
End If
 For i = 10 To Len(List2.Text)
   k = Mid(List2.Text, i, 1)
   If k = ":" Then
      cod_orca = Trim(Mid(List2.Text, i + 1, 15))
      Exit For
   End If
 Next i
orcamento.Index = "cod_orca"
orcamento.Seek "=", cod_orca
Command4.Enabled = True
Command6.Enabled = True
data = orcamento.Fields(2)
texto.Text = ""
texto.TextRTF = orcamento.Fields(3)
Numero.Caption = orcamento.Fields(0)
Command5.Visible = True

End Sub



Public Sub nome_Change()

If option1.Value Then

Set cad_cli = frm_principal.arquivo.OpenRecordset("select * from clientes where nome like '" & nome.Text & "*'", dbOpenDynaset)
List1.Clear

Do While Not cad_cli.EOF
   List1.AddItem Format(cad_cli.Fields(0), "0000") & cad_cli.Fields(1) & " --- " & cad_cli.Fields(2)
   cad_cli.MoveNext
Loop
   
   If List1.ListCount = 1 Then
      List1.ListIndex = 0
   End If
Else

Set cad_cli = frm_principal.arquivo.OpenRecordset("select * from clientes where endereco like '" & nome.Text & "*'", dbOpenDynaset)
List1.Clear

Do While Not cad_cli.EOF
   List1.AddItem Format(cad_cli.Fields(0), "0000") & cad_cli.Fields(2) & "  /   " & cad_cli.Fields(1)
   cad_cli.MoveNext
Loop
   
   If List1.ListCount = 1 Then
      List1.ListIndex = 0
   End If
End If
End Sub

Private Sub nome_GotFocus()

If option1.Value Then

  Set cad_cli = frm_principal.arquivo.OpenRecordset("select * from clientes order by nome", dbOpenDynaset)
  List1.Clear
  List1.Visible = True
  Do While Not cad_cli.EOF
     List1.AddItem cad_cli.Fields(0) & Space(15 - ((Len(cad_cli.Fields(0))) * 2)) & cad_cli.Fields(1) & " --- " & cad_cli.Fields(2)
     cad_cli.MoveNext
  Loop
Else
   
  Set cad_cli = frm_principal.arquivo.OpenRecordset("select * from clientes order by endereco", dbOpenDynaset)
  List1.Clear
  List1.Visible = True
  Do While Not cad_cli.EOF
     List1.AddItem cad_cli.Fields(0) & Space(15 - ((Len(cad_cli.Fields(0))) * 2)) & cad_cli.Fields(2) & " --- " & cad_cli.Fields(1)
     cad_cli.MoveNext
  Loop
End If
End Sub



Private Sub nome_LostFocus()
List1.Visible = False

End Sub





Private Sub Option1_Click()
Label2.Caption = "Nome cliente:"
'nome.SetFocus
End Sub

Private Sub Option2_Click()
Label2.Caption = "Endereço:"
'nome.SetFocus
End Sub



Private Sub texto_Change()


If texto.Text = "" Then
   Command1.Enabled = False
   Command5.Enabled = False
   Command8.Enabled = False
Else


nome.Enabled = False

Command1.Enabled = True
Command5.Enabled = True
Command8.Enabled = True
End If


End Sub

Private Sub mostra_orca()
Dim orca As Recordset

Set orca = frm_principal.arquivo.OpenRecordset("select * from orcamento where cod_cli = " & cod_cli & " order by data", dbOpenDynaset)

List2.Clear
  Do While Not orca.EOF
     List2.AddItem orca.Fields(2) & " Orçamento número:  " & orca.Fields(0)
     orca.MoveNext
  Loop
  
  If List2.ListCount = 0 Then
     List2.Visible = False
  Else
     List2.Visible = True
  End If
  
orca.Close
Command6.Enabled = False
End Sub





