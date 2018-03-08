VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_certif 
   BackColor       =   &H00404080&
   Caption         =   "Certificado de Garantia"
   ClientHeight    =   5745
   ClientLeft      =   1650
   ClientTop       =   1695
   ClientWidth     =   8940
   Icon            =   "frm_ceritf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport IMPRESSORA 
      Left            =   6000
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   ".\imp_certificado.rpt"
      WindowLeft      =   0
      WindowTop       =   0
      WindowWidth     =   0
      WindowHeight    =   0
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3150
      ItemData        =   "frm_ceritf.frx":08CA
      Left            =   240
      List            =   "frm_ceritf.frx":08CC
      TabIndex        =   17
      Top             =   960
      Width           =   8535
   End
   Begin MSMask.MaskEdBox data 
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483637
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox periodo 
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
      Height          =   375
      Left            =   5520
      MaxLength       =   2
      TabIndex        =   12
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox bairro 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   10
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Nova pesquisa"
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
      Left            =   3840
      Picture         =   "frm_ceritf.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4800
      Width           =   1455
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
      Left            =   7440
      Picture         =   "frm_ceritf.frx":0D10
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
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
      Left            =   240
      Picture         =   "frm_ceritf.frx":1152
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00404080&
      Caption         =   "Por nome"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00404080&
      Caption         =   "Por endereço"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   7200
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox nome 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   5175
   End
   Begin VB.TextBox ender 
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
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   8535
   End
   Begin VB.Label Label6 
      BackColor       =   &H00404080&
      Caption         =   "COMO VENCIMENTO"
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
      TabIndex        =   15
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404080&
      Caption         =   "A   PARTIR  DA   EMISSÃO   DESTE   CERTIFICADO   DEVIDAMENTE   ASSINADO,   TENDO "
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
      TabIndex        =   14
      Top             =   3240
      Width           =   8055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404080&
      Caption         =   "ANOS CONTADOS"
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
      Left            =   6600
      TabIndex        =   13
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      Caption         =   "CERTIFICAMOS NOSSO PRODUTO PELO PERÍODO DE"
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
      TabIndex        =   11
      Top             =   2760
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404080&
      Caption         =   "Bairro:"
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
      Left            =   5520
      TabIndex        =   9
      Top             =   1200
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9480
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label titulo 
      BackColor       =   &H00404080&
      Caption         =   "Nome do cliente:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404080&
      Caption         =   "Endereço:"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
End
Attribute VB_Name = "frm_certif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private tab_aux As Recordset
Private cad_cli As Recordset
Private cod As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Command1_Click()
On Error GoTo arq_erro

Dim matriz(12) As String
Dim imp_certif As Recordset
Dim mes As String
Dim dia As String
Dim ano As String
Dim datatexto As String
Dim prox_numer As Integer


'grava a data de envio de certif, para calculo do venc. da garantia

cad_cli.Edit

  cad_cli("data_serv") = Format(CDate(Date), "dd/mm/yyyy")

cad_cli.Update


Set imp_certif = frm_principal.arquivo.OpenRecordset("certificado", dbOpenTable)

matriz(1) = "janeiro  "
matriz(2) = "fevereiro"
matriz(3) = "março    "
matriz(4) = "abril    "
matriz(5) = "maio     "
matriz(6) = "junho    "
matriz(7) = "julho    "
matriz(8) = "agosto   "
matriz(9) = "setembro "
matriz(10) = "outubro  "
matriz(11) = "novembro "
matriz(12) = "dezembro "



mes = Val(Mid(Date, 4, 2))
dia = Mid(Date, 1, 2)

If Val(Mid(Format(Date, "dd/mm/yy"), 7, 2)) > 35 Then
    ano = 19 & Mid(Date, 7, 2)
Else
   ano = 20 & Mid(Date, 7, 2)
End If

datatexto = "São Paulo, " & dia & " de " & matriz(mes) & " " & ano

'----grava certificado para impressão

imp_certif.MoveFirst
imp_certif.Edit
   imp_certif("nome") = nome.Text
   imp_certif("endereco") = ender.Text
   imp_certif("periodo") = periodo.Text
   imp_certif("data_texto") = datatexto
   imp_certif("data") = data
imp_certif.Update

prox_numer = imp_certif.Fields(3) + 1

'-- imprime

    DoEvents
    Sleep (555)
IMPRESSORA.ReportFileName = App.Path + "\imp_certificado.rpt"

IMPRESSORA.Action = 0

'apaga-------
imp_certif.MoveFirst
imp_certif.Delete
imp_certif.Close
'cria novoe registro em branco com o proximo numero

frm_principal.arquivo.Execute "insert into certificado values(null,null,null," & prox_numer & ",null,null)"

Exit Sub

arq_erro:

  If (Err.Number = 20513) Or (Err.Number = 20526) Then
      MsgBox "Impressora não esta pronta", , "Pesquisa"

  End If
  
End Sub

Private Sub Command2_Click()
cad_cli.Close
Unload frm_certif

End Sub

Private Sub Command3_Click()
cod = 0
If Option2.Value Then
   nome.Text = ""
   List1.Top = 1560
     Set tab_aux = frm_principal.arquivo.OpenRecordset("select * from clientes order by nome", dbOpenDynaset)
        List1.Clear
        Do While Not tab_aux.EOF
            List1.AddItem tab_aux.Fields(0) & Space(15 - ((Len(tab_aux.Fields(0))) * 2)) & tab_aux.Fields(1) & " --- " & tab_aux.Fields(2)
            tab_aux.MoveNext
        Loop
        If List1.ListCount = 1 Then
            List1.ListIndex = 0
        End If
        List1.Visible = True
   
   
   
Else
   ender.Text = ""
   List1.Top = 2280
    Set tab_aux = frm_principal.arquivo.OpenRecordset("select * from clientes order by endereco", dbOpenDynaset)
        List1.Clear
        Do While Not tab_aux.EOF
            List1.AddItem tab_aux.Fields(0) & Space(15 - ((Len(tab_aux.Fields(0))) * 2)) & tab_aux.Fields(2) & " --- " & tab_aux.Fields(1)
            tab_aux.MoveNext
        Loop
        If List1.ListCount = 1 Then
            List1.ListIndex = 0
        End If
        List1.Visible = True
   
   
End If


End Sub


Private Sub ender_Change()

If Option3.Value Then
If ender.Text <> "" Then
    If cod = 0 Then
        Set tab_aux = frm_principal.arquivo.OpenRecordset("select * from clientes where endereco like '" & ender.Text & "*' order by endereco", dbOpenDynaset)
        
        
        List1.Clear
        Do While Not tab_aux.EOF
            List1.AddItem tab_aux.Fields(0) & Space(15 - ((Len(tab_aux.Fields(0))) * 2)) & tab_aux.Fields(2) & " --- " & tab_aux.Fields(1)
            tab_aux.MoveNext
        Loop
        If List1.ListCount = 1 Then
            List1.ListIndex = 0
        End If
    End If

  
End If
End If
End Sub

Private Sub Form_Load()
Set cad_cli = frm_principal.arquivo.OpenRecordset("clientes", dbOpenTable)
Set tab_aux = frm_principal.arquivo.OpenRecordset("select * from clientes  order by nome", dbOpenDynaset)
List1.Clear

Do While Not tab_aux.EOF
    List1.AddItem tab_aux.Fields(0) & Space(15 - ((Len(tab_aux.Fields(0))) * 2)) & tab_aux.Fields(1) & " --- " & tab_aux.Fields(2)
    tab_aux.MoveNext
Loop
If List1.ListCount = 1 Then
   List1.ListIndex = 0
ElseIf List1.ListCount > 1 Then
   List1.Visible = True
Else
    Command1.Enabled = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
tab_aux.Close
frm_principal.Enabled = True

End Sub

Private Sub List1_Click()
cod = Trim(Mid(List1.Text, 1, 10))

'Set tab_aux = frm_principal.arquivo.OpenRecordset("select * from clientes where cod = " & cod, dbOpenDynaset)
 
cad_cli.Index = "cod"

cad_cli.Seek "=", cod


nome.Text = cad_cli.Fields(1)
ender.Text = cad_cli.Fields(2)
bairro.Text = cad_cli.Fields(5)

List1.Visible = False

End Sub

Private Sub nome_Change()
If Option2.Value Then

If nome.Text <> "" Then
 If cod = 0 Then
  Set tab_aux = frm_principal.arquivo.OpenRecordset("select * from clientes where nome like '" & nome.Text & "*' order by nome", dbOpenDynaset)

    List1.Clear

    Do While Not tab_aux.EOF
        List1.AddItem tab_aux.Fields(0) & Space(15 - ((Len(tab_aux.Fields(0))) * 2)) & tab_aux.Fields(1) & " --- " & tab_aux.Fields(2)
        tab_aux.MoveNext
    Loop
    If List1.ListCount = 1 Then
        List1.ListIndex = 0
    End If
  End If

  

End If
End If
End Sub

Private Sub Option2_Click()
Call Command3_Click

End Sub

Private Sub Option3_Click()
Call Command3_Click
End Sub
