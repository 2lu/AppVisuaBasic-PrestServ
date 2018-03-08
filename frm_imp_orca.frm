VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_imp_orca 
   BackColor       =   &H00404080&
   Caption         =   "Imprimindo orçamento"
   ClientHeight    =   6795
   ClientLeft      =   1845
   ClientTop       =   810
   ClientWidth     =   9480
   Icon            =   "frm_imp_orca.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00404080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9255
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000C&
         Caption         =   "Orçamentos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1935
         Left            =   2400
         TabIndex        =   12
         Top             =   2400
         Visible         =   0   'False
         Width           =   4335
         Begin VB.ListBox List2 
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
            Height          =   1410
            ItemData        =   "frm_imp_orca.frx":08CA
            Left            =   120
            List            =   "frm_imp_orca.frx":08CC
            TabIndex        =   13
            Top             =   360
            Width           =   4095
         End
      End
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
         Height          =   1605
         ItemData        =   "frm_imp_orca.frx":08CE
         Left            =   120
         List            =   "frm_imp_orca.frx":08D0
         TabIndex        =   11
         Top             =   2160
         Width           =   9015
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
         Left            =   240
         TabIndex        =   7
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
         Left            =   7080
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox nome 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   1680
         Width           =   7095
      End
      Begin VB.CommandButton Command1 
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
         Height          =   855
         Left            =   360
         Picture         =   "frm_imp_orca.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5520
         Width           =   1455
      End
      Begin VB.TextBox data 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   405
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   4920
         Width           =   4095
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
         Left            =   7320
         Picture         =   "frm_imp_orca.frx":0D14
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5520
         Width           =   1455
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
         Left            =   3960
         Picture         =   "frm_imp_orca.frx":1156
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5520
         Width           =   1455
      End
      Begin Crystal.CrystalReport IMPRESSORA 
         Left            =   8760
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   ".\imp_orcamento.rpt"
         WindowLeft      =   0
         WindowTop       =   0
         WindowWidth     =   0
         WindowHeight    =   0
         DiscardSavedData=   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404080&
         Caption         =   "Data de edição do orçamento"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   4920
         Width           =   3975
      End
      Begin VB.Label titulo 
         BackColor       =   &H00404080&
         Caption         =   "Nome do cliente:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   9240
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404080&
         Caption         =   "Orçamento No. :"
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
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Numero 
         BackColor       =   &H00404080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm_imp_orca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private tab_orca As Recordset
Private cod_orca As Long
Private cod As Long


Private Sub Command1_Click()
On Error GoTo arq_erro

IMPRESSORA.ReplaceSelectionFormula ("{orcamento.cod_orca} = " & cod_orca)

IMPRESSORA.ReportFileName = App.Path + "\imp_orcamento.rpt"

IMPRESSORA.Action = 0
arq_erro:
 If (Err.Number = 20513) Or (Err.Number = 20526) Then
      MsgBox "Impressora não esta pronta", , "Pesquisa"
  
  End If
End Sub

Private Sub Command2_Click()

Unload frm_imp_orca
End Sub

Private Sub Command3_Click()
Dim tab_aux As Recordset

cod = 0

'inicializa nova pesquisa
Frame2.Visible = False
Numero.Caption = ""
data.Text = ""
Command1.Enabled = False


nome.Text = ""
List1.Clear


If Option2.Value Then
     Set tab_aux = frm_principal.arquivo.OpenRecordset("select * from clientes order by nome", dbOpenDynaset)
     
     Do While Not tab_aux.EOF
            List1.AddItem tab_aux.Fields(0) & Space(15 - ((Len(tab_aux.Fields(0))) * 2)) & tab_aux.Fields(1) & " --- " & tab_aux.Fields(2)
            tab_aux.MoveNext
     Loop
     
Else
 
    Set tab_aux = frm_principal.arquivo.OpenRecordset("select * from clientes order by endereco", dbOpenDynaset)
    
    Do While Not tab_aux.EOF
            List1.AddItem tab_aux.Fields(0) & Space(15 - ((Len(tab_aux.Fields(0))) * 2)) & tab_aux.Fields(2) & " --- " & tab_aux.Fields(1)
            tab_aux.MoveNext
    Loop
    
End If

        
If List1.ListCount = 1 Then
    List1.ListIndex = 0
Else
    List1.Visible = True
End If

End Sub

Private Sub Form_Load()
Dim tab_aux As Recordset


Set tab_orca = frm_principal.arquivo.OpenRecordset("orcamento", dbOpenTable)

Set tab_aux = frm_principal.arquivo.OpenRecordset("select * from clientes  order by nome", dbOpenDynaset)

List1.Clear

Do While Not tab_aux.EOF
    List1.AddItem tab_aux.Fields(0) & Space(15 - ((Len(tab_aux.Fields(0))) * 2)) & tab_aux.Fields(1) & " --- " & tab_aux.Fields(2)
    tab_aux.MoveNext
Loop
If List1.ListCount = 1 Then
   List1.ListIndex = 0
Else
   List1.Visible = True
End If



End Sub

Private Sub Form_Unload(Cancel As Integer)
tab_orca.Close
frm_principal.Enabled = True

End Sub

Private Sub List1_Click()
Dim tab_aux As Recordset
cod = Trim(Mid(List1.Text, 1, 10))

Set tab_orca = frm_principal.arquivo.OpenRecordset("select * from orcamento where cod_cli = " & cod & " order by data", dbOpenDynaset)
Set tab_aux = frm_principal.arquivo.OpenRecordset("select * from clientes where cod =" & cod, dbOpenDynaset)


List2.Clear
Do While Not tab_orca.EOF
   List2.AddItem tab_orca.Fields(2) & Space(10) & "Orçamento No: " & tab_orca.Fields(0)
   tab_orca.MoveNext
Loop


If List2.ListCount = 0 Then
    MsgBox "Não existe orçamento para este cliente, escolha a opção 'CRIAR ORÇAMENTO' no menu principal", , "Pesquisa"
    Exit Sub
Else
   Frame2.Visible = True
End If
  

nome.Text = tab_aux.Fields(1)


List1.Visible = False

End Sub

Private Sub List2_Click()
Dim matriz(12) As String
Dim i As Integer
Dim k As String
Dim pesquisa As String
Dim mes As String
Dim dia As String
Dim ano As String
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

For i = 10 To Len(List2.Text)
   k = Mid(List2.Text, i, 1)
   If k = ":" Then
      cod_orca = Trim(Mid(List2.Text, i + 1, 15))
      Exit For
   End If
Next i
pesquisa = "cod_orca = " & cod_orca
tab_orca.FindFirst pesquisa

Numero.Caption = tab_orca.Fields(0)

mes = Val(Mid(tab_orca.Fields(2), 4, 2))
dia = Mid(tab_orca.Fields(2), 1, 2)

If Val(Mid(tab_orca.Fields(2), 7, 2)) > 35 Then
    ano = 19 & Mid(tab_orca.Fields(2), 7, 2)
Else
   ano = 20 & Mid(tab_orca.Fields(2), 7, 2)
End If

data.Text = "São Paulo, " & dia & " de " & matriz(mes) & " " & ano

Command1.Enabled = True

End Sub

Private Sub nome_Change()
Dim cad_cli As Recordset

If Option2.Value Then

Set cad_cli = frm_principal.arquivo.OpenRecordset("select * from clientes where nome like '" & nome.Text & "*'", dbOpenDynaset)
List1.Clear

Do While Not cad_cli.EOF
   List1.AddItem cad_cli.Fields(0) & Space(15 - ((Len(cad_cli.Fields(0))) * 2)) & cad_cli.Fields(1) & " --- " & cad_cli.Fields(2)
   cad_cli.MoveNext
Loop
   
   If List1.ListCount = 1 Then
      List1.ListIndex = 0
   End If
Else

Set cad_cli = frm_principal.arquivo.OpenRecordset("select * from clientes where endereco like '" & nome.Text & "*'", dbOpenDynaset)
List1.Clear

Do While Not cad_cli.EOF
   List1.AddItem cad_cli.Fields(0) & Space(15 - ((Len(cad_cli.Fields(0))) * 2)) & cad_cli.Fields(2) & " --- " & cad_cli.Fields(1)
   cad_cli.MoveNext
Loop
   
   If List1.ListCount = 1 Then
      List1.ListIndex = 0
   End If
End If
End Sub

Private Sub Option2_Click()
Call Command3_Click
End Sub

Private Sub Option3_Click()
Call Command3_Click
End Sub
