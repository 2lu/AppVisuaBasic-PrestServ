VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_recibo 
   BackColor       =   &H00404080&
   Caption         =   "Imprimindo recibo"
   ClientHeight    =   6450
   ClientLeft      =   6480
   ClientTop       =   3405
   ClientWidth     =   9480
   Icon            =   "frm_recibo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9480
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
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   9255
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
         Height          =   1590
         Left            =   240
         TabIndex        =   19
         Top             =   1920
         Width           =   8535
      End
      Begin VB.TextBox valor 
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
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   13
         Top             =   960
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
         Picture         =   "frm_recibo.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5040
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
         Left            =   7560
         Picture         =   "frm_recibo.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5040
         Width           =   1215
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
         ForeColor       =   &H80000000&
         Height          =   405
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   4560
         Width           =   4095
      End
      Begin VB.TextBox ativ_2 
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
         MaxLength       =   66
         TabIndex        =   9
         Top             =   3960
         Width           =   8535
      End
      Begin VB.TextBox ativ_1 
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
         MaxLength       =   66
         TabIndex        =   8
         Top             =   3720
         Width           =   8535
      End
      Begin VB.TextBox import_2 
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
         LinkTimeout     =   30
         MaxLength       =   24
         TabIndex        =   7
         Top             =   3120
         Width           =   3255
      End
      Begin VB.TextBox import_1 
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
         MaxLength       =   66
         TabIndex        =   6
         Top             =   2880
         Width           =   8535
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
         TabIndex        =   5
         Top             =   2280
         Width           =   8535
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
         Picture         =   "frm_recibo.frx":114E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5040
         Width           =   1335
      End
      Begin VB.TextBox nome 
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
         TabIndex        =   3
         Top             =   1560
         Width           =   8535
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
         TabIndex        =   2
         Top             =   360
         Width           =   1695
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
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin Crystal.CrystalReport IMPRESSORA 
         Left            =   8760
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   "imp_recibo.rpt"
         WindowLeft      =   0
         WindowTop       =   0
         WindowWidth     =   0
         WindowHeight    =   0
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         DiscardSavedData=   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404080&
         Caption         =   "Referente a atividade:"
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
         TabIndex        =   18
         Top             =   3480
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404080&
         Caption         =   "Importância de:"
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
         TabIndex        =   17
         Top             =   2640
         Width           =   1215
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
         TabIndex        =   16
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404080&
         Caption         =   "Valor:   R$"
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
         Left            =   6360
         TabIndex        =   15
         Top             =   1080
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   9240
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
         TabIndex        =   14
         Top             =   1320
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frm_recibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public arquivo As Database
Public area As Workspace
Private tab_aux, tab_imp As Recordset
Private cod As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Command1_Click()
On Error GoTo arq_erro

'grava informacoes na table de impressao
Set tab_imp = frm_principal.arquivo.OpenRecordset("imp_recibo", dbOpenTable)

 tab_imp.AddNew
    'tab_imp("numero") = Numero.Caption
    tab_imp("nome") = nome.Text
    tab_imp("endereco") = ender.Text
    tab_imp("data") = data.Text
    tab_imp("valor") = valor.Text
    tab_imp("importancia1") = import_1.Text
    tab_imp("importancia2") = import_2.Text
    tab_imp("atividade1") = ativ_1
    tab_imp("atividade2") = ativ_2
 tab_imp.Update



    DoEvents
    Sleep (555)

IMPRESSORA.ReportFileName = App.Path + "\imp_recibo.rpt"


IMPRESSORA.Action = 0

tab_imp.MoveFirst 'deleta os dois registros
tab_imp.Delete
tab_imp.Close
Exit Sub

arq_erro:

  If (Err.Number = 20513) Or (Err.Number = 20526) Then
     MsgBox "Impressora não esta pronta", , "Pesquisa"
    
      'Resume Next
  
  End If
  
End Sub



Private Sub Command2_Click()

Unload frm_recibo


End Sub

Private Sub Command3_Click()


cod = 0
If Option2.Value Then
   nome.Text = ""
   List1.Top = 1920
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
   List1.Top = 2640
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






Private Sub Form_Unload(Cancel As Integer)
tab_aux.Close
frm_principal.Enabled = True
End Sub

Private Sub List1_Click()
cod = Trim(Mid(List1.Text, 1, 10))

Set tab_aux = frm_principal.arquivo.OpenRecordset("select * from clientes where cod = " & cod, dbOpenDynaset)


nome.Text = tab_aux.Fields(1)
ender.Text = tab_aux.Fields(2)
ender.Text = ender.Text & "  -  " & tab_aux.Fields(5)
ativ_1.Text = ""
ativ_2.Text = ""
import_1.Text = ""
import_2.Text = ""

valor.Text = ""


List1.Visible = False
End Sub

Private Sub Form_Load()
Dim mes As String
Dim ano As String
Dim dia As String

Frame1.Visible = True


Dim matriz(12) As String

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

mes = Val(Mid(Date, 4, 2))
dia = Mid(Date, 1, 2)

If Val(Mid(Date, Len(Str(Date)) - 1, 2)) > 35 Then
    ano = 19 & Mid(Date, Len(Str(Date)) - 1, 2)
Else
   ano = 20 & Mid(Date, Len(Str(Date)) - 1, 2)
End If

data.Text = "São Paulo, " & dia & " de " & matriz(mes) & " " & ano





End Sub

Private Sub mnu_sair_Click()

End
End Sub

Private Sub mnu_senha_Click()
frm_senha.Show
frm_principal.Enabled = False
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



Private Sub valor_LostFocus()
valor.Text = Format(valor.Text, "###,###,###.00")
End Sub

