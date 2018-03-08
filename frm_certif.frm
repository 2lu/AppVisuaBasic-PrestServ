VERSION 5.00
Begin VB.Form frm_ceritf 
   Caption         =   "Certificado de Garantia"
   ClientHeight    =   5475
   ClientLeft      =   1575
   ClientTop       =   1110
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   9480
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
      Picture         =   "frm_ceritf.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4560
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
      Picture         =   "frm_ceritf.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
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
      Left            =   120
      Picture         =   "frm_ceritf.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      ItemData        =   "frm_ceritf.frx":0CC6
      Left            =   120
      List            =   "frm_ceritf.frx":0CC8
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   9255
   End
   Begin VB.OptionButton Option2 
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
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton Option3 
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
      Height          =   255
      Left            =   7200
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox nome 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   5295
   End
   Begin VB.TextBox ender 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   8535
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9480
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label titulo 
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
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
End
Attribute VB_Name = "frm_ceritf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tab_aux, tab_certif As Recordset

Private cod As Long

Private Sub Command1_Click()
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



mes = Val(Mid(Date, 4, 2))
dia = Mid(Date, 1, 2)

If Val(Mid(Date, 7, 2)) > 0 Then
    ano = 19 & Mid(Date, 7, 2)
Else
   ano = 20 & Mid(data, 6, 2)
End If

data.Text = "São Paulo, " & dia & " de " & matriz(mes) & " " & ano

End Sub

Private Sub Command2_Click()
tab_aux.Close
Unload frm_certif

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
            List1.AddItem tab_aux.Fields(0) & Space(25 - ((Len(tab_aux.Fields(0))) * 2)) & tab_aux.Fields(2) & " --- " & tab_aux.Fields(1)
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
