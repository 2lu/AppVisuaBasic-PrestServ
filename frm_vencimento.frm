VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_vencimento 
   BackColor       =   &H00404080&
   Caption         =   "Mala direta"
   ClientHeight    =   6615
   ClientLeft      =   6315
   ClientTop       =   2865
   ClientWidth     =   10350
   Icon            =   "frm_vencimento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   10350
   Begin VB.CommandButton Command4 
      Caption         =   "Atualizar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404080&
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   2895
      Begin VB.OptionButton ord_e 
         BackColor       =   &H00404080&
         Caption         =   "Ordenar por  ENDEREÇO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   2655
      End
      Begin VB.OptionButton ord_n 
         BackColor       =   &H00404080&
         Caption         =   "Ordenar por NOME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404080&
      Caption         =   "Clientes encontrados:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   3960
      TabIndex        =   11
      Top             =   1200
      Width           =   2055
      Begin VB.Label qt_CLI 
         Alignment       =   2  'Center
         BackColor       =   &H00404080&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
   End
   Begin Crystal.CrystalReport impressora2 
      Left            =   9960
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   ".\imp_lista.rpt"
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport impressora1 
      Left            =   9360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   ".\imp_etiq.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mala direta"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   1080
      Picture         =   "frm_vencimento.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404080&
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   10095
      Begin VB.TextBox lim 
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
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   9120
         MaxLength       =   2
         TabIndex        =   9
         Text            =   " 5"
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton venc 
         BackColor       =   &H00404080&
         Caption         =   "Clientes com certificado de garantia vencido no prazo  de          anos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   372
         Left            =   4080
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   5895
      End
      Begin VB.OptionButton todos 
         BackColor       =   &H00404080&
         Caption         =   "Todos os clientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.ListBox LISTA_CLI 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2580
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Para emitir uma etiqueta deste CLIENTE  clique duas vezes na linha"
      Top             =   2520
      Width           =   10335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7920
      Picture         =   "frm_vencimento.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Listagem"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   4680
      Picture         =   "frm_vencimento.frx":114E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00404080&
      Height          =   975
      Left            =   6960
      TabIndex        =   0
      Top             =   1200
      Width           =   3255
      Begin VB.CheckBox sem_contac 
         BackColor       =   &H00404080&
         Caption         =   "Listar clientes não contactados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox contac 
         BackColor       =   &H00404080&
         Caption         =   "Listar clientes ja contactados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frm_vencimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private db1 As Recordset


Private Sub Command1_Click(Index As Integer)

On Error GoTo arq_erro
'----- imprime as etiquetas
impressora1.ReportFileName = App.Path + "\IMP_ETIQ.rpt"

impressora1.Action = 0
'--------- registra data de envio de corresp na base cli ---
 LISTA_CLI.ListIndex = -1

Do
    LISTA_CLI.ListIndex = LISTA_CLI.ListIndex + 1
    db1.FindFirst "cod = " & Val(Mid(LISTA_CLI.Text, 1, 10))

    db1.Edit
        db1("data_mala") = Format(Date, "dd/mm/yyyy")
    db1.Update
    
    
Loop Until LISTA_CLI.ListIndex >= LISTA_CLI.ListCount - 1
'-----------------------------------
LISTA_CLI.ListIndex = 0 'posiciona a primieira linha
LISTA_CLI.ListIndex = -1 'da lista_cli

arq_erro:

Exit Sub

 If (Err.Number = 20513) Or (Err.Number = 20526) Or (Err.Number = 20504) Then
      MsgBox "Impressora não esta pronta", , "Pesquisa"
    Exit Sub
  End If
End Sub

Private Sub Command2_Click(Index As Integer)
On Error GoTo arq_erro
Dim IMPRESSORA As CrystalReport

'----- imprimi lista de clientes
IMPRESSORA.ReportFileName = App.Path + "\imp_lista.rpt"
impressora2.Action = 0
'------------------
Exit Sub
arq_erro:

 If (Err.Number = 20513) Or (Err.Number = 20526) Or (Err.Number = 20504) Then
      MsgBox "Impressora não esta pronta", , "Pesquisa"
   
  End If
End Sub

Private Sub Command3_Click()

Unload frm_vencimento
End Sub

Private Sub Command4_Click()
Call RERODAR
End Sub

Private Sub contac_Click()
Call RERODAR
End Sub

Private Sub Form_Load()

Call RERODAR
End Sub
Private Sub RERODAR()
Dim CAMPO As String * 10

CAMPO = IIf(ord_n, "NOME", "ENDERECO")
'-- logia para as opcoes da tela de vencimentos  /c atualiza online ----
'-- cria o arquivo de impressao na tab. ACESS e a um recordset com os reg. da lista_cli

If todos Then
    frm_principal.arquivo.Execute "DROP TABLE ETIQ_TEMP"
    Set db1 = frm_principal.arquivo.OpenRecordset("select cod,nome,endereco,telefone,DATA_SERV,DATA_MALA,bairro,cidade,cep from clientes order by nome", dbOpenDynaset)
    frm_principal.arquivo.Execute "select clientes.* into etiq_temp from clientes"
Else
    If contac And sem_contac Then
        frm_principal.arquivo.Execute "DROP TABLE ETIQ_TEMP"
        Set db1 = frm_principal.arquivo.OpenRecordset("select cod,nome,endereco,telefone,DATA_SERV,DATA_MALA,bairro,cidade,cep from clientes order by nome", dbOpenDynaset)
        frm_principal.arquivo.Execute "select clientes.* into etiq_temp from clientes ORDER BY " & CAMPO
    ElseIf contac Then
        frm_principal.arquivo.Execute "DROP TABLE ETIQ_TEMP"
        Set db1 = frm_principal.arquivo.OpenRecordset("select cod,nome,endereco,telefone,DATA_SERV,DATA_MALA,bairro,cidade,cep from clientes where data_mala <> null and DATA_MALA <> '----------' order by nome", dbOpenDynaset)
        frm_principal.arquivo.Execute "select clientes.* into etiq_temp from clientes where  DATA_MALA <> null and DATA_MALA <> '----------' ORDER BY " & CAMPO
    ElseIf sem_contac Then
        frm_principal.arquivo.Execute "DROP TABLE ETIQ_TEMP"
        Set db1 = frm_principal.arquivo.OpenRecordset("select cod,nome,endereco,telefone,DATA_SERV,DATA_MALA,bairro,cidade,cep  from clientes where data_mala = null OR DATA_MALA = '----------'  order by nome ", dbOpenDynaset)
        frm_principal.arquivo.Execute "select clientes.* into etiq_temp from clientes where data_mala = null OR DATA_MALA ='----------' ORDER BY " & CAMPO
    End If
 End If
 '---------------
   Call LISTAR_CLIENTES



End Sub
Private Sub LISTAR_CLIENTES()
'-- lista os clienete na lista_cli excluindo os casos menor q o LIM "
Dim QT As Integer
Dim pos As Integer
Dim db1 As Recordset
Set db1 = frm_principal.arquivo.OpenRecordset("etiq_temp", dbOpenDynaset)
LISTA_CLI.Clear

Do While Not db1.EOF
If ord_n Then
    pos = 1
Else
    pos = 2
End If

   
   If venc Then
   
       
       If DateDiff("yyyy", db1("DATA_SERV"), Date) >= Val(lim) Then
                 
          LISTA_CLI.AddItem db1("cod") & String(9 - (Len(LTrim(RTrim(Str(db1("cod")))))), " ") + db1.Fields(pos) + String(51 - (Len(LTrim(RTrim(db1.Fields(pos))))), ".") + db1.Fields(3 - pos) + String(52 - (Len(LTrim(RTrim(db1.Fields(3 - pos))))), ".") + db1("TELEFONE")
          QT = QT + 1  'contador
       Else
          db1.Delete
       End If
       
    Else
   
          LISTA_CLI.AddItem db1("cod") & String(9 - (Len(LTrim(RTrim(Str(db1("cod")))))), " ") + db1.Fields(pos) + String(51 - (Len(LTrim(RTrim(db1.Fields(pos))))), ".") + db1.Fields(3 - pos) + String(52 - (Len(LTrim(RTrim(db1.Fields(3 - pos))))), ".") + db1("TELEFONE")
          QT = QT + 1 'contador
    End If
 
    
    
    db1.MoveNext

    If db1.EOF Then
     Exit Do
    End If

Loop

If LISTA_CLI.ListCount > 0 Then
    Command2(0).Enabled = True
    Command1(0).Enabled = True
Else
    Command2(0).Enabled = False
    Command1(0).Enabled = False
End If

qt_CLI.Caption = QT
db1.Close
End Sub


Private Sub Form_Unload(Cancel As Integer)
db1.Close
End Sub

Private Sub LISTA_CLI_DblClick()
On Error GoTo arq_erro
Dim op As Integer

op = MsgBox("Confirma a impressão do cliente selecionado ?", vbYesNo, "Imprimir")
If op = 6 Then
'------- imprime etiqueta
    frm_principal.arquivo.Execute "drop table etiq_temp"
    frm_principal.arquivo.Execute "select clientes.* into etiq_temp from clientes where cod  = " & Val(Mid(LISTA_CLI.Text, 1, 10))
    impressora1.Action = 1
    
    '---------- registra data de envio de corresp na base cli ----
    
    db1.FindFirst "cod = " & Val(Mid(LISTA_CLI.Text, 1, 10))

    db1.Edit
        db1("data_mala") = Format(Date, "dd/mm/yyyy")
    db1.Update
   '--------------------------
   
End If

arq_erro:

 If (Err.Number = 20513) Or (Err.Number = 20526) Or (Err.Number = 20504) Then
      MsgBox "Impressora não esta pronta", , "Pesquisa"
     Exit Sub

  End If
End Sub



Private Sub ord_e_Click()
Call RERODAR
End Sub

Private Sub ord_n_Click()
Call RERODAR
End Sub



Private Sub sem_contac_Click()
Call RERODAR
End Sub

Private Sub todos_Click()

   Frame.Enabled = False


Call RERODAR
End Sub

Private Sub venc_Click()

If venc Then
   Frame.Enabled = True
   Call RERODAR
   
End If
End Sub
