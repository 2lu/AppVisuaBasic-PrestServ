VERSION 5.00
Begin VB.Form frm_cad_emp_tel 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alterando dados do estabelecimento"
   ClientHeight    =   3720
   ClientLeft      =   3045
   ClientTop       =   3405
   ClientWidth     =   7140
   FillColor       =   &H80000006&
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   11.25
      Charset         =   0
      Weight          =   900
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000A&
   Icon            =   "frm_cad_emp_tel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000010&
      Caption         =   "Alterar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Picture         =   "frm_cad_emp_tel.frx":17B2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6855
      Begin VB.TextBox TEL 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   14
         TabIndex        =   1
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox EMP 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   35
         TabIndex        =   0
         Top             =   600
         Width           =   6615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404080&
         Caption         =   "Nome fantasia:"
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
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404080&
         Caption         =   "Telefone (Exemplo: 99-99999999):"
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
         TabIndex        =   5
         Top             =   1320
         Width           =   4095
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000010&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      Picture         =   "frm_cad_emp_tel.frx":1BF4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
End
Attribute VB_Name = "frm_cad_emp_tel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private ARMAZENA_EMP_TEL As Recordset
            
            

Private Sub Command2_Click()
Unload frm_cad_emp_tel

End Sub



Private Sub Command3_Click()
    Call grava_tel_emp(EMP.Text, TEL.Text)
    MsgBox " Dados alterados com sucesso.", vbInformation, "Atenção"
    Unload frm_cad_emp_tel
End Sub

Private Sub EMP_Change()
If EMP.Text = "" Then
    Command3.Enabled = False
Else
    Command3.Enabled = True
End If

End Sub

Private Sub Form_Load()
Set ARMAZENA_EMP_TEL = frm_principal.arquivo.OpenRecordset("RECIBO", dbOpenTable)
EMP.Text = ARMAZENA_EMP_TEL("EMP")
TEL.Text = ARMAZENA_EMP_TEL("TEL")
End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_principal.Enabled = True
ARMAZENA_EMP_TEL.Close
End Sub

Sub grava_tel_emp(EMP As String, TEL As String)

            ARMAZENA_EMP_TEL.Edit
                ARMAZENA_EMP_TEL("EMP") = EMP
                ARMAZENA_EMP_TEL("TEL") = TEL
            ARMAZENA_EMP_TEL.Update

            
End Sub

