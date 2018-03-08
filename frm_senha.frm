VERSION 5.00
Begin VB.Form frm_senha 
   BackColor       =   &H00404080&
   Caption         =   "Alterando senha "
   ClientHeight    =   3840
   ClientLeft      =   8805
   ClientTop       =   4665
   ClientWidth     =   5505
   ControlBox      =   0   'False
   Icon            =   "frm_senha.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   5505
   Begin VB.CommandButton Command2 
      Caption         =   "Sair "
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
      Left            =   3360
      Picture         =   "frm_senha.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gravar "
      Enabled         =   0   'False
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
      Left            =   600
      Picture         =   "frm_senha.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000B&
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
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      Caption         =   "Confirme:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404080&
      Caption         =   "Digite a senha:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frm_senha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Dim registro As geral.record
Dim x As String
Open App.Path & "\ipsbit.dll" For Binary As #1

Line Input #1, x
Line Input #1, registro.disco
Line Input #1, registro.senha
Line Input #1, registro.nome_emp
Line Input #1, registro.data
Line Input #1, registro.mes_ano

registro.disco = Mid(registro.disco, 1, 9)
registro.senha = Mid(registro.senha, 1, 19)
registro.nome_emp = Mid(registro.nome_emp, 1, 29)
registro.data = Mid(registro.data, 1, 10)
registro.mes_ano = Mid(registro.mes_ano, 1, 10)

registro.senha = UCase(Text1.Text & Space(20 - Len(Text1.Text)))

Close #1
Kill App.Path & "\ipsbit.dll"

Open App.Path & "\ipsbit.dll" For Binary As #1

Put #1, 1, registro 'GRAVA REGISTRO
Close #1
frm_principal.Enabled = True
Unload frm_senha
End Sub

Private Sub Command2_Click()
frm_principal.Enabled = True
Unload frm_senha
End Sub


Private Sub Text1_Change()
Dim i As Integer

If Text1.Text <> "" Then
   Text2.Enabled = True
   If verifica_integridade() Then
      Command1.Enabled = True
      For i = 0 To 5
          Beep
      Next i
   Else
      Command1.Enabled = False
   End If
Else
   Command1.Enabled = False
   Text2.Enabled = False
End If
End Sub

Private Sub Text2_Change()
Dim i As Integer
If verifica_integridade() Then
   Command1.Enabled = True
   For i = 0 To 5
     Beep
   Next i
Else
   Command1.Enabled = False
End If


End Sub
Private Function verifica_integridade()
    If Text1.Text = Text2.Text Then
        verifica_integridade = True
    Else
        verifica_integridade = False
    End If
End Function

