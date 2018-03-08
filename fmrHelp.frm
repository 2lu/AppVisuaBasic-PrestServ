VERSION 5.00
Begin VB.Form fmrHelp 
   Caption         =   "Help"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   12240
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox AcroPDF1 
      DragMode        =   1  'Automatic
      Height          =   12615
      Left            =   0
      ScaleHeight     =   12555
      ScaleWidth      =   12195
      TabIndex        =   0
      Top             =   0
      Width           =   12255
   End
End
Attribute VB_Name = "fmrHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
'AcroPDF1.src = App.Path & "\PrestServ.pdf"
Shell App.Path & "\PrestServ.pdf"

End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_principal.Enabled = True
End Sub
