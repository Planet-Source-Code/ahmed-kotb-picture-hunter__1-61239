VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "view"
   ClientHeight    =   3285
   ClientLeft      =   165
   ClientTop       =   780
   ClientWidth     =   4005
   LinkTopic       =   "Form4"
   ScaleHeight     =   3285
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1815
   End
   Begin VB.Menu note 
      Caption         =   "Note : You can stretch this window"
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Icon = Form1.Icon
End Sub

Private Sub Form_Resize()
Image1.Width = Me.Width - 100
Image1.Height = Me.Height - 700
End Sub
