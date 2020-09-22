VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   1665
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5685
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1149.213
   ScaleMode       =   0  'User
   ScaleWidth      =   5338.509
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   4320
      TabIndex        =   0
      Top             =   1080
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5183.565
      Y1              =   662.609
      Y2              =   662.609
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1050
      TabIndex        =   2
      Top             =   120
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   5183.565
      Y1              =   662.609
      Y2              =   662.609
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   3
      Top             =   660
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Kotb Corp. 2004                                                            All Rights Reseved To Ahmed Kotb"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub
Private Sub Form_Load()
Me.Icon = Form1.Icon
    Image1.Picture = Form1.Icon
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title & " By Ahmed Kotb"
End Sub
