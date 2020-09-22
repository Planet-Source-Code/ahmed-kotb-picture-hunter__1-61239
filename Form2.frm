VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4890
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin PictureHunter.chameleonButton command1 
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   840
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      BTYPE           =   2
      TX              =   "Browse"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form2.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PictureHunter.chameleonButton Command3 
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form2.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PictureHunter.chameleonButton command2 
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form2.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PictureHunter.chameleonButton command4 
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   2640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Sample"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form2.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0E0FF&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2640
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0E0FF&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1920
      Width           =   4695
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000080&
      Caption         =   "AutoSave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Command button Skin :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose program Sensitivity :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Default dir path :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Settings :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmr As Long
Dim s As Integer
Private Sub Combo2_Click()
command4.ButtonType = Combo2.ListIndex + 1
End Sub

Private Sub Command1_Click()
Dim p As String
p = Text1.text
Text1.text = showDir("Enter new folder path :")
If Text1.text = "" Then Text1.text = p
End Sub

Private Sub Command2_Click()

Open App.path & "\config.dat" For Output As #1
Write #1, Text1.text, Check1.Value, Combo1.ListIndex, Combo2.ListIndex
Close
Form1.path = Text1.text
Form1.auto = Check1.Value
Form1.Timer1.Interval = tmr
Form1.ind = Combo1.ListIndex
If Check1.Value = 1 Then Form1.Label3.Caption = "AutoSave is On"
If Check1.Value = 0 Then Form1.Label3.Caption = "AutoSave is Off"
Form1.Timer1.Enabled = True
If s <> Combo2.ListIndex Then MsgBox "the button skin will be changed when you startup the program", vbInformation, "Skin Changed"
Unload Me
End Sub

Private Sub Command3_Click()
Form1.Timer1.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
Combo1.AddItem "low -checks the clipboard every 2 seconds-"
Combo1.AddItem "mediam -checks the clipboard every 1 seconds"
Combo1.AddItem "High -checks the clipboard every 1/2 seconds"
For i = 1 To 14
Combo2.AddItem "Type " & i
Next i
Combo2.ListIndex = bt
On Error Resume Next
Text1.text = Form1.path
Check1.Value = Form1.auto
Combo1.ListIndex = Form1.ind
Combo2.ListIndex = Form1.bt
s = Combo2.ListIndex
Select Case Form1.ind
Case 0
tmr = 2000
Case 1
tmr = 1000
Case 3
tmr = 500
End Select

End Sub

