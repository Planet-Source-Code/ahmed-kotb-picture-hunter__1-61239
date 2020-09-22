VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Picture Hunter"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4650
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin PictureHunter.chameleonButton command9 
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "About"
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
      MICON           =   "Form1.frx":2CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PictureHunter.chameleonButton command8 
      Height          =   495
      Left            =   2160
      TabIndex        =   12
      Top             =   4320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "Activate"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      MICON           =   "Form1.frx":2CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PictureHunter.chameleonButton command7 
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "View"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
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
      MICON           =   "Form1.frx":2D02
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PictureHunter.chameleonButton command6 
      Height          =   855
      Left            =   3720
      TabIndex        =   10
      Top             =   3960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "Quit"
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
      MICON           =   "Form1.frx":2D1E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PictureHunter.chameleonButton Command5 
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   3000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Settings"
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
      MICON           =   "Form1.frx":2D3A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PictureHunter.chameleonButton Command4 
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   2520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Save All"
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
      MICON           =   "Form1.frx":2D56
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
      Left            =   3720
      TabIndex        =   7
      Top             =   2040
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Clear All"
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
      MICON           =   "Form1.frx":2D72
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PictureHunter.chameleonButton Command2 
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
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
      MICON           =   "Form1.frx":2D8E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PictureHunter.chameleonButton command1 
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   1080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Delete"
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
      MICON           =   "Form1.frx":2DAA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0E0FF&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   3495
   End
   Begin VB.PictureBox pic 
      Height          =   375
      Index           =   0
      Left            =   1800
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1440
      Top             =   4320
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Auto save is off"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Picture Hunter By Ahmed Kotb"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "aaass"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Image imgview 
      Height          =   1935
      Left            =   480
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   120
      Picture         =   "Form1.frx":2DC6
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3525
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim picnum As Long
Dim nval As Long
Public path As String
Public auto As Integer
Public ind As Integer
Public st As Boolean
Public bt As Integer
Private Sub Command1_Click()
If Combo1.ListIndex = -1 Then Exit Sub
imgview.Picture = LoadPicture("")
Dim a(0) As String
a(0) = Combo1.ListIndex


If Combo1.ListIndex + 1 = picnum Then GoTo del

For i = Combo1.ListIndex + 1 + 1 To picnum
pic(i - 1).Picture = pic(i).Picture
Next i

Combo1.ListIndex = Combo1.ListIndex + 1
imgview.Picture = pic(Combo1.ListIndex).Picture

del:
Unload pic(picnum)
picnum = picnum - 1
Label1.Caption = "Number of Picture = " & picnum
Combo1.RemoveItem a(0)

Erase a
End Sub

Private Sub Combo1_Click()
imgview.Picture = pic(Combo1.ListIndex + 1).Picture
End Sub

Private Sub Command2_Click()
If imgview.Picture = 0 Then Exit Sub
Dim dist As String
Dim dlg1 As New dlg
dlg1.Filter = "Bitmaps (*.bmp)|*.bmp"
dlg1.DefaultExtension = "*.bmp"
dist = dlg1.FileSave
If dist <> "" Then SavePicture imgview.Picture, dist
End Sub

Private Sub Command3_Click()
For i = 1 To picnum
Unload pic(i)
Next i
imgview.Picture = LoadPicture("")
picnum = 0
Label1.Caption = "Number of Picture = " & picnum
Combo1.Clear
End Sub

Private Sub Command4_Click()
If picnum = 0 Then Exit Sub
Form3.Show vbModal
st = False
For i = 1 To picnum
SavePicture pic(i).Picture, path & "\Pic " & i & ".bmp"
DoEvents
If st = True Then Exit Sub
Next i
Unload Form3
End Sub

Private Sub Command5_Click()
Timer1.Enabled = False
Form2.Show vbModal
End Sub

Private Sub command6_Click()
MsgBox "Thank you 4 Using my program" & vbCrLf & "Ahmed Kotb", vbInformation, "Kotb Inc."
End
End Sub

Private Sub command7_Click()
imgview_Click
End Sub

Private Sub command8_Click()
Select Case command8.Caption
Case "Activate"
Clipboard.Clear
Timer1.Enabled = True
command8.Caption = "Stop"
Case "Stop"
Timer1.Enabled = False
command8.Caption = "Activate"
End Select
End Sub

Private Sub command9_Click()
frmAbout.Show
End Sub

Private Sub Form_Load()
Label1.Caption = "Number of Picture = " & picnum
Clipboard.Clear
auto = False
path = App.path & "\pics"
ind = 2
On Error Resume Next
MkDir App.path & "\pics"
Open App.path & "\config.dat" For Input As #1
Input #1, path, auto, ind, bt
Close
If bt = 0 Then bt = 3
If auto = 1 Then Label3.Caption = "Auto save is On"
Select Case ind
Case 0
Timer1.Interval = "2000"
Case 1
Timer1.Interval = "1000"
Case 2
Timer1.Interval = "500"
End Select
changebtskin bt + 1
command8_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Thank you 4 Using my program" & vbCrLf & "Ahmed Kotb", vbInformation, "Kotb Inc."
End

End Sub

Private Sub imgview_Click()
If imgview.Picture = 0 Then Exit Sub
Form4.Image1.Picture = imgview.Picture
Form4.Show vbModal
'Form4.Height = Form4.Image1.Picture.Height
'Form4.Width = Form4.Image1.Picture.Width
End Sub

Private Sub Timer1_Timer()
pic(0).Picture = Clipboard.GetData
If pic(0).Picture <> 0 Then
Clipboard.Clear
picnum = picnum + 1
Label1.Caption = "Number of Picture = " & picnum
Load pic(picnum)
pic(picnum).Picture = pic(0).Picture
If auto = 1 Then
SavePicture pic(0).Picture, path & "\" & picnum & "-" & Hour(Now) & "-" & Minute(Now) & "-" & Second(Now) & ".bmp"
End If
pic(0).Picture = LoadPicture("")
Combo1.AddItem Now
End If
End Sub
Public Function changebtskin(t As Integer)
command1.ButtonType = t
command2.ButtonType = t
Command3.ButtonType = t
command4.ButtonType = t
Command5.ButtonType = t
command6.ButtonType = t
command7.ButtonType = t
command8.ButtonType = t
command9.ButtonType = t
Form2.command1.ButtonType = t
Form2.command2.ButtonType = t
Form2.Command3.ButtonType = t
Form2.command4.ButtonType = t
Form1.Refresh
End Function
