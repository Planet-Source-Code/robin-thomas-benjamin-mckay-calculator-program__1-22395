VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Calculator"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   3555
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   3555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command47 
      Caption         =   ".0"
      Height          =   255
      Left            =   600
      TabIndex        =   73
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command57 
      Caption         =   "Len"
      Height          =   255
      Left            =   600
      TabIndex        =   72
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command56 
      Caption         =   "Oct"
      Height          =   255
      Left            =   0
      TabIndex        =   70
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command55 
      Caption         =   "Hex"
      Height          =   255
      Left            =   3000
      TabIndex        =   68
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Command54 
      Caption         =   "Mem"
      Height          =   255
      Left            =   3000
      TabIndex        =   66
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command53 
      Caption         =   "Sqr 3"
      Height          =   255
      Left            =   2400
      TabIndex        =   65
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Command52 
      Caption         =   "Sqr 2"
      Height          =   255
      Left            =   1800
      TabIndex        =   64
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Command51 
      Caption         =   "abc"
      Height          =   255
      Left            =   1200
      TabIndex        =   63
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Command50 
      Appearance      =   0  'Flat
      Caption         =   "Tan"
      Height          =   255
      Left            =   1800
      TabIndex        =   62
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command49 
      Caption         =   "Cos"
      Height          =   255
      Left            =   1200
      TabIndex        =   61
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command48 
      Caption         =   "Sin"
      Height          =   255
      Left            =   600
      TabIndex        =   60
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command44 
      Caption         =   "Hyp"
      Height          =   255
      Left            =   2400
      TabIndex        =   59
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command43 
      Caption         =   "Sgn"
      Height          =   255
      Left            =   0
      TabIndex        =   58
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Command42 
      Caption         =   "Rnd"
      Height          =   255
      Left            =   2400
      TabIndex        =   57
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command46 
      Caption         =   "Int"
      Height          =   255
      Left            =   3000
      TabIndex        =   56
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command45 
      Caption         =   "Abs"
      Height          =   255
      Left            =   600
      TabIndex        =   55
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Command41 
      Caption         =   "Dec"
      Height          =   255
      Left            =   0
      TabIndex        =   54
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command40 
      Caption         =   "x4"
      Height          =   255
      Left            =   3000
      TabIndex        =   53
      Top             =   2280
      Width           =   495
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   3120
      Top             =   0
   End
   Begin VB.CommandButton Command39 
      Caption         =   "Mod"
      Height          =   255
      Left            =   1800
      TabIndex        =   51
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command38 
      Caption         =   "1/x"
      Height          =   255
      Left            =   1200
      TabIndex        =   50
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "0."
      Height          =   255
      Left            =   0
      TabIndex        =   49
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command36 
      Caption         =   "CE"
      Height          =   255
      Left            =   3000
      TabIndex        =   48
      Top             =   1200
      Width           =   495
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   2760
      Top             =   0
   End
   Begin VB.CommandButton Command35 
      Caption         =   "+/-"
      Height          =   255
      Left            =   3000
      TabIndex        =   43
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Diam"
      Height          =   255
      Left            =   2400
      TabIndex        =   42
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Radi"
      Height          =   255
      Left            =   1800
      TabIndex        =   41
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command32 
      Caption         =   "n!"
      Height          =   255
      Left            =   1200
      TabIndex        =   40
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Not"
      Height          =   255
      Left            =   0
      TabIndex        =   39
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Command30 
      Caption         =   "log"
      Height          =   255
      Left            =   600
      TabIndex        =   37
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command29 
      Caption         =   "%"
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Exp"
      Height          =   255
      Left            =   1200
      TabIndex        =   34
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Command27 
      Caption         =   "sA"
      Height          =   255
      Left            =   600
      TabIndex        =   33
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Command26 
      Caption         =   "sP"
      Height          =   255
      Left            =   3000
      TabIndex        =   32
      Top             =   840
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3480
      Top             =   0
   End
   Begin VB.CommandButton Command25 
      Caption         =   "XoR"
      Height          =   255
      Left            =   2400
      TabIndex        =   30
      Top             =   2280
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   0
   End
   Begin VB.CommandButton Command24 
      Caption         =   "CTK"
      Height          =   255
      Left            =   1800
      TabIndex        =   28
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Command23 
      Caption         =   "cC"
      Height          =   255
      Left            =   2400
      TabIndex        =   27
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command22 
      Caption         =   "cA"
      Height          =   255
      Left            =   2400
      TabIndex        =   26
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Pi"
      Height          =   255
      Left            =   1200
      TabIndex        =   25
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command20 
      Caption         =   "x3"
      Height          =   255
      Left            =   3000
      TabIndex        =   24
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command19 
      Caption         =   "x2"
      Height          =   255
      Left            =   3000
      TabIndex        =   23
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Sqrt"
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command17 
      Caption         =   "C"
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command16 
      Caption         =   "."
      Height          =   255
      Left            =   1200
      TabIndex        =   20
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      Caption         =   "0"
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Caption         =   "/"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Caption         =   "="
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      Caption         =   "X"
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "-"
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "+"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label17 
      Caption         =   "0"
      Height          =   495
      Left            =   1800
      TabIndex        =   74
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "0"
      Height          =   495
      Left            =   1320
      TabIndex        =   71
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "0"
      Height          =   495
      Left            =   1200
      TabIndex        =   69
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "0"
      Height          =   495
      Left            =   1200
      TabIndex        =   67
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "0"
      Height          =   495
      Left            =   840
      TabIndex        =   52
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "0"
      Height          =   495
      Left            =   1560
      TabIndex        =   47
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "0"
      Height          =   495
      Left            =   840
      TabIndex        =   46
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "0"
      Height          =   495
      Left            =   840
      TabIndex        =   45
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "0"
      Height          =   495
      Left            =   840
      TabIndex        =   44
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "1"
      Height          =   495
      Left            =   840
      TabIndex        =   38
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
      Height          =   495
      Left            =   720
      TabIndex        =   36
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2520
      TabIndex        =   31
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "The computation has been added to memory"
      Height          =   375
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "The Answer Is:"
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   495
      Left            =   1200
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Clear 
         Caption         =   "&Clear"
         Shortcut        =   {F6}
      End
      Begin VB.Menu border2 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Copy 
         Caption         =   "&Copy"
         Begin VB.Menu Number 
            Caption         =   "Number"
         End
         Begin VB.Menu Answer 
            Caption         =   "&Answer"
         End
      End
      Begin VB.Menu Paste 
         Caption         =   "&Paste"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu HelpTopics 
         Caption         =   "&Help Topics"
      End
      Begin VB.Menu border1 
         Caption         =   "-"
      End
      Begin VB.Menu AboutCalculator 
         Caption         =   "&About Calculator"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Calculator -- Written By Robin McKay
Const Pi As Variant = "3.1415926535897932384626433832795"

Private Sub AboutCalculator_Click()
Form2.Show
End Sub

Private Sub Answer_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText Label3.Caption
End Sub

Private Sub Clear_Click()
Call Command17_Click
End Sub

Private Sub Command1_Click()
Label17 = 0
Text1.SelText = CDec(1)
End Sub

Private Sub Command10_Click()
On Error Resume Next
Label17 = 0
If Text1 = Label3 Then
    Text1 = ""
    Label1 = 0
ElseIf Label1 = 0 Then
    Label1 = 1
    Label2 = 0 + Label2 + Text1
    Text1 = ""
    Label1 = 0
ElseIf Label1 = 1 Then
    Label1 = 0
    Label2 = Label2 - Text1
    Text1 = ""
ElseIf Label1 = 2 Then
    Label2 = Label2 * Text1
    Label1 = 0
    Text1 = ""
ElseIf Label1 = 3 Then
    Label2 = Label2 / Text1
    Label1 = 0
    Text1 = ""
End If
End Sub

Private Sub Command11_Click()
On Error Resume Next
Label17 = 0
If Text1 = Label3 Then
    Label1 = 1
    Text1 = ""
ElseIf Label1 = 0 Then
    Label2 = 0 + Label2 + Text1
    Label1 = 1
    Text1 = ""
    Label1 = 1
ElseIf Label1 = 2 Then
    Label2 = Label2 * Text1
    Label1 = 1
    Text1 = ""
ElseIf Label1 = 1 Then
    Label2 = Label2 - Text1
    Text1 = ""
ElseIf Label1 = 3 Then
    Label2 = Label2 / Text1
    Label1 = 1
    Text1 = ""
End If
End Sub

Private Sub Command12_Click()
On Error Resume Next
Label17 = 0
If Text1 = Label3 Then
    Label1 = 2
    Text1 = ""
ElseIf Label1 = 0 Then
    Label2 = 0 + Label2 + Text1
    Label1 = 2
    Text1 = ""
ElseIf Label1 = 1 Then
    Label2 = Label2 - Text1
    Label1 = 2
    Text1 = ""
ElseIf Label1 = 2 Then
    Label2 = Label2 * Text1
    Text1 = ""
ElseIf Label1 = 3 Then
    Label2 = Label2 / Text1
    Label1 = 2
    Text1 = ""
End If
End Sub

Private Sub Command13_Click()
On Error Resume Next
Dim i As Integer
If Label17 = 1 Then
    Label2 = Label2 + 0
ElseIf Label1 = 0 Then
    Label2 = 0 + Label2 + Text1
    Label3 = Label2
    Label1 = 0
    Label17 = 1
    Text1 = Label3
ElseIf Label1 = 1 Then
    Label2 = Label2 - Text1
    Label3 = Label2
    Label1 = 0
    Text1 = Label3
    Label17 = 1
ElseIf Label1 = 2 Then
    Label2 = Label2 * Text1
    Label3 = Label2
    Label1 = 0
    Text1 = Label3
    Label17 = 1
ElseIf Label1 = 3 Then
    Label2 = CDec(Label2) / CDec(Text1)
    Label3 = Label2
    Label1 = 0
    Text1 = Label3
    Label17 = 1
End If
Label8 = 1
End Sub

Private Sub Command14_Click()
On Error Resume Next
Label17 = 0
If Text1 = Label3 Then
    Label1 = 3
    Text1 = ""
ElseIf Label1 = 0 Then
    Label2 = 0 + Label2 + Text1
    Label1 = 3
    Text1 = ""
ElseIf Label1 = 1 Then
    Label2 = Label2 - Text1
    Label1 = 3
    Text1 = ""
ElseIf Label1 = 2 Then
    Label2 = Label2 * Text1
    Label1 = 3
    Text1 = ""
ElseIf Label1 = 3 Then
    Label2 = Label2 / Text1
    Text1 = ""
End If
End Sub

Private Sub Command15_Click()
' Handles division by Zero
Label17 = 0
Text1.SelText = 0
If (Label1 = 3) And (Text1 = 0) Then
    Label5.Caption = "Cannot divide by Zero"
    Label5.Visible = True
    Timer3.Enabled = True
    Label12 = 1
End If
End Sub

Private Sub Command16_Click()
Label17 = 0
Text1.SelText = "."
End Sub

Private Sub Command17_Click()
On Error Resume Next
If Label12 = 1 Then
    Label1 = 0
    Label3 = 0
    Label2 = 0
    Label17 = 0
    Label10 = 0
    Text1 = ""
ElseIf Label12 = 0 Then
    Label1 = 0
    Label2 = 0
    Label3 = 0
    Label17 = 0
    Label10 = 0
    Text1 = ""
End If
End Sub

Private Sub Command18_Click()
On Error Resume Next
Label17 = 0
Text1 = Sqr(Text1)
End Sub

Private Sub Command19_Click()
Label17 = 0
Text1 = Text1 * Text1
End Sub

Private Sub Command2_Click()
Label17 = 0
Text1.SelText = CDec(2)
End Sub

Private Sub Command20_Click()
Label17 = 0
Text1 = Text1 * Text1 * Text1
End Sub

Private Sub Command21_Click()
On Error Resume Next
Label17 = 0
Text1.Text = Pi
End Sub



Private Sub Command22_Click()
On Error Resume Next
Label17 = 0
Text1 = Pi * Text1 * Pi * Text1
End Sub

Private Sub Command23_Click()
On Error Resume Next
Label17 = 0
Text1 = 2 * Pi * Text1
End Sub

Private Sub Command24_Click()
On Error Resume Next
Label17 = 0
Timer1.Enabled = True
Label5.Caption = "The computation has been added to Memory"
Label5.Visible = True
Text1 = Text1 * 1.6
End Sub

Private Sub Command25_Click()
On Error GoTo cancelerr:
Dim i As Integer
Dim i2 As Integer
Dim i3 As Integer
Label17 = 0
    i = InputBox("Please enter the highest computation value:", "Highest")
    i2 = InputBox("Please enter the lowest computation value:", "Lowest")
    i3 = i - i2
    Label5.Caption = "The XOR is equal to:"
    Label5.Visible = True
    Label6.Visible = True
    Label6 = i3
    Timer2.Enabled = True
cancelerr:
Exit Sub
End Sub



Private Sub Command26_Click()
Label17 = 0
Text1 = 0 + Text1 + Text1 + Text1 + Text1
End Sub

Private Sub Command27_Click()
Label17 = 0
Text1 = Text1 * Text1
End Sub

Private Sub Command28_Click()
On Error Resume Next
Label17 = 0
Text1.SelText = ".e" + "+" + "0"
End Sub

Private Sub Command29_Click()
On Error Resume Next
' This performs a percentage calculation
Label17 = 0
Text1 = Text1 / 100
Text1 = Text1 * 100
Text1.Text = Text1.Text + "%"
End Sub

Private Sub Command3_Click()
Label17 = 0
Text1.SelText = CDec(3)
End Sub
Private Sub Command30_Click()
On Error Resume Next
Label17 = 0
log10 = Log(Text1) / Log(10)
Text1 = log10
End Sub

Private Sub Command31_Click()
On Error Resume Next
Label17 = 0
If Label10 = 0 Then
    Label10 = Text1.Text
    Label11 = 1
    Text1.Text = "-" + Text1.Text - 1
ElseIf Label11 = 1 Then
    Text1.Text = Label10
    Label11 = 0
    Label10 = 0
End If
End Sub

Private Sub Command32_Click()
' This program supports factorial 10
Label17 = 0
If Text1 = 1 Then Text1 = 1
If Text1 = 2 Then Text1 = 2
If Text1 = 3 Then Text1 = 6
If Text1 = 4 Then Text1 = 24
If Text1 = 5 Then Text1 = 120
If Text1 = 6 Then Text1 = 720
If Text1 = 7 Then Text1 = 5040
If Text1 = 8 Then Text1 = 40320
If Text1 = 9 Then Text1 = 362880
If Text1 = 10 Then Text1 = 3628800
End Sub

Private Sub Command33_Click()
On Error Resume Next
Label17 = 0
Text1 = Text1 / 2
End Sub

Private Sub Command34_Click()
On Error Resume Next
Label17 = 0
Text1 = Text1 * 2
End Sub





Private Sub Command35_Click()
' Use this to toggle plus and minus
On Error Resume Next
Label17 = 0
If Label9.Caption = 0 Then
    Text1 = Text1 - Text1 - Text1
    Label9 = 1
ElseIf Label9 = 1 Then
    Text1 = Text1 - Text1 - Text1
    Label9 = 0
End If
End Sub

Private Sub Command36_Click()
On Error Resume Next
Label17 = 0
Text1 = ""
End Sub

Private Sub Command37_Click()
On Error Resume Next
Label17 = 0
Text1.SelText = "0."
End Sub



Private Sub Command38_Click()
On Error Resume Next
Label17 = 0
Text1 = Text1 / Text1 / Text1
End Sub



Private Sub Command39_Click()
On Error Resume Next
Dim i As Integer
Dim i2 As Integer
Dim i3 As Integer
Label17 = 0
Label13 = 0
    i = InputBox("Please enter the 1st number:", "Number 1")
    i2 = InputBox("Please enter the 2nd number:", "Number 2")
    i3 = i Mod i2
    Label13 = i3
    Label5.Caption = "The MOD is " + Label13
    Label5.Visible = True
    Label13 = ""
    Timer4.Enabled = True
End Sub

Private Sub Command4_Click()
Label17 = 0
Text1.SelText = CDec(4)
End Sub



Private Sub Command40_Click()
On Error Resume Next
Label17 = 0
Text1 = Text1 * Text1 * Text1 * Text1
End Sub

Private Sub Command41_Click()
' A procedure to convert a number to decimal
On Error Resume Next
Label17 = 0
Text1 = Text1 / 100
End Sub



Private Sub Command42_Click()
On Error Resume Next
Randomize
Text1 = Rnd(Text1)
Label17 = 0
End Sub

Private Sub Command43_Click()
On Error Resume Next
Label17 = 0
Text1 = Sgn(Text1)
End Sub

Private Sub Command44_Click()
On Error GoTo cancelerr:
Dim i As Integer
Dim i2 As Integer
Dim i3 As Integer
Dim i4 As Integer
Dim i5 As Integer
Label17 = 0
    i = InputBox("Please enter number:", "1st Number")
    i2 = InputBox("Please enter number:", "2nd Number")
    i3 = i * i
    i4 = i2 * i2
    i5 = i3 + i4
    Text1 = Sqr(i5)
cancelerr:
Exit Sub
End Sub

Private Sub Command45_Click()
On Error Resume Next
Label17 = 0
Text1 = Abs(Text1)
End Sub

Private Sub Command46_Click()
On Error Resume Next
Label17 = 0
Text1 = Int(Text1)
End Sub














Private Sub Command47_Click()
' This avoids the division by zero
' error.
On Error Resume Next
Label17 = 0
Text1.SelText = "0"
End Sub

Private Sub Command48_Click()
' This works out the sin of a number
On Error Resume Next
Label17 = 0
Text1 = Text1 * Pi / 180
Text1 = Sin(Text1)
End Sub

Private Sub Command49_Click()
' This works out the cosine of a number
On Error Resume Next
Label17 = 0
Text1 = Text1 * Pi / 180
Text1 = Cos(Text1)
End Sub

Private Sub Command5_Click()
Label17 = 0
Text1.SelText = CDec(5)
End Sub



Private Sub Command50_Click()
' This works out the tangent of a number
On Error Resume Next
Label17 = 0
Text1 = Text1 * Pi / 180
Text1 = Tan(Text1)
End Sub

Private Sub Command51_Click()
Form3.Show
End Sub

Private Sub Command52_Click()
' This function takes the square of the number originally squared and squares it again
On Error Resume Next
Label17 = 0
Text1 = Sqr(Text1)
Text1 = Sqr(Text1)
End Sub

Private Sub Command53_Click()
On Error Resume Next
Label17 = 0
Text1 = Sqr(Text1)
Text1 = Sqr(Text1)
Text1 = Sqr(Text1)
End Sub

Private Sub Command54_Click()
' This will store a number in temporary
' memory so that it can be retrieved at
' a later time.
On Error Resume Next
Label17 = 0
If Label15.Caption = 0 Then
    Label15.Caption = 1
    Label14.Caption = Text1
Else
If Label15.Caption = 1 Then
    Text1.Text = Label14.Caption
    Label15.Caption = 0
End If
End If
End Sub

Private Sub Command55_Click()
On Error Resume Next
Label15 = Hex(Text1)
Text1 = Label15
End Sub

Private Sub Command56_Click()
On Error Resume Next
Label16 = Oct(Text1)
Text1 = Label16
End Sub

Private Sub Command57_Click()
On Error Resume Next
Text1 = Len(Text1)
End Sub

Private Sub Command6_Click()
Label17 = 0
Text1.SelText = CDec(6)
End Sub

Private Sub Command7_Click()
Label17 = 0
Text1.SelText = CDec(7)
End Sub

Private Sub Command8_Click()
Label17 = 0
Text1.SelText = CDec(8)
End Sub

Private Sub Command9_Click()
Label17 = 0
Text1.SelText = CDec(9)
End Sub

Private Sub Degrees_Click()
On Error Resume Next
If Degrees.Checked = False Then
    Degrees.Checked = True
    Radians.Checked = False
End If
End Sub

Private Sub Exit_Click()
On Error Resume Next
Unload Form1
Unload Form2
Set Form1 = Nothing
Set Form2 = Nothing
End Sub

Private Sub Form_Load()
On Error Resume Next
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Label1.Visible = True
Label6.Visible = False
Radians.Checked = True
End Sub

Private Sub HelpTopics_Click()
MsgBox "Some functions use Input Boxes now. I have cleaned up the code slightly and added Sin, Cos and Tan. If you do Sin, Cos or Tan on some small numbers, you may see a minus 2 at the end. This means that you take your answer and move the decimal place 2 places to the left." + vbCrLf + "When you work out fractions, after you have entered all the values, press the + button to add the fractions together, - to take them away, X to times them together or / to divide them. All fraction answers are not given in their simplest form, but still give the correct answer.", vbInformation
End Sub

Private Sub Number_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText Text1.Text
End Sub

Private Sub Paste_Click()
On Error Resume Next
Text1 = Clipboard.GetText
End Sub

Private Sub Radians_Click()
On Error Resume Next
If Degrees.Checked = False Then
    Degrees.Checked = False
    Radians.Checked = True
End If
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Label5.Visible = False
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Label5.Visible = False
Label6.Visible = False
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
Label5.Visible = False
Call Clear_Click
End Sub

Private Sub Timer4_Timer()
Timer4.Enabled = False
Label5.Visible = False
End Sub
