VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   18
      Top             =   0
      Width           =   4575
   End
   Begin VB.CommandButton Command18 
      Caption         =   "E"
      Height          =   735
      Left            =   4200
      TabIndex        =   17
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      Caption         =   "C"
      Height          =   1215
      Left            =   4200
      TabIndex        =   16
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "/"
      Height          =   975
      Left            =   4200
      TabIndex        =   15
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "+"
      Height          =   1455
      Left            =   3240
      TabIndex        =   14
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      Caption         =   "-"
      Height          =   615
      Left            =   3240
      TabIndex        =   13
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      Caption         =   "*"
      Height          =   615
      Left            =   3240
      TabIndex        =   12
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      Caption         =   "="
      Height          =   615
      Left            =   2280
      TabIndex        =   11
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command11 
      Caption         =   "0"
      Height          =   615
      Left            =   1320
      TabIndex        =   10
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "."
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   615
      Left            =   2280
      TabIndex        =   8
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   615
      Left            =   1320
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   615
      Left            =   2280
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "4"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "5"
      Height          =   615
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Op, fn As Integer

Dim Op, fn As Integer

Private Sub Command1_Click()
Text1.Text = Text1.Text & 1
End Sub

Private Sub Command10_Click()
Text1.Text = Text1.Text & "."
End Sub

Private Sub Command11_Click()
Text1.Text = Text1.Text & 0
End Sub

Private Sub Command12_Click()
If Op = 1 Then
Text1.Text = Val(fn) + Val(Text1.Text)
ElseIf Op = 2 Then
Text1.Text = Val(fn) - Val(Text1.Text)
ElseIf Op = 3 Then
Text1.Text Val(fn) * Val(Text1.Text)
ElseIf Op = 4 Then
Text1.Text Val(fn) / Val(Text1.Text)
End If
End Sub

Private Sub Command13_Click()
Op = 3
fn = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command14_Click()
Op = 2
fn = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command15_Click()
Op = 1
fn = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command16_Click()
Op = 4
fn = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command17_Click()
Text1.Text = ""
End Sub

Private Sub Command18_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Text1.Text = Text1.Text & 2
End Sub

Private Sub Command3_Click()
Text1.Text = Text1.Text & 3
End Sub

Private Sub Command4_Click()
Text1.Text = Text1.Text & 4
End Sub

Private Sub Command5_Click()
Text1.Text = Text1.Text & 5
End Sub

Private Sub Command6_Click()
Text1.Text = Text1.Text & 6
End Sub

Private Sub Command7_Click()
Text1.Text = Text1.Text & 7
End Sub

Private Sub Command8_Click()
Text1.Text = Text1.Text & 8
End Sub

Private Sub Command9_Click()
Text1.Text = Text1.Text & 9
End Sub
