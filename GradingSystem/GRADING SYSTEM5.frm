VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H0080FF80&
   Caption         =   "Form5"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12960
   LinkTopic       =   "Form5"
   ScaleHeight     =   9045
   ScaleWidth      =   12960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Prelim Grade"
      Height          =   3495
      Left            =   2520
      TabIndex        =   31
      Top             =   5400
      Width           =   8415
      Begin VB.CommandButton Command3 
         Caption         =   "SUMMARY"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4680
         TabIndex        =   34
         Top             =   1440
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "COMPUTE FINAL GRADE"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4680
         TabIndex        =   33
         Top             =   360
         Width           =   3135
      End
      Begin VB.CommandButton Command5 
         Caption         =   "BACK"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4680
         TabIndex        =   32
         Top             =   2400
         Width           =   3135
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   44
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   43
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   42
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   41
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   40
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Equivalent"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   39
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Final Grade"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   38
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Class Standing"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Class Average"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   35
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Assignment"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   8520
      TabIndex        =   24
      Top             =   120
      Width           =   3975
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   47
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   26
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ASSIGNMENT #3"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   46
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   30
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Assignment Rating:"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ASSIGNMENT #2"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ASSIGNMENT #1"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Prelim Exam"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4320
      TabIndex        =   17
      Top             =   3000
      Width           =   8175
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   19
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   18
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   23
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Final Exam RATING"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Final Exam SCORE"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   " # of Items"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Recitation"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   4320
      TabIndex        =   10
      Top             =   120
      Width           =   3975
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   48
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RECITATION #3"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   45
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Recitation Rating:"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RECITATION #2"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RECITATION #1"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Quizzes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "&COMPUTE"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   4
         Top             =   4080
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1800
         TabIndex        =   3
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1800
         TabIndex        =   2
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1800
         TabIndex        =   9
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RATING:"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "QUIZ #3"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "QUIZ #2"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "QUIZ #1"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim q1, q2, q3, qrate As Integer
q1 = Val(Text2.Text)
q2 = Val(Text3.Text)
q3 = Val(Text4.Text)
qrate = (q1 + q2 + q3) / 3
Label6.Caption = qrate
Text5.SetFocus
End Sub
Private Sub Command2_Click()
Dim cs, fg, ca As Double
Dim rating, assrating, recrate As Integer
rating = Val(Label6.Caption)
assrating = Val(Label26.Caption)
recrate = Val(Label24.Caption)
cs = (rating + recrate + assrating) / 3
Label16.Caption = cs
ca = (cs * 2 + Label20.Caption) / 3
Label27.Caption = ca
fg = (ca * 2 + Form4.Label17.Caption) / 3
Label17.Caption = fg

Dim remarks As String
remarks = Val(Label17.Caption)
Select Case remarks
Case Is <= 100
Label19.Caption = "Passed!"
Case Is >= 75
Label19.Caption = "Passed!"
Case Is <= 74.99
Label19.Caption = "Failed!"
End Select

Dim equivalent As String
equivalent = Label17
Select Case equivalent
Case Is <= 100
Label18.Caption = "1.00"
Case Is >= 97
Label18.Caption = "1.00"
Case Is >= 94
Label18.Caption = "1.25"
Case Is >= 91
Label18.Caption = "1.50"
Case Is >= 88
Label18.Caption = "1.75"
Case Is >= 85
Label18.Caption = "2.00"
Case Is >= 82
Label18.Caption = "2.25"
Case Is >= 79
Label18.Caption = "2.50"
Case Is >= 76
Label18.Caption = "2.75"
Case Is = 75
Label18.Caption = "3.00"
Case Is <= 74.99
Label18.Caption = "5.00"
End Select
End Sub
Private Sub Command3_Click()
Dim a As String
Dim b, c, d, e, f, g, h, i, j, k, m, l, n, o As Integer
a = Val(Form3.Text1.Text)
b = Val(Label6.Caption)
c = Val(Label24.Caption)
d = Val(Label26.Caption)
e = Val(Label16.Caption)
f = Val(Label20.Caption)
g = Val(Label18.Caption)
h = Val(Label17.Caption)
i = Val(Form3.Label17.Caption)
j = Val(Form3.Label18.Caption)
k = Val(Form4.Label17.Caption)
m = Val(Form4.Label8.Caption)
l = Val(Form5.Label17.Caption)
n = Val(Form5.Label18.Caption)
o = Val(Label27.Caption)
MsgBox "Student Name: " + Form3.Text1.Text & vbNewLine & "Quiz Rating: " & b & vbNewLine & "Recitation Rating: " & c & vbNewLine & "Assign. Rating: " & d & vbNewLine & "Finals Exam Rating: " & f & vbNewLine & "Class Standing: " & e & vbNewLine & "Class Average: " & o & vbNewLine & vbNewLine & "Prelim Grade: " & i & vbNewLine & "Prelim Grade Equivalent: " + Form3.Label18.Caption & vbNewLine & "Midterm Grade: " & k & vbNewLine & "Midterm Grade Equivalent: " + Form4.Label18.Caption & vbNewLine & "Finals Grade: " & h & vbNewLine & "Finals Grade Equivalent: " + Me.Label18.Caption

End Sub
Private Sub Command5_Click()
Dim respond As Integer
respond = MsgBox("Back to Prelim Grading System?", vbYesNo + vbQuestion, "CONFIRMATION")
If respond = vbYes Then
Form4.Show
Form5.Hide
End If
End Sub
Private Sub Text11_Change()
Dim rec1, rec2, rec3, recrating As Integer
Text5.Refresh
Text6.Refresh
Text11.Refresh

rec2 = Val(Text5.Text)
rec1 = Val(Text6.Text)
rec3 = Val(Text11.Text)
recrating = (rec1 + rec2 + rec3) / 3
Label24.Caption = recrating

End Sub
Private Sub Text12_Change()
Dim ass1, ass2, ass3, assignrating As Integer
Text9.Refresh
Text10.Refresh
Text12.Refresh

ass1 = Val(Text9.Text)
ass2 = Val(Text10.Text)
ass3 = Val(Text12.Text)
assignrating = (ass1 + ass2 + ass3) / 3
Label26.Caption = assignrating

End Sub

Private Sub Text8_Change()
Dim num1, num2, rate As Integer
Text7.Refresh
Text8.Refresh

num2 = Val(Text7.Text)
num1 = Val(Text8.Text)
rate = (num1 / num2) * 50 + 50
Label20.Caption = rate
End Sub
