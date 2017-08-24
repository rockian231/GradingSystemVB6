VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0080FF80&
   Caption         =   "Form2"
   ClientHeight    =   4890
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   4500
   LinkTopic       =   "Form2"
   ScaleHeight     =   4890
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option3 
      Caption         =   "FINALS GRADING SYSTEM"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   3615
   End
   Begin VB.OptionButton Option2 
      Caption         =   "MIDTERM GRADING SYSTEM"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   3615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "PRELIM GRADING SYSTEM"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   3960
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ClearAllOptionButtons()
  Dim c As Control
  For Each c In Form2.Controls
    If TypeOf c Is OptionButton Then c.Value = False
  Next
End Sub
Private Sub Command1_Click()
Form2.Hide
Form1.Show
Form3.Hide

End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
MsgBox "PROCEED COMPUTATION??", vbQuestion + vbYesNo, "CONFIRM SELECTION"
Form3.Show
Form2.Hide
Form1.Hide

End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
MsgBox "PLEASE SELECT PRELIM GRADE FIRST", vbCritical + vbOKOnly, "SELECT ERROR"
Form2.Show
Option2.Value = False
End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
MsgBox "PLEASE SELECT PRELIM GRADE FIRST", vbCritical + vbOKOnly, "SELECT ERROR"
Form2.Show
Option3.Value = False
End If

End Sub
