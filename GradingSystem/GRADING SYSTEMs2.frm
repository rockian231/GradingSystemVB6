VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FF80&
   Caption         =   "Grading System"
   ClientHeight    =   4575
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   1
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&GRADING SYSTEM"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim respond As Integer
respond = MsgBox("CONTINUE TO GRADING SYSTEM", vbYesNo + vbQuestion, "PLEASE CONFIRM")
If respond = vbYes Then
Form2.Show
Else
Form1.Show
End If
End Sub

Private Sub Command2_Click()
Dim respond As Integer
respond = MsgBox("TERMINATE?", vbYesNo + vbQuestion, "CONFIRMATION")
If respond = vbYes Then
Unload Me
End If
End Sub
