VERSION 5.00
Begin VB.Form OptionForm 
   BackColor       =   &H00FFC0C0&
   Caption         =   "OptionForm"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4275
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Staff Login"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Student Login"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Admin Login"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "OptionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
LoginForm.Show
Unload Me
End Sub

Private Sub Command2_Click()
StaffLogin.Show
Unload Me
End Sub

Private Sub Command3_Click()
StudentLogin.Show
Unload Me
End Sub

Private Sub Command4_Click()
End
End Sub
