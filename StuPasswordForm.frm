VERSION 5.00
Begin VB.Form StuPasswordForm 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Change Password"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox Text0 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton Butclose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   7920
      TabIndex        =   4
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton butSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Retype Password"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "RegNo"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
End
Attribute VB_Name = "StuPasswordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Butclose_Click()
Unload Me
End Sub

Private Sub butSave_Click()
If Not Text2 = Text3 Then
MsgBox "New password and retyped password are not correct"
Exit Sub
End If

If tRS.State = 1 Then tRS.Close
tRS.Open "select * from studentTab where regNo='" & RegVar & "' and pword='" & Text1 & "'", Conn
If tRS.EOF = True Then
MsgBox "The old password is not correct please check"
Exit Sub
End If

Conn.Execute "update studenttab set pWord='" & Text3 & "' where  regNo='" & RegVar & "'"

MsgBox "Password changes successfully"
End Sub

Private Sub Form_Load()
Text0 = RegVar
End Sub
