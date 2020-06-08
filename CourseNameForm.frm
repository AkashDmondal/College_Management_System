VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CourseNameForm 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Course Name Form"
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
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Butclose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   7920
      TabIndex        =   9
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton ButDis 
      Caption         =   "&Display"
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton butDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton butModify 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton butSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton ButNew 
      Caption         =   "&New"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   8295
      Left            =   10200
      TabIndex        =   7
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   14631
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "CourseNameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CourseNameVar As String
Private Sub Butclose_Click()
Unload Me
End Sub

Private Sub butDelete_Click()
If Text1 = "" Then
MsgBox "Please Enter The CourseName"
Exit Sub
End If


Conn.Execute "delete from CourseTab where CourseName ='" & CourseNameVar & "'"

butSave.Enabled = False
ButNew.Enabled = True
butModify.Enabled = True
butDelete.Enabled = True
End Sub

Private Sub ButDis_Click()
MSF.Clear
MSF.Cols = 2
MSF.TextMatrix(0, 0) = "CourseName"
MSF.TextMatrix(0, 1) = "Details"

MSF.ColWidth(0) = 2000
MSF.ColWidth(1) = 2000

I = 1
If tRS.State = 1 Then tRS.Close
tRS.Open "select * from CourseTab  order by CourseName", Conn
Do While tRS.EOF = False
MSF.TextMatrix(I, 0) = tRS(0)
MSF.TextMatrix(I, 1) = tRS(1)
tRS.MoveNext
I = I + 1
MSF.Rows = I + 5
Loop
End Sub

Private Sub butModify_Click()
If Text1 = "" Then
MsgBox "Please Enter The CourseName"
Exit Sub
End If

Conn.Execute "delete from CourseTab where CourseName ='" & CourseNameVar & "'"
Conn.Execute "insert into CourseTab values('" & Text1 & "','" & Text2 & "')"
butSave.Enabled = False
ButNew.Enabled = True
butModify.Enabled = False
butDelete.Enabled = False
End Sub

Private Sub ButNew_Click()
Text1.Text = ""
Text2.Text = ""


butSave.Enabled = True
ButNew.Enabled = False
butModify.Enabled = False
butDelete.Enabled = False

End Sub

Private Sub butSave_Click()
If Text1 = "" Then
MsgBox "Please Enter The CourseName"
Exit Sub
End If
If tRS.State = 1 Then tRS.Close
tRS.Open "select * from CourseTab where CourseName='" & Text1.Text & "'", Conn
If tRS.EOF = False Then
MsgBox ("This CourseName is allready present,Please check")
Exit Sub
End If

Conn.Execute "insert into CourseTab values('" & Text1 & "','" & Text2 & "')"
butSave.Enabled = False
ButNew.Enabled = True
butModify.Enabled = False
butDelete.Enabled = False
End Sub






Private Sub Form_Load()

'If tRS.State = 1 Then tRS.Close
'tRS.Open "select JobName from JobTypeTab order by JobName", Conn
'Do While tRS.EOF = False
'Combo1.AddItem tRS(0)
'tRS.MoveNext
'Loop
ButDis_Click

End Sub

Private Sub MSF_DblClick()
If MSF.TextMatrix(MSF.Row, 0) = "" Then Exit Sub
CourseNameVar = MSF.TextMatrix(MSF.Row, 0)

If tRS.State = 1 Then tRS.Close
tRS.Open "select * from CourseTab where CourseName ='" & CourseNameVar & "'", Conn
If tRS.EOF = False Then
Text1 = tRS(0)
Text2 = tRS(1)

End If
butSave.Enabled = False
ButNew.Enabled = True
butModify.Enabled = True
butDelete.Enabled = True
End Sub

