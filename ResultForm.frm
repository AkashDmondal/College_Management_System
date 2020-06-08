VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ResultForm 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Exam Result Form"
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
   Begin VB.CommandButton Command2 
      Caption         =   "All Student Report"
      Height          =   495
      Left            =   12360
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   0
      Top             =   7440
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   4815
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8493
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
   Begin VB.Label Label1 
      Caption         =   "Enter RegNo"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   960
      Width           =   3375
   End
End
Attribute VB_Name = "ResultForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MSF.Clear
MSF.Cols = 5
MSF.Rows = 20
MSF.ColWidth(0) = 2400
MSF.ColWidth(1) = 1400
MSF.ColWidth(2) = 1400
MSF.ColWidth(3) = 1200
If tRS.State = 1 Then tRS.Close
tRS.Open "select sName,courseName,semister from StudentTab where regno='" & Text1 & "'", Conn
If tRS.EOF = False Then
MSF.TextMatrix(0, 0) = "Reg No"
MSF.TextMatrix(0, 1) = Text1
MSF.TextMatrix(1, 0) = "Name"
MSF.TextMatrix(1, 1) = tRS(0)
MSF.TextMatrix(2, 0) = "CourseName"
MSF.TextMatrix(2, 1) = tRS(1)
MSF.TextMatrix(3, 0) = "Semister"
MSF.TextMatrix(3, 1) = tRS(2)
Else
MSF.Clear
End If

MSF.TextMatrix(5, 0) = "Subject Name"
MSF.TextMatrix(5, 1) = "Practical"
MSF.TextMatrix(5, 2) = "Theory"
MSF.TextMatrix(5, 3) = "Total"
I = 6

If tRS.State = 1 Then tRS.Close
tRS.Open "select subjectName,pMarks,tMarks from MarksTab where regno='" & Text1 & "' order by subjectName", Conn
Do While tRS.EOF = False
MSF.TextMatrix(I, 0) = tRS(0)
MSF.TextMatrix(I, 1) = tRS(1)
MSF.TextMatrix(I, 2) = tRS(2)
MSF.TextMatrix(I, 3) = Val(MSF.TextMatrix(I, 1)) + Val(MSF.TextMatrix(I, 2))
tRS.MoveNext
I = I + 1
Loop


End Sub

Private Sub Command2_Click()

MarksListAll.Show
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Text1_LostFocus()
Text1 = UCase(Text1)
End Sub
