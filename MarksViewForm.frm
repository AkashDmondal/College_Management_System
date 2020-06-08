VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MarksViewForm 
   BackColor       =   &H00C0E0FF&
   Caption         =   "MarksViewForm"
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
   Begin VB.ComboBox Combo3 
      Height          =   360
      ItemData        =   "MarksViewForm.frx":0000
      Left            =   12000
      List            =   "MarksViewForm.frx":0016
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   12000
      TabIndex        =   3
      Top             =   480
      Width           =   3015
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
      Left            =   12000
      TabIndex        =   1
      Top             =   3240
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
      Left            =   11880
      TabIndex        =   0
      Top             =   7920
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   8295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Semister"
      Height          =   255
      Left            =   12000
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Course Name"
      Height          =   375
      Left            =   12000
      TabIndex        =   4
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "MarksViewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hNo As Long
Dim colVar, rowVar As Integer
Private Sub Command1_Click()
MSF.Clear
MSF.Cols = 6
MSF.Rows = 5

MSF.TextMatrix(0, 0) = "rNo"
MSF.TextMatrix(0, 1) = "RegNo"
MSF.TextMatrix(0, 2) = "SName"
MSF.TextMatrix(0, 3) = "SubName"
MSF.TextMatrix(0, 4) = "Practical"
MSF.TextMatrix(0, 5) = "Exam Marks"

MSF.ColWidth(0) = 0
MSF.ColWidth(1) = 1400
MSF.ColWidth(2) = 3000
MSF.ColWidth(3) = 1000
MSF.ColWidth(4) = 1000
If tRS.State = 1 Then tRS.Close
tRS.Open "select max(rNo) from MarksTab"
hNo = IIf(IsNull(tRS(0)), 1000, tRS(0)) + 1
I = 1

If tRS.State = 1 Then tRS.Close
tRS.Open "select RegNo,sName,subjectName,pMarks,tMarks from MarksTab  where coursename='" & Combo1 & "' and semister='" & Combo3 & "'  order by regNo,subjectName", Conn
Do While tRS.EOF = False
MSF.TextMatrix(I, 0) = hNo
MSF.TextMatrix(I, 1) = tRS(0)
MSF.TextMatrix(I, 2) = tRS(1)
MSF.TextMatrix(I, 3) = tRS(2)
MSF.TextMatrix(I, 4) = tRS(3)
MSF.TextMatrix(I, 5) = tRS(4)
tRS.MoveNext
hNo = hNo + 1
I = I + 1
MSF.Rows = I + 5
Loop


End Sub



Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
If tRS.State = 1 Then tRS.Close
tRS.Open "select CourseName from CourseTab  order by CourseName", Conn
Do While tRS.EOF = False
Combo1.AddItem (tRS(0))
tRS.MoveNext
Loop


End Sub

Private Sub MSF_KeyPress(KeyAscii As Integer)

rowVar = MSF.Row
colVar = MSF.Col
If colVar = 0 Or colVar = 1 Or colVar = 2 Or colVar = 3 Then Exit Sub
'MsgBox KeyAscii
    If KeyAscii = 13 Then
    MSF.Row = MSF.Row + 1
    MSF.SetFocus
    Exit Sub
    End If

If KeyAscii = 8 Then
    If Len(MSF.TextMatrix(rowVar, colVar)) > 0 Then MSF.TextMatrix(rowVar, colVar) = Left(MSF.TextMatrix(rowVar, colVar), Len(MSF.TextMatrix(rowVar, colVar)) - 1)
Else
MSF.TextMatrix(rowVar, colVar) = MSF.TextMatrix(rowVar, colVar) & Chr((KeyAscii))
End If


End Sub

