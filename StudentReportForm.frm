VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form StudentReportForm 
   BackColor       =   &H00C0E0FF&
   Caption         =   "StudentReportForm"
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
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
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
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
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
      Top             =   960
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
      TabIndex        =   3
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
End
Attribute VB_Name = "StudentReportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MSF.Clear
MSF.Cols = 6
MSF.Rows = 5
tAmt = 0
MSF.TextMatrix(0, 0) = "RegNo"
MSF.TextMatrix(0, 1) = "Name"
MSF.TextMatrix(0, 2) = "Address"
MSF.TextMatrix(0, 3) = "Phone"
MSF.TextMatrix(0, 4) = "CourseName"

MSF.ColWidth(0) = 1000
MSF.ColWidth(1) = 1400
MSF.ColWidth(2) = 1000
MSF.ColWidth(3) = 1400
MSF.ColWidth(4) = 1400

I = 1
If tRS.State = 1 Then tRS.Close
tRS.Open "select * from StudentTab order by CourseName,regno", Conn
Do While tRS.EOF = False
MSF.TextMatrix(I, 0) = tRS(0)
MSF.TextMatrix(I, 1) = tRS(1)
MSF.TextMatrix(I, 2) = tRS(2)
MSF.TextMatrix(I, 3) = tRS(3)
MSF.TextMatrix(I, 4) = tRS(4)

tRS.MoveNext
I = I + 1
MSF.Rows = I + 5
Loop
I = I + 1
End Sub

Private Sub Command2_Click()
If tRS.State = 1 Then tRS.Close
tRS.Open "select * from StudentTab order by CourseName,Semister,regno", Conn
Set StudentList.DataSource = tRS
StudentList.Show
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

