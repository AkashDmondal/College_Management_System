VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form StudentDetailsForm 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Student  Details Form"
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
      ItemData        =   "StudentDetailsForm.frx":0000
      Left            =   2160
      List            =   "StudentDetailsForm.frx":0016
      TabIndex        =   17
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   3000
      Width           =   3135
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   2160
      TabIndex        =   4
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   1200
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton Butclose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   7920
      TabIndex        =   12
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton ButDis 
      Caption         =   "&Display"
      Height          =   495
      Left            =   6360
      TabIndex        =   11
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton butDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton butModify 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton butSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   5
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
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   8295
      Left            =   10200
      TabIndex        =   10
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Semister"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Course Name"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Reg No"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "StudentDetailsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pkVar As String
Private Sub Butclose_Click()
Unload Me
End Sub

Private Sub butDelete_Click()
If Text1 = "" Then
MsgBox "Please Enter The RegNo"
Exit Sub
End If

Conn.Execute "delete from StudentTab where RegNo ='" & pkVar & "'"

butSave.Enabled = False
ButNew.Enabled = True
butModify.Enabled = False
butDelete.Enabled = False
End Sub

Private Sub ButDis_Click()
MSF.Clear
MSF.Cols = 2
MSF.TextMatrix(0, 0) = "RegNo"
MSF.TextMatrix(0, 1) = "Name"

MSF.ColWidth(0) = 2000
MSF.ColWidth(1) = 2000

I = 1
If tRS.State = 1 Then tRS.Close
tRS.Open "select * from StudentTab  order by RegNo", Conn
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
MsgBox "Please Enter The RegNo"
Exit Sub
End If

Conn.Execute "delete from StudentTab where RegNo ='" & pkVar & "'"
Conn.Execute "insert into StudentTab values('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Combo1 & "','" & Combo3 & "','abcd')"

butSave.Enabled = False
ButNew.Enabled = True
butModify.Enabled = False
butDelete.Enabled = False
End Sub

Private Sub ButNew_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

butSave.Enabled = True
ButNew.Enabled = False
butModify.Enabled = False
butDelete.Enabled = False

End Sub

Private Sub butSave_Click()
If Text1 = "" Then
MsgBox "Please Enter The RegNo"
Exit Sub
End If


Conn.Execute "insert into StudentTab values('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Combo1 & "','" & Combo3 & "','abcd')"
butSave.Enabled = False
ButNew.Enabled = True
butModify.Enabled = False
butDelete.Enabled = False
End Sub






Private Sub Form_Load()
If tRS.State = 1 Then tRS.Close
tRS.Open "select courseName from courseTab order by courseName", Conn
Do While tRS.EOF = False
Combo1.AddItem (tRS(0) & "")
tRS.MoveNext
Loop



ButDis_Click

End Sub

Private Sub MSF_DblClick()
If MSF.TextMatrix(MSF.Row, 0) = "" Then Exit Sub
pkVar = MSF.TextMatrix(MSF.Row, 0)

If tRS.State = 1 Then tRS.Close
tRS.Open "select * from StudentTab where RegNo ='" & pkVar & "'", Conn
If tRS.EOF = False Then
Text1 = tRS(0)
Text2 = tRS(1)
Text3 = tRS(2)
Text4 = tRS(3)
Combo1 = tRS(4)
Combo3 = tRS(5)

End If
butSave.Enabled = False
ButNew.Enabled = True
butModify.Enabled = True
butDelete.Enabled = True
End Sub

