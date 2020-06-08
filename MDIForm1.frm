VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "College Management"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   1155
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu MasEntMenu 
      Caption         =   "Master Entries"
      Begin VB.Menu CourseDetMenu 
         Caption         =   "Course Details"
      End
      Begin VB.Menu SubDetMenu 
         Caption         =   "Subject Details"
      End
   End
   Begin VB.Menu StaffDetMenu 
      Caption         =   "Staff Details"
   End
   Begin VB.Menu StudDetMenu 
      Caption         =   "Student Admission"
   End
   Begin VB.Menu AttednaceEntryMenu 
      Caption         =   "Attendance Entry"
   End
   Begin VB.Menu MarksEntryMenu 
      Caption         =   "Marks Entry"
   End
   Begin VB.Menu ReportsMenu 
      Caption         =   "Reports"
      Begin VB.Menu ExamResultForm 
         Caption         =   "Exam Result"
      End
      Begin VB.Menu StudentdetListMenu 
         Caption         =   "Student Details List"
      End
      Begin VB.Menu StaffDetListMenu 
         Caption         =   "Staff Details List"
      End
      Begin VB.Menu AdminAttnMenu 
         Caption         =   "View Attendance"
      End
      Begin VB.Menu AdminMarksMenu 
         Caption         =   "View Marks"
      End
      Begin VB.Menu MarksListSumMenu 
         Caption         =   "Marks List Summary"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CollegeDetMenu_Click()
CollegeDetailsForm.Show
End Sub

Private Sub AdminAttnMenu_Click()
AttendanceViewForm.Show
End Sub

Private Sub AdminMarksMenu_Click()
MarksViewForm.Show
End Sub

Private Sub AdminTimeTableMenu_Click()
ttMainVerForm.Show
End Sub

Private Sub AttednaceEntryMenu_Click()
AttendanceForm.Show
End Sub

Private Sub CourseDetMenu_Click()
CourseNameForm.Show
End Sub

Private Sub DesignationMenu_Click()
DesignationForm.Show
End Sub

Private Sub ExamDatesMenu_Click()
ExamDatesForm.Show
End Sub

Private Sub ExamResultForm_Click()
ResultForm.Show
End Sub

Private Sub HallTicketForm_Click()
HallTicketGenForm.Show
End Sub

Private Sub MarksEntryMenu_Click()
MarksEntryForm.Show
End Sub

Private Sub MarksListSumMenu_Click()
If tRS.State = 1 Then tRS.Close
tRS.Open "select regno,sName,Coursename,semister,subjectname,pmarks,tmarks from MarksTab order by CourseName,Semister,regno", Conn
Set MarksList.DataSource = tRS
MarksList.Show

End Sub

Private Sub MDIForm_Load()
If Conn.State = 0 Then Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CollegeManagement.mdb;Persist Security Info=False"





End Sub

Private Sub QualificationMenu_Click()
QualificationForm.Show
End Sub

Private Sub StaffCPMenu_Click()
ChangePasswordForm.Show
End Sub

Private Sub StaffDetListMenu_Click()
StaffReportForm.Show
End Sub

Private Sub StaffDetMenu_Click()
StaffDetailsForm.Show
End Sub

Private Sub StaffTTMenu_Click()
TTViewForm.Show
End Sub

Private Sub StaffViewAtteMenu_Click()
AttendanceViewForm.Show
End Sub

Private Sub StaffViewMarks_Click()
MarksViewForm.Show
End Sub

Private Sub StuAttnMenu_Click()
StuAttnForm.Show
End Sub

Private Sub StuCPMenu_Click()
StuPasswordForm.Show
End Sub

Private Sub StudDetMenu_Click()
StudentDetailsForm.Show
End Sub

Private Sub StudentdetListMenu_Click()
StudentReportForm.Show
End Sub

Private Sub StuMarksMenu_Click()
StuMarksForm.Show
End Sub

Private Sub StuTTMenu_Click()
TTViewForm.Show
End Sub

Private Sub SubDetMenu_Click()
SubjectNamesForm.Show
End Sub
