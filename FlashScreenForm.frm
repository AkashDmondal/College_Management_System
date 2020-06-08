VERSION 5.00
Begin VB.Form FlashScreenForm 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "FlashScreen"
   ClientHeight    =   10200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   19035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   2040
   End
   Begin VB.Image Image1 
      Height          =   4290
      Left            =   2520
      Picture         =   "FlashScreenForm.frx":0000
      Top             =   4920
      Width           =   13575
   End
End
Attribute VB_Name = "FlashScreenForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
'Me.Height = 4425
'Me.Width = 13845
I = 1
If Conn.State = 0 Then Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Collegemanagement.mdb;Persist Security Info=False"

End Sub

Private Sub Timer1_Timer()
I = I + 1
If I > 5 Then
Unload Me
LoginForm.Show
End If
End Sub
