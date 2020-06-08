Attribute VB_Name = "Module1"
Public Conn As New ADODB.Connection
Public tRS As New ADODB.Recordset
Public I, J, K, RNo, cNo, NoVar As Long
Public tCodeVar As String
Public EmpVar, aYearVar, RegVar, uTypeVar As String
Function DateFormat(vdate1)
DateFormat = Format(vdate1, "dd-MMM-yyyy")
End Function

Public Function CheckNum(KeyAsciiVar)
'KeyAscii = CheckNum(KeyAscii)
If KeyAsciiVar = 8 Then CheckNum = KeyAsciiVar: Exit Function
If KeyAsciiVar < 46 Or KeyAsciiVar > 57 Then
CheckNum = 0
MsgBox ("Please Enter Numbers Only")
Else
CheckNum = KeyAsciiVar
End If
If KeyAsciiVar = 47 Then CheckNum = 0
End Function
