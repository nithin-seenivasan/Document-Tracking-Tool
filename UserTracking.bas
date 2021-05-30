Attribute VB_Name = "UserTracking"
Option Compare Database

Sub Login_Tracking(Optional Splash As String = "Null")

'tracks who opened and when

Dim dbs As DAO.Database
Set dbs = CurrentDb

Dim UN As String
Dim CN As String
Dim DL As String
Dim TL As String
Dim uSQL As String
Dim Win_name As String
Win_name = "Splash Screen"


UN = Environ$("username")
CN = Environ$("computername")
DL = Date
TL = Time()

If Splash <> "Null" Then Win_name = Splash

DoCmd.SetWarnings False

DoCmd.RunSQL "INSERT INTO UserTracker (UserName, ComputerName, Date_Login, Time_Login, FrontEnd) VALUES ('" & UN & "','" & CN & "','" & DL & "','" & TL & "','" & Win_name & "');"

DoCmd.SetWarnings True

End Sub


Sub Current_User(Optional Splash As String = "Null")

'tracks current user for code

Dim dbs As DAO.Database
Set dbs = CurrentDb

Dim UN As String
Dim CN As String
Dim DL As String
Dim TL As String
Dim uSQL As String
Dim Win_name As String
Win_name = "Splash Screen"


UN = Environ$("username")
CN = Environ$("computername")
DL = Date
TL = Time()

If Splash <> "Null" Then Win_name = Splash

DoCmd.SetWarnings False

DoCmd.RunSQL "UPDATE CurrentUser SET UserName = '" & UN & "', ComputerName = '" & CN & "' WHERE Identifier = '1'"

DoCmd.SetWarnings True

End Sub
