Attribute VB_Name = "MainFormCode"
Option Compare Database

Public EditStatus As Boolean
Public AllFieldsComplete As Boolean

Sub CheckUserEntry(FormName As String)

Dim ctr As Control
Dim strMsg As String
Dim responseMsgBox



'Loop through every control on the form
For Each ctr In Forms(FormName).Controls
   
   'Look for a Particular Tag
   If ctr.Tag = "BlkChk" Then
      'Create a List of empty questions
      If IsNull(ctr) Then
         strMsg = strMsg & "- " & ctr.ControlTipText & vbCrLf
         
      End If
   End If
Next ctr

'Did We Find Any Unanswered Questions?
If strMsg <> "" Then
   responseMsgBox = MsgBox("The following fields require a response" & vbCrLf & vbCrLf & _
   strMsg & vbCrLf & vbCrLf & "Please fill in these fields before submitting", vbOKOnly, "Required data")
   
   If responseMsgBox = vbNo Then Glob_Var.QuitConfirmation
   
   
   AllFieldsComplete = False
    
Else
AllFieldsComplete = True

End If


End Sub



