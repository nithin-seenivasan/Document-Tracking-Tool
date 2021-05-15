VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_AddDocumentToProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ReasonsFormClose_Click()

Dim x As Integer
Dim y As Integer
Dim z As Integer
Dim DocumentTitle As String
Dim DueDate As Date

'Get the ProjectDataID that was sent from EditScope
x = Nz(Me.ProjectDataID.Value, 0)

'Count the number of entries matching the ProjectDataID
y = Nz(DCount("DocumentNumber", "DocumentData", "DocumentData.ProjectDataID = " & x & ""), 0)

'Data to be written has to have the document number that's +1 of the count
z = y + 1

DocumentTitle = Nz("M - " + Me.DocumentTitle.Value, "N/A")
DueDate = Nz(Me.DueDate.Value, "00:00:00")

If DueDate = "00:00:00" Or DocumentTitle = "N/A" Then
    MsgBox "Please enter the document title and date"

Else

    DoCmd.SetWarnings False
    'add data to table DocumentData
    ssql = "INSERT INTO DocumentData(ProjectDataID, DocumentNumber, DocumentTitle, DueDate) VALUES(" & x & "," & z & ",'" & DocumentTitle & "','" & DueDate & "')"
    DoCmd.RunSQL ssql
    DoCmd.SetWarnings True
    MsgBox "Successfully added the document to this project"
    DoCmd.Close

End If

End Sub
