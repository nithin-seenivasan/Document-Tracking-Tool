VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Visualization_Filter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command21_Click()

End Sub



Private Sub Command8_Click()

'Filtering Code -> Goal: To be as flexible with the parameters as possible

'Define an array with variables for all the filter elements
Dim A(1 To 5) As String
A(1) = Nz([scope], "Null")
A(2) = Nz([project_id], "Null")
A(3) = Nz([subsuppliers_name], "Null")
A(4) = Nz([DocumentTitle], "Null")
A(5) = Nz([FilterDueDate], "Null")
A(5) = Replace(A(5), ".", "/") 'Replace date to DD/mm/YYYY format

'Define an array with the concatenation commands for the string
Dim B(1 To 5) As String
B(1) = "[scope]=" & "'" & A(1) & "'"
B(2) = "[contract_name]=" & "'" & A(2) & "'"
B(3) = "[supplier_name]=" & "'" & A(3) & "'"
B(4) = "[DocumentTitle]=" & "'" & A(4) & "'"
B(5) = "[DueDate]<" & "#" & A(5) & "#"

'Define the filter command
Dim filter As String
Dim firstFilter As Integer
Dim criteriaString As String

firstFilter = 0 'Variable that tracks the first element in the array
filterCount = 0 'Count that returns how many items actually have a value in the array

'Get filter count
For i = 1 To 5
    If A(i) <> "Null" Then
        filterCount = filterCount + 1
    End If
Next i

'Build the filter query that will be used as the Visualization_View page's data source
'Built to keep the parameters as flexible as possible
For j = 1 To 5
    If A(j) <> "Null" Then 'Only select non-Null values -> the Nz function above sets these values in case of null
        If firstFilter = 0 Then 'For the first element
            filter = filter & B(j)
        ElseIf firstFilter > 0 And firstFilter < filterCount Then 'Between the first and last element
            filter = filter & "And" & B(j)
        ElseIf firstFilter = filterCount Then 'For the last element
            filter = filter & "And" & B(j)
        End If
        firstFilter = firstFilter + 1 'increment the firstFilter variable
    End If
    Next j


If filterCount = 0 Then 'If no parameter is selected
    MsgBox "Please at least one criteria to filter"
Else
    DoCmd.SetWarnings False

    'Opens "Visualization_View" and filters based on the cocatenated string "filter" above
    DoCmd.OpenForm "Visualization_View", , , filter
    'Sets the label in the Viz page to show the current criteria
    Forms!Visualization_View!criteriaString.Caption = "Due Date: " & CStr(A(5)) & " || Contract: " & A(2) & " || Supplier: " & A(3) & " || Scope: " & A(1) & " || Document Title: " & A(4) & ""
    UserTracking.Login_Tracking ("Filtered View Opened -" & A(1) & " || " & A(2) & " || " & A(3) & " || " & A(4) & " || " & A(5))
    DoCmd.SetWarnings True
    
End If


End Sub




Private Sub project_id_AfterUpdate()

'After Project Name is selected, set the Row Source for Supplier name.
subsuppliers_name.Value = Null
scope.Value = Null
scope.RowSource = ""
DocumentTitle.Value = Null
DocumentTitle.RowSource = ""
subsuppliers_name.RowSource = "SELECT DISTINCT ProjectData.subsuppliers_name FROM ProjectData WHERE ProjectData.project_id = " & Me.project_id.Column(1) & ";"

If subsuppliers_name.listCount > 0 Then
    subsuppliers_name.Visible = True
Else
    MsgBox "No data for this selection"
    subsuppliers_name.Visible = False
    scope.Visible = False
    IncludeDeleted.Visible = False
    DocumentTitle.Visible = False
    Command8.Visible = False
End If



End Sub


Private Sub subsuppliers_name_AfterUpdate()
x = Me.project_id.Column(1)
y = Me.subsuppliers_name

'After Supplier is selected, set row source for Scope
scope.Value = Null
DocumentTitle.Value = Null
DocumentTitle.RowSource = ""
scope.RowSource = "SELECT DISTINCT ProjectData.scope FROM ProjectData WHERE ProjectData.project_id = " & x & " AND ProjectData.subsuppliers_name = '" & y & "';"

'LEsson Learnt -> If you want to put a number variable inside a string, simply break the string using a ", add variable between two &s, add another " and continue with statement. For a string variable, add ' (see above, Y)

If scope.listCount > 0 Then
    scope.Visible = True
Else
    MsgBox "No data for this selection"
    scope.Visible = False
    IncludeDeleted.Visible = False
    DocumentTitle.Visible = False
    Command8.Visible = False
End If
End Sub


Private Sub scope_AfterUpdate()
'Sets the document title to only those marked not deleted

x = Nz(Me.project_id.Column(1), 0) 'ProjectID selected
y = Nz(Me.subsuppliers_name, A) 'Subsupplier name selected
z = Nz(Me.scope, A) 'Scope selected

'After Proj/Supplier/Scope is selected, set row source for DocumentTitle

'Find the ProjectDataID that is used in DocumentData
e = DLookup("ProjectData.ID", "ProjectData", "ProjectData.project_id = " & x & " AND ProjectData.subsuppliers_name = '" & y & "' AND ProjectData.scope = '" & z & "'")
'Set rowsource of the Documents list here
DocumentTitle.Value = Null
DocumentTitle.RowSource = "SELECT DISTINCT DocumentData.DocumentTitle FROM DocumentData WHERE DocumentData.ProjectDataID = " & e & " AND DocumentData.Deleted = False ;"

If DocumentTitle.listCount > 0 Then
    DocumentTitle.Visible = True
    'IncludeDeleted.Visible = True
Else
    MsgBox "No data for this selection"
    IncludeDeleted.Visible = False
    'DocumentTitle.Visible = False
    Command8.Visible = False
End If


End Sub


Private Sub IncludeDeleted_AfterUpdate()
'Includes deleted items also

x = Nz(Me.project_id.Column(1), 0) 'ProjectID selected
y = Nz(Me.subsuppliers_name, A) 'Subsupplier name selected
z = Nz(Me.scope, A) 'Scope selected
chkbox = Me.IncludeDeleted

'After Proj/Supplier/Scope is selected, set row source for DocumentTitle

If chkbox = True Then
    'Find the ProjectDataID that is used in DocumentData
    e = DLookup("ProjectData.ID", "ProjectData", "ProjectData.project_id = " & x & " AND ProjectData.subsuppliers_name = '" & y & "' AND ProjectData.scope = '" & z & "'")
    'Set rowsource of the Documents list here
    DocumentTitle.RowSource = "SELECT DISTINCT DocumentData.DocumentTitle FROM DocumentData WHERE DocumentData.ProjectDataID = " & e & ";"
    
Else
    Call scope_AfterUpdate
End If

End Sub

Private Sub DocumentTitle_AfterUpdate()

Command8.Visible = True

End Sub

Private Sub ToggleDueDateFilter_AfterUpdate()

If ToggleDueDateFilter.Value = True Then
    FilterDueDate.Visible = True

Else
    FilterDueDate.Visible = False
    
End If


End Sub
