VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditDocumentStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command21_Click()

End Sub

Private Sub Command8_Click()


'Define variables for all the filter elements
x = Nz([scope], 0)
y = Nz([project_id], 0)
z = Nz([subsuppliers_name], 0)
a = Nz([DocumentTitle], 0)

If x = 0 Or y = 0 Or z = 0 Or a = 0 Then
    MsgBox "Please enter all the data to filter"
Else
    DoCmd.Close
    'Filters based on 3 clauses
    DoCmd.OpenForm "EditEachDocumentData", , , "[scope]=" & "'" & x & "'" & "And" & "[contract_name]=" & "'" & y & "'" & "And" & "[supplier_name]=" & "'" & z & "'" & "And" & "[DocumentTitle]=" & "'" & a & "'"

End If

'Lesson Learnt -> Even "AND" needs to be encaspulated in "", since the SQL query is a string

End Sub



Private Sub project_id_AfterUpdate()

'After Project Name is selected, set the Row Source for Supplier name.
subsuppliers_name.RowSource = "SELECT DISTINCT ProjectData.subsuppliers_name FROM ProjectData WHERE ProjectData.project_id = " & Me.project_id.Column(1) & ";"


End Sub


Private Sub subsuppliers_name_AfterUpdate()
x = Me.project_id.Column(1)
y = Me.subsuppliers_name

'After Supplier is selected, set row source for Scope
scope.RowSource = "SELECT DISTINCT ProjectData.scope FROM ProjectData WHERE ProjectData.project_id = " & x & " AND ProjectData.subsuppliers_name = '" & y & "';"

'LEsson Learnt -> If you want to put a number variable inside a string, simply break the string using a ", add variable between two &s, add another " and continue with statement. For a string variable, add ' (see above, Y)

End Sub


Private Sub scope_AfterUpdate()
'Sets the document title to only those marked not deleted

x = Nz(Me.project_id.Column(1), 0) 'ProjectID selected
y = Nz(Me.subsuppliers_name, a) 'Subsupplier name selected
z = Nz(Me.scope, a) 'Scope selected

'After Proj/Supplier/Scope is selected, set row source for DocumentTitle

'Find the ProjectDataID that is used in DocumentData
e = DLookup("ProjectData.ID", "ProjectData", "ProjectData.project_id = " & x & " AND ProjectData.subsuppliers_name = '" & y & "' AND ProjectData.scope = '" & z & "'")
'Set rowsource of the Documents list here
DocumentTitle.RowSource = "SELECT DISTINCT DocumentData.DocumentTitle FROM DocumentData WHERE DocumentData.ProjectDataID = " & e & " AND DocumentData.Deleted = False ;"


End Sub


Private Sub IncludeDeleted_AfterUpdate()
'Includes deleted items also

x = Nz(Me.project_id.Column(1), 0) 'ProjectID selected
y = Nz(Me.subsuppliers_name, a) 'Subsupplier name selected
z = Nz(Me.scope, a) 'Scope selected
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
