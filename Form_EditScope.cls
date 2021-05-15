VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Command1023_Click()

End Sub



Private Sub WarningBack_Click()
Call Step3Previous_Click
End Sub

Private Sub WarningNext_Click()
Set tabCtl = Forms!EditScope!TabControl
tabCtl.Value = 2
End Sub

Private Sub Step1Next_Click()

'Set Row Sources for all controls on form
x = Nz(Me.project_id_ES.Column(1), 0) 'ProjectID selected
y = Nz(Me.subsuppliers_name_ES, a) 'Subsupplier name selected
z = Nz(Me.scope_ES, a) 'Scope selected

'After Proj/Supplier/Scope is selected, set row source for Contract_Form
e = DLookup("ProjectData.contract_form", "ProjectData", "ProjectData.project_id = " & x & " AND ProjectData.subsuppliers_name = '" & y & "' AND ProjectData.scope = '" & z & "'")
contract_form.Value = e

'After Proj/Supplier/Scope is selected, set row source for supplier_contact_name
e = DLookup("ProjectData.supplier_contact_name", "ProjectData", "ProjectData.project_id = " & x & " AND ProjectData.subsuppliers_name = '" & y & "' AND ProjectData.scope = '" & z & "'")
supplier_contact_name.Value = e

'After Proj/Supplier/Scope is selected, set row source for supplier_contact_phone
e = DLookup("ProjectData.supplier_contact_phone", "ProjectData", "ProjectData.project_id = " & x & " AND ProjectData.subsuppliers_name = '" & y & "' AND ProjectData.scope = '" & z & "'")
supplier_contact_phone.Value = e

'After Proj/Supplier/Scope is selected, set row source for supplier_contact_email
e = DLookup("ProjectData.supplier_contact_email", "ProjectData", "ProjectData.project_id = " & x & " AND ProjectData.subsuppliers_name = '" & y & "' AND ProjectData.scope = '" & z & "'")
supplier_contact_email.Value = e

'After Proj/Supplier/Scope is selected, set row source for C
e = DLookup("ProjectData.contract_date", "ProjectData", "ProjectData.project_id = " & x & " AND ProjectData.subsuppliers_name = '" & y & "' AND ProjectData.scope = '" & z & "'")
C.Value = e

'After Proj/Supplier/Scope is selected, set row source for S
e = DLookup("ProjectData.start_erection_date", "ProjectData", "ProjectData.project_id = " & x & " AND ProjectData.subsuppliers_name = '" & y & "' AND ProjectData.scope = '" & z & "'")
S.Value = e

'After Proj/Supplier/Scope is selected, set row source for W
e = DLookup("ProjectData.mech_completion_date", "ProjectData", "ProjectData.project_id = " & x & " AND ProjectData.subsuppliers_name = '" & y & "' AND ProjectData.scope = '" & z & "'")
W.Value = e

Set tabCtl = Forms!EditScope!TabControl
tabCtl.Value = 1


End Sub

Private Sub SaveAndCloseStep2_Click()
'Updates data in ProjectData based on input from this form

'Set Row Sources for all controls on form
x = Nz(Me.project_id_ES.Column(1), 0) 'ProjectID selected
y = Nz(Me.subsuppliers_name_ES, a) 'Subsupplier name selected
z = Nz(Me.scope_ES, a) 'Scope selected

DoCmd.SetWarnings False

If x <> 0 And y <> a And z <> a Then
    'Update table ProjectData
    ssql = "UPDATE ProjectData SET contract_form = '" & Me.contract_form & "', supplier_contact_name = '" & Me.supplier_contact_name & "', supplier_contact_email = '" & Me.supplier_contact_email & "', supplier_contact_phone = '" & Me.supplier_contact_phone & "', contract_date = '" & Me.C & "', start_erection_date = '" & Me.S & "', mech_completion_date =  '" & Me.W & "' WHERE ProjectData.project_id = " & x & " AND ProjectData.subsuppliers_name = '" & y & "' AND ProjectData.scope = '" & z & "'"
    DoCmd.RunSQL ssql
    MsgBox "Successfully updated Scope Data"
                                                                                                                                                                                                        ', contract_date, start_erection_date, mech_completion_date
                                                                                                                                                                                                        ' '" & Me.C & "','" & Me.S & "','" & Me.W & "'
Else
    MsgBox "Invalid selection"
End If

DoCmd.SetWarnings True
DoCmd.Close

End Sub


Private Sub Command874_Click()
'Updates all Due Dates in DocumentData based on changes on this form

'Set Row Sources for all controls on form
x = Nz(Me.project_id_ES.Column(1), 0) 'ProjectID selected
y = Nz(Me.subsuppliers_name_ES, a) 'Subsupplier name selected
z = Nz(Me.scope_ES, a) 'Scope selected

'Find the ProjectDataID that is used in DocumentData
e = DLookup("ProjectData.ID", "ProjectData", "ProjectData.project_id = " & x & " AND ProjectData.subsuppliers_name = '" & y & "' AND ProjectData.scope = '" & z & "'")


'Update DueData to DocumentData tables

DoCmd.SetWarnings False

'Update DocumentData
Dim strSQL As String 'Table name for recordset
Dim rs As DAO.Recordset
strSQL = "DocumentsTemplate"
Set rs = CurrentDb.OpenRecordset(strSQL) 'Open record set which is the table
i = 1 'initialize first document

Do Until rs.EOF 'Until end of the table
    'Find document Title for the appropriate number (directly from the DocumentsTemplate table)
    DocumentTitleX = DLookup("DocumentsTemplate.DocumentTitle", "DocumentsTemplate", "DocumentsTemplate.ID = " & i & "")
    'Calculate dates (using data directly from DocumentTemplate table)
    CalculatedDateX = DateAdd("ww", DLookup("DocumentsTemplate.DueDateWeek", "DocumentsTemplate", "DocumentsTemplate.ID = " & i & ""), Me.C)
    'Update the SQL table
    ssql = "UPDATE DocumentData SET DueDate = '" & CalculatedDateX & "' WHERE DocumentData.ProjectDataID = " & e & " AND DocumentData.DocumentNumber = " & i & ""
    DoCmd.RunSQL ssql
    i = i + 1
    rs.MoveNext
Loop

DoCmd.SetWarnings True

MsgBox "Successfully Updated DueDates"

'Update all other scope data in ProjectData
Call SaveAndCloseStep2_Click
    
    
End Sub

Private Sub Step2Previous_Click()
'Navigate tab control

Set tabCtl = Forms!EditScope!TabControl
tabCtl.Value = 0
End Sub

Private Sub Step3Previous_Click()
'Navigate tab control

Set tabCtl = Forms!EditScope!TabControl
tabCtl.Value = 1
End Sub


Private Sub Step2Next_Click()
'Navigate tab control

Set tabCtl = Forms!EditScope!TabControl
tabCtl.Value = 3
End Sub

Private Sub OpenAddDocumentToProject_Click()
'Set the ProjectDataID in the sub form, so that it can be used for adding additional documents to the DocumentData table from the form

'Set Row Sources for all controls on form
x = Nz(Me.project_id_ES.Column(1), 0) 'ProjectID selected
y = Nz(Me.subsuppliers_name_ES, a) 'Subsupplier name selected
z = Nz(Me.scope_ES, a) 'Scope selected

'Find the ProjectDataID that is used in DocumentData
e = DLookup("ProjectData.ID", "ProjectData", "ProjectData.project_id = " & x & " AND ProjectData.subsuppliers_name = '" & y & "' AND ProjectData.scope = '" & z & "'")

DoCmd.OpenForm ("AddDocumentToProject") 'Open the form

Forms!AddDocumentToProject!ProjectDataID.Value = e 'Set the value of the ProjectDataID to this invisible field



End Sub

Private Sub project_id_ES_Update()
'After Project Name is selected, set the Row Source for Supplier name
subsuppliers_name_ES.RowSource = "SELECT DISTINCT ProjectData.subsuppliers_name FROM ProjectData WHERE ProjectData.project_id = " & Me.project_id_ES.Column(1) & ";"

End Sub

Private Sub project_id_ES_Change()
Call project_id_ES_Update

End Sub

Private Sub subsuppliers_name_ES_AfterUpdate()
'After Supplier is selected, set row source for Scope

x = Nz(Me.project_id_ES.Column(1), 0) 'ProjectID selected
y = Nz(Me.subsuppliers_name_ES, a) 'Subsupplier name selected
z = Nz(Me.scope_ES, a) 'Scope selected

'After Supplier is selected, set row source for Scope
scope_ES.RowSource = "SELECT DISTINCT ProjectData.scope FROM ProjectData WHERE ProjectData.project_id = " & x & " AND ProjectData.subsuppliers_name = '" & y & "';"
End Sub


