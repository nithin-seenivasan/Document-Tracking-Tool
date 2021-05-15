VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_AddScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command345_Click()
Set tabCtl = Forms!AddScope!TabControl
tabCtl.Value = 1
End Sub

Private Sub Command714_Click()
Set tabCtl = Forms!AddScope!TabControl
tabCtl.Value = 2
End Sub

Private Sub Command865_Click()
Set tabCtl = Forms!AddScope!TabControl
tabCtl.Value = 0
End Sub

Private Sub Command873_Click()
Set tabCtl = Forms!AddScope!TabControl
tabCtl.Value = 1
End Sub

Private Sub Command874_Click()
'Writes Data to the ProjectData and DocumentData tables
    
    DoCmd.SetWarnings False
        
    'add data to table ProjectData
    ssql = "INSERT INTO ProjectData(project_id, subsuppliers_id, subsuppliers_name, scope, contract_form, supplier_contact_name, supplier_contact_phone, supplier_contact_email, contract_date, start_erection_date, mech_completion_date) VALUES('" & Me.contract_name.Column(1) & "','" & Me.supplier_name.Column(1) & "','" & Me.supplier_name.Column(0) & "','" & Me.scope & "','" & Me.contract_form & "','" & Me.supplier_contact_name & "','" & Me.supplier_contact_phone & "','" & Me.supplier_contact_email & "','" & Me.C & "','" & Me.S & "','" & Me.W & "')"
    DoCmd.RunSQL ssql
    'Find the ID of ProjectData
    ProjectDataID = DMax("ID", "ProjectData")
    
    'insert DocumentData
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

        ssql = "INSERT INTO DocumentData(ProjectDataID, DocumentNumber, DocumentTitle, DueDate) VALUES('" & ProjectDataID & "','" & i & "','" & DocumentTitleX & "','" & CalculatedDateX & "')"
        DoCmd.RunSQL ssql
        i = i + 1
        rs.MoveNext
    Loop
    
    DoCmd.SetWarnings True
    
    MsgBox "Successfully Added Documents List"
    
    DoCmd.Close
    
    
    
End Sub

