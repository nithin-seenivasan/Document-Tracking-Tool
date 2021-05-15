VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ReasonsForDueDateChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub SubmitForecastReason_Click()

Dim comment As String
Dim dateChanged As String

comment = Me.Reasons.Value
dateChanged = Format(Now, "DD-MM-YYYY")

'Has to be non-null, otherwise it won't be appended using the formula below
existingString = Nz(Forms!EditEachDocumentData!Comments.Value, "-")

'Save the selected value in the DocumentsData table
Forms!EditEachDocumentData!Comments.Value = existingString + " || Due Date Change, " + dateChanged + " : " + comment


DoCmd.Close
End Sub
