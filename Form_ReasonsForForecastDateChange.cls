VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ReasonsForForecastDateChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub SubmitActualDate_Click()

Dim comment As String
Dim dateChanged As String
comment = Me.Reasons.Column(0)


dateChanged = Format(Now, "DD-MM-YYYY")

'Has to be non-null, otherwise it won't be appended using the formula below
existingString = Nz(Forms!EditEachDocumentData!Reason_DeviationForecastDate.Value, "-")

'Save the selected value in the DocumentsData table
Forms!EditEachDocumentData!Reason_DeviationForecastDate.Value = existingString + " || " + dateChanged + " : " + comment

'MsgBox Forms!EditEachDocumentData!Reason_DeviationActualDate.Value
'MsgBox comment
DoCmd.Close

End Sub
