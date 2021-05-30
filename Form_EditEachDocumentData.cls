VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditEachDocumentData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub SaveClose_Click()


UserTracking.Login_Tracking ("Document Modified for (project name),(supplier),(equipment/scope), (Doc.Number): " & [contract_name] & "," & [supplier_name] & "," & [scope] & "," & [DocumentNumber])

DoCmd.Close
DoCmd.OpenForm "EditDocumentStatus"

End Sub
