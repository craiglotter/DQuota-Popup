Attribute VB_Name = "Module1"
Private Sub Main()
Dim frm As New MainDialog
Dim obj As Object
frm.Show vbModal
' When the modal form is dismissed, save a
' reference to one of its controls.
Unload frm
Set frm = Nothing
End Sub

