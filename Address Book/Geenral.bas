Attribute VB_Name = "Geenral"
Public Sub Cleartxt(frm As Form)
    For Each oCtrl In frm
        If TypeOf oCtrl Is TextBox Then oCtrl.Text = ""
    Next
End Sub

Public Sub HideCtrl()

End Sub

