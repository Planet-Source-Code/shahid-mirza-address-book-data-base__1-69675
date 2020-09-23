Attribute VB_Name = "SetRecordset"
Public RstLogin As New ADODB.Recordset '4 Login
Public RstNUser As New ADODB.Recordset '4 New User
Public RstContry As New ADODB.Recordset '4 Country
Public RstCity As New ADODB.Recordset '4 City
Public RstAB As New ADODB.Recordset '4 Address Book
Public RstABD As New ADODB.Recordset '4 Address Book Details.
Public RstSearch As New ADODB.Recordset '4 Searching the Records

Public Sub GetDBase()
    RstSQL = "SELECT * FROM " & tbl_login
    RstLogin.Open RstSQL, DB_Con, adOpenStatic, adLockOptimistic
    
    RstSQL = "SELECT * FROM " & tbl_NUser
    RstNUser.Open RstSQL, DB_Con, adOpenStatic, adLockOptimistic

    RstSQL = "SELECT * FROM " & tbl_country
    RstContry.Open RstSQL, DB_Con, adOpenStatic, adLockOptimistic

    RstSQL = "SELECT * FROM " & tbl_city
    RstCity.Open RstSQL, DB_Con, adOpenStatic, adLockOptimistic
    
    RstSQL = "SELECT * FROM " & tbl_AB
    RstAB.Open RstSQL, DB_Con, adOpenStatic, adLockOptimistic

    RstSQL = "SELECT * FROM " & tbl_ABD
    RstABD.Open RstSQL, DB_Con, adOpenStatic, adLockOptimistic
    
    RstSQL = "SELECT * FROM " & tbl_AB & " , " & tbl_ABD
    RstSearch.Open RstSQL, DB_Con, adOpenStatic, adLockOptimistic
End Sub

Public Sub CloseAll()
    RstLogin.Close: RstNUser.Close: RstContry.Close: RstCity.Close
    RstAB.Close: RstABD.Close: LoginName = "":  LoginUID = ""
End Sub

Public Sub GetLogin()
    RstLogin.Close: RstLogin.Open "SELECT * FROM " & tbl_login
    Call LoginAction(frmMain, RstLogin)
End Sub

Public Sub SaveNewUser(RstNU As Recordset, RstLog As Recordset, frm As Form)
    If frm.txtFullName.Text = "" Then MsgBox "Please enter the correct User's Full Name", vbCritical, "Error! Empty Full Name": SendKeys "{Home}+{End}": frm.txtFullName.SetFocus: Exit Sub
    If frm.txtNewUser.Text = "" Then MsgBox "Please enter the correct User ID", vbCritical, "Error! Empty User ID": SendKeys "{Home}+{End}": frm.txtNewUser.SetFocus: Exit Sub
    If frm.txtPWD.Text = "" Then MsgBox "Please enter the correct User's Password", vbCritical, "Error! Empty Password": SendKeys "{Home}+{End}": frm.txtPWD.SetFocus: Exit Sub
        With RstNU
            .AddNew
                .Fields(0).Value = frm.txtFullName.Text: .Fields(1).Value = frm.txtNewUser.Text
                .Fields(2).Value = "Administrator"
            .Update
        End With
        With RstLog
            .AddNew
                .Fields(0).Value = frm.txtNewUser.Text: .Fields(1).Value = frm.txtPWD.Text
                .Fields(2).Value = "Active"
            .Update
        End With
        msgRes = MsgBox("New User for Login has been Created Successfully." & vbCrLf & "Do you want to Create more User", vbExclamation + vbYesNo, "Create New User")
        If msgRes = vbYes Then
            Call Cleartxt(frmMain): Call Setting_Grp(frmMain, "New User")
        ElseIf msgRes = vbNo Then
            Call Cleartxt(frmMain): Call Setting_Grp(frmMain, "Login")
        End If
End Sub

Public Sub SaveData()
With frmMain
    Select Case ViewOpt
        Case "Country" 'For Country Information
            If .txtCountryCode.Text = "" Then MsgBox "Not allow the empty field." & vbCrLf & _
                "Please enter the valid Country Code.", vbCritical, "Error! Country Code...": _
                .txtCountryCode.SetFocus: Exit Sub
            If .txtCountry.Text = "" Then MsgBox "Not allow the empty field." & vbCrLf & _
                "Please enter the valid Country Name.", vbCritical, "Error! Country Name...": _
                .txtCountry.SetFocus: Exit Sub
                
                RstContry.Close: RstContry.Open "SELECT * FROM " & tbl_country & " WHERE Country_Code='" & CntCODE & "'"
                If RstContry.RecordCount <= 0 Then
                    RstContry.AddNew
                        RstContry.Fields(0).Value = CntCODE
                        RstContry.Fields(1).Value = .txtCountry.Text
                    RstContry.Update
                    MsgBox "New Country Information has been saved successfully", vbInformation, "Country Data Save..."
                    .txtID.Text = "(Auto Number)": Call Cleartxt(frmMain): .txtCountryCode.SetFocus
                    .ImgSubmit.Visible = False: .ImgCancel.Visible = False
                    .ImgSubmit_Dis.Visible = True: .ImgCancel_Dis.Visible = True
                ElseIf RstContry.RecordCount > 0 Then
                    MsgBox "Country Code : " & .txtCountryCode & vbCrLf & " Country Name : " & _
                    RstContry.Fields(1).Value & vbCrLf & " already exit."
                    SendKeys "{Home}+{End}": .txtCountryCode.SetFocus: Exit Sub
                End If
                
        Case "City" 'For City Information
            If .txtCityCode.Text = "" Then MsgBox "Not allow the empty field." & vbCrLf & _
                "Please enter the valid City Code.", vbCritical, "Error! City Code...": _
                .txtCityCode.SetFocus: Exit Sub
            If .txtcity.Text = "" Then MsgBox "Not allow the empty field." & vbCrLf & _
                "Please enter the valid City Name.", vbCritical, "Error! City Name...": _
                .txtcity.SetFocus: Exit Sub
                
                RstCity.Close: RstCity.Open "SELECT * FROM " & tbl_city & " WHERE City_Code='" & CitCODE & "'"
                If RstCity.RecordCount <= 0 Then
                    RstCity.AddNew
                        RstCity.Fields(0).Value = CitCODE
                        RstCity.Fields(1).Value = .lblCntCode
                        RstCity.Fields(2).Value = .txtcity.Text
                    RstCity.Update
                    MsgBox "New City Information has been saved successfully", vbInformation, "City Data Save..."
                    .txtID.Text = "(Auto Number)": Call Cleartxt(frmMain): .cmbcountry.SetFocus
                    .ImgSubmit.Visible = False: .ImgCancel.Visible = False
                    .ImgSubmit_Dis.Visible = True: .ImgCancel_Dis.Visible = True
                ElseIf RstCity.RecordCount > 0 Then
                    MsgBox "Country Code : " & .txtCountryCode & vbCrLf & " Country Name : " & _
                    RstContry.Fields(1).Value & vbCrLf & " already exit."
                    SendKeys "{Home}+{End}": .txtCountryCode.SetFocus: Exit Sub
                End If
        Case "New" 'For New Record Information
            If .txtName.Text = "" Then MsgBox "Not allow the empty field." & vbCrLf & _
                "Please enter the valid Name.", vbCritical, "Error! Name Field...": _
                .txtName.SetFocus: Exit Sub

            If .txtCountry.Text = "Select Country" Then MsgBox "Not allow such selected field." & vbCrLf & _
                "Please select valid Country Name.", vbCritical, "Error! Country Name...": _
                .cmbcountry.SetFocus: Exit Sub

            If .cmbcity.Text = "Select City" Then MsgBox "Not allow the such selected field." & vbCrLf & _
                "Please select valid City Name.", vbCritical, "Error! City Name...": _
                .cmbcity.SetFocus: Exit Sub
                
                RstAB.Close: RstAB.Open "SELECT * FROM " & tbl_AB & " WHERE AB_Name='" & .txtName.Text & "'"
                If RstAB.RecordCount <= 0 Then
                    RstAB.AddNew
                        RstAB.Fields(0).Value = .txtID.Text
                        RstAB.Fields(1).Value = .txtName.Text
                        RstAB.Fields(2).Value = .txtAddress.Text
                    RstAB.Update
                    
                    For IntI = 1 To frmMain.Grid.Rows - 1
                        RstABD.AddNew
                            RstABD.Fields(0).Value = .txtID.Text
                            RstABD.Fields(1).Value = frmMain.Grid.TextMatrix(IntI, 2)
                            RstABD.Fields(2).Value = frmMain.Grid.TextMatrix(IntI, 3)
                            RstABD.Fields(3).Value = frmMain.Grid.TextMatrix(IntI, 4)
                            RstABD.Fields(4).Value = frmMain.Grid.TextMatrix(IntI, 5)
                            RstABD.Fields(5).Value = frmMain.Grid.TextMatrix(IntI, 6)
                        RstABD.Update
                    Next
                    
                    MsgBox "New Information has been saved successfully", vbInformation, "Data Save..."
                    .txtID.Text = "(Auto Number)": Call Cleartxt(frmMain): .txtName.SetFocus
                     Call FillCombo 'use to fill the Combo Boxes
                    .ImgSubmit.Visible = False: .ImgCancel.Visible = False: .lblclick.Visible = False
                    .ImgSubmit_Dis.Visible = True: .ImgCancel_Dis.Visible = True
                    
                ElseIf RstAB.RecordCount > 0 Then
                    MsgBox "Name : " & .txtName.Text & vbCrLf & " Country Name : " & _
                    RstAB.Fields(1).Value & vbCrLf & " already exit."
                    SendKeys "{Home}+{End}": .txtName.SetFocus: Exit Sub
                End If
    End Select
    RstContry.Close: RstContry.Open "SELECT * FROM " & tbl_country
    RstCity.Close: RstCity.Open "SELECT * FROM " & tbl_city
    Call InitGrid(frmMain.Grid): Call FillGrid(frmMain, frmMain.Grid)
End With
End Sub
