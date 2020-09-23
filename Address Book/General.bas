Attribute VB_Name = "General"

Public Sub Cleartxt(frm As Form)
    For Each oCtrl In frm
        If TypeOf oCtrl Is TextBox Then oCtrl.Text = ""
        If TypeOf oCtrl Is ComboBox Then oCtrl.Clear
    Next
    Call InitGrid(frmMain.Grid) 'use to initiate the Grid.
End Sub

Public Sub frmSetting(frm As Form)
    With frm
        .Height = 6040: .Width = 9150
        .Image1.Width = 9020: .Image1.Height = 5920: .Image1.Top = 60: .Image1.Left = 80
        .Image2.Width = 375: .Image2.Height = 301: .Image2.Top = 2891.97: .Image2.Left = 2850
        
         Set .LstSearch.SmallIcons = .i16x16: Set .LstSearch.Icons = .i16x16
         Set .Image2.Picture = .i16x16.ListImages(7).Picture

        Call Cleartxt(frm): .txtID.Text = "(Auto Number)"
        Call InitGrid(frm.Grid) 'use to initialize the Grid View
        Call FillCombo 'use 4 filling the Grid Boxes
        Call FillGrid(frmMain, frmMain.Grid) 'use to fill the Grid.
        .ImgSubmit_Dis.Visible = True: .ImgCancel_Dis.Visible = True
    End With
End Sub

Public Sub set_menubtn(frm As Form)
    With frm
        .ImgHome.Visible = Not .ImgHome.Visible
        .ImgAbout.Visible = Not .ImgAbout.Visible
        .ImgCountry.Visible = Not .ImgCountry.Visible
        .ImgCity.Visible = Not .ImgCity.Visible
        .ImgNewRec.Visible = Not .ImgNewRec.Visible
        .ImgSearch.Visible = Not .ImgSearch.Visible
        .ImgContact.Visible = Not .ImgContact.Visible

        .ImgSubmit.Visible = False: .ImgCancel.Visible = False
        .ImgSubmit_Dis.Visible = False: .ImgCancel_Dis.Visible = False
    End With
End Sub

Public Sub Prior_Reload(frm As Form)
    For Each oCtrl In frm
        If TypeOf oCtrl Is TextBox Then oCtrl.Visible = False
        If TypeOf oCtrl Is Label Then oCtrl.Visible = False
        If TypeOf oCtrl Is ComboBox Then oCtrl.Visible = False
        If TypeOf oCtrl Is MSHFlexGrid Then oCtrl.Visible = False
        If TypeOf oCtrl Is CommandButton Then oCtrl.Visible = False
        If TypeOf oCtrl Is Frame Then oCtrl.Visible = False
    Next
    With frmMain
        .ImgFind.Visible = False
        .ImgButter.Visible = False
        .ImgMyPic.Visible = False
    End With
End Sub

Public Sub Setting_Grp(frm As Form, GrpOpt As String)
    With frm
        ViewOpt = GrpOpt: Call Prior_Reload(frm) 'use to intilize before settingup.
        Call Cleartxt(frmMain) 'use to wash out the all textboxes of form
        If GrpOpt = "New" Then
            .lblID.Top = 1150: .lblID.Left = 2820: .lblID.Visible = True 'for record id
            .txtID.Top = 1150: .txtID.Left = 4320: .txtID.Visible = True: .txtID.Text = "(Auto Number)"
            
            .lblName.Width = 1095: .lblName.Caption = "Name : " 'for name
            .lblName.Top = 1530.334: .lblName.Left = 1680: .lblName.Visible = True
            .txtName.Top = 1500.209: .txtName.Left = 2800: .txtName.Visible = True
            
            .lblAddress.Top = 1927.98: .lblAddress.Left = 1680: .lblAddress.Visible = True 'for address
            .txtAddress.Top = 1927.98: .txtAddress.Left = 2800: .txtAddress.Visible = True
            
            .lblCity.Top = 2379.85: .lblCity.Left = 5760: .lblCity.Visible = True 'for city
            .cmbcity.Top = 2349.726: .cmbcity.Left = 6840: .cmbcity.Visible = True
            
            .lblCountry.Top = 1958.105: .lblCountry.Left = 5760: .lblCountry.Visible = True 'for country
            .cmbcountry.Top = 1927.98: .cmbcountry.Left = 6840: .cmbcountry.Visible = True
            
            .lblNumber.Top = 2891.97: .lblNumber.Left = 1550: .lblNumber.Visible = True 'for number indication entry
            .lblNumber.Caption = "Enter Numbers "
            
            .lblLandline.Top = 2891.97: .lblLandline.Left = 2800: .lblLandline.Visible = True 'for landline number
            .txtPhone.Top = 3200.246: .txtPhone.Left = 2800: .txtPhone.Visible = True
            
            .lblCell.Top = 2891.97: .lblCell.Left = 4620: .lblCell.Visible = True 'for cell number
            .txtMobile.Top = 3200.246: .txtMobile.Left = 4620: .txtMobile.Visible = True
            
            .lblEmail.Top = 2891.97: .lblEmail.Left = 6455: .lblEmail.Visible = True 'for email id
            .txtEmail.Top = 3200.246: .txtEmail.Left = 6455: .txtEmail.Visible = True
            
            .lblHead.Caption = "New Record": .lblHead.Visible = True: .Image2.Visible = True 'Arrow Image
            .Grid.Visible = True: .lbllogout.Visible = True:
            .lblDisplayUID = "System Login as " & LoginUID: .lblDisplayUID.Visible = True
            
            .lblHello.Caption = "HELLO   Mr. " & LoginName: .lblUserName.Caption = LoginName
            .lblHello.Visible = True: .lblUserName.Visible = True: .txtName.SetFocus
            Call InitGrid(frm.Grid): Call FillGrid(frmMain, frm.Grid) 'use to Fill the Data in Grid.
             .ImgSubmit_Dis.Visible = True: .ImgCancel_Dis.Visible = True
            .ImgSubmit.Visible = False: .ImgCancel.Visible = False
        '----------For New Country Setting--------------------
        ElseIf GrpOpt = "Country" Then

            .lblID.Top = 1150: .lblID.Left = 2350: .lblID.Visible = True 'for record id
            .txtID.Top = 1150: .txtID.Left = 4320: .txtID.Visible = True: .txtID.Text = "(Auto Number)"
            
            
            .lblName.Width = .lblCountry.Width: .lblName.Caption = "Code # : ": 'for country code
            .lblName.Left = 2760: .lblName.Top = 1958.105: .lblName.Visible = True
            .txtCountryCode.Left = 4320: .txtCountryCode.Top = .lblName.Top
            .txtCountryCode.Visible = True
            
            .lblCountry.Left = 2760: .lblCountry.Top = 2349.726: .lblCountry.Visible = True 'for country name
            .txtCountry.Left = 4320: .txtCountry.Top = 2349.726: .txtCountry.Visible = True
            
            .lblHead.Caption = "New Country": .lblHead.Visible = True: .Image2.Visible = False 'Arrow Image
            .Grid.Visible = True: .txtCountryCode.SetFocus: .lbllogout.Visible = True
            .lblDisplayUID = "System Login as " & LoginUID: .lblDisplayUID.Visible = True
            
            .lblHello.Caption = "HELLO   Mr. " & LoginName: .lblUserName.Caption = LoginName
            .lblHello.Visible = True: .lblUserName.Visible = True
             Call InitGrid(frm.Grid): Call FillGrid(frmMain, frm.Grid) 'use to Fill the Data in Grid.
             .ImgSubmit_Dis.Visible = True: .ImgCancel_Dis.Visible = True
             .ImgSubmit.Visible = False: .ImgCancel.Visible = False
        '----------For New City Entries-----------------------
        ElseIf GrpOpt = "City" Then
            
            .lblID.Top = 1150: .lblID.Left = 2350: .lblID.Visible = True 'for record id
            .txtID.Top = 1150: .txtID.Left = 4320: .txtID.Visible = True: .txtID.Text = "(Auto Number)"
    
            .lblCountry.Left = 2760: .lblCountry.Top = 1500.209: .lblCountry.Visible = True 'for country selection
            .cmbcountry.Left = 4320: .cmbcountry.Top = 1500.209: .cmbcountry.Visible = True
            
            .lblName.Width = .lblCountry.Width: .lblName.Caption = "Code # : " 'for city code
            .lblName.Left = 2760: .lblName.Top = 1958.105: .lblName.Visible = True
            .txtCityCode.Left = 4320: .txtCityCode.Top = .lblName.Top
            .txtCityCode.Visible = True
            
            .lblCity.Left = 2760: .lblCity.Top = 2349.726: .lblCity.Visible = True 'for city name
            .txtcity.Left = 4320: .txtcity.Top = 2349.726: .txtcity.Visible = True
            
            .lblHead.Caption = "New City": .lblHead.Visible = True: .Image2.Visible = False 'Arrow Image
            .Grid.Visible = True: .cmbcountry.SetFocus: .lbllogout.Visible = True
            .lblDisplayUID = "System Login as " & LoginUID: .lblDisplayUID.Visible = True

            .lblHello.Caption = "HELLO   Mr. " & LoginName: .lblUserName.Caption = LoginName
            .lblHello.Visible = True: .lblUserName.Visible = True
             Call InitGrid(frm.Grid): Call FillGrid(frmMain, frm.Grid) 'use to Fill the Data in Grid.
             .ImgSubmit_Dis.Visible = True: .ImgCancel_Dis.Visible = True
             .ImgSubmit.Visible = False: .ImgCancel.Visible = False
        '----------For Search Criteria-----------------------
        ElseIf GrpOpt = "Search" Then
            .lblNumber.Top = 1150: .lblNumber.Left = 700 'use for search by Option
            .lblNumber.Caption = "Search By : ": .lblNumber.Visible = True
            
            .txtSearName.Left = 2280: .txtSearName.Top = 1150: .txtSearName.Visible = True
            .lblbyName.Left = 2280: .lblbyName.Top = 1450: .lblbyName.Visible = True
            
            .txtSearLandLine.Left = 4560: .txtSearLandLine.Top = 1150: .txtSearLandLine.Visible = True
            .lblByLandLine.Left = 4560: .lblByLandLine.Top = 1450: .lblByLandLine.Visible = True
            
            .txtSearCell.Left = 6840: .txtSearCell.Top = 1150: .txtSearCell.Visible = True
            .lblByMobile.Left = 6840: .lblByMobile.Top = 1450: .lblByMobile.Visible = True
            
            .ImgFind.Left = 7500: .ImgFind.Top = 1800: .ImgFind.Visible = True
            .lblSearchedName.Left = 4560: .lblSearchedName.Top = 1800
            
            .lblOR1.Left = 4160: .lblOR1.Top = 1150: .lblOR1.Visible = True
            .lblOR2.Left = 6440: .lblOR2.Top = 1150: .lblOR2.Visible = True
            
            .lblHead.Caption = "Search Record": .lblHead.Visible = True: .txtSearName.SetFocus
            .lbllogout.Visible = True: .lblDisplayUID = "System Login as " & LoginUID: .lblDisplayUID.Visible = True
            .Image2.Visible = False 'Arrow Image
            
            .lblHello.Caption = "HELLO   Mr. " & LoginName: .lblUserName.Caption = LoginName
            .lblHello.Visible = True: .lblUserName.Visible = True
            
            .ImgSubmit_Dis.Visible = False: .ImgCancel_Dis.Visible = False
            .ImgSubmit.Visible = False: .ImgCancel.Visible = False
        '----------For Create New User----------------------
        ElseIf GrpOpt = "New User" Then
            .lblFullName.Top = 1928: .lblFullName.Left = 1500: .lblFullName.Visible = True 'use for user full name
            .txtFullName.Top = 1928: .txtFullName.Left = 3500: .txtFullName.Visible = True
            
            .lblNewUser.Caption = "Enter New User ID : "
            .lblNewUser.Top = 2410: .lblNewUser.Left = 1500: .lblNewUser.Visible = True 'use for new user.
            .txtNewUser.Top = 2410: .txtNewUser.Left = 3500: .txtNewUser.Visible = True
            
            .lblPWD.Top = 2892: .lblPWD.Left = 1500: .lblPWD.Visible = True 'use for new user password.
            .txtPWD.Top = 2892: .txtPWD.Left = 3500: .txtPWD.Visible = True
            
            .lblConPWD.Top = 3374: .lblConPWD.Left = 1500: .lblConPWD.Visible = True 'use for confirm password.
            .txtConPWD.Top = 3374: .txtConPWD.Left = 3500: .txtConPWD.Visible = True
            
            .lblHelp1.Height = 450: .lblHelp1.Width = 3000: .lblHelp1.Visible = True
            
            .lblHead.Caption = "Create New User": .lblHead.Visible = True: .txtFullName.SetFocus
            .lbllogout.Visible = True: .lblDisplayUID = "System Login as " & LoginUID: .lblDisplayUID.Visible = True
            .Image2.Visible = False 'Arrow Image
            
            .ImgSubmit_Dis.Visible = False: .ImgCancel_Dis.Visible = False
            .ImgSubmit.Visible = False: .ImgCancel.Visible = False
        '----------For Log in-------------------------------
        ElseIf GrpOpt = "Login" Then
            .lblNewUser.Top = 1928: .lblNewUser.Left = 2500: .lblNewUser.Caption = "Select User ID : ": .lblNewUser.Visible = True 'use for user selection.
            .CmbUID.Top = 1928: .CmbUID.Left = 4500: .CmbUID.Visible = True
            .lblPWD.Top = 2410: .lblPWD.Left = 2500: .lblPWD.Visible = True 'use for login user password.
            .txtLoginPWD.Top = 2410: .txtLoginPWD.Left = 4500: .txtLoginPWD.Visible = True
            .lblHelp1.Height = 859: .lblHelp1.Width = 4815: .lblHelp1.Top = 3300: .lblHelp1.Left = 3000
            .lblHelp1.Visible = True: .Image2.Visible = False 'Arrow Image
            .ImgSubmit_Dis.Visible = False: .ImgCancel_Dis.Visible = False
            .ImgSubmit.Visible = False: .ImgCancel.Visible = False
            Call GetLogin 'use to login the Address Book System.
            frmMain.CmbUID.SetFocus 'Setfocus to User ID Combo Box.
        '----------For Welcome Screen-----------------------
        ElseIf GrpOpt = "Welcome" Then 'USE 4 WELCOME SCREEN
            .lblwelcome.Left = 3900: .lblwelcome.Top = 800: .lblwelcome.Visible = True
            .lblTo.Left = 4450: .lblTo.Top = 1250: .lblTo.Visible = True
            .lbl0.Left = 3500: .lbl0.Top = 1700: .lbl0.Visible = True
            .lbl1.Left = 2800: .lbl1.Top = 2000: .lbl1.Visible = True
            .lbl2.Left = 1800: .lbl2.Top = 2700: .lbl2.Visible = True
            .lbl3.Left = 2400: .lbl3.Top = 4000: .lbl3.Visible = True
            .lblContact.Left = 3200: .lblContact.Top = 4500: .lblContact.Visible = True
            .Image2.Visible = False 'Arrow Image
            .ImgSubmit_Dis.Visible = False: .ImgCancel_Dis.Visible = False
            .ImgSubmit.Visible = False: .ImgCancel.Visible = False
            .lbl2.Caption = "This system completly helps / supports the Address Book Options it will manage the all users's data with securities.  Thanks to every one who helps me in the field of Programming espacialy to Almighty GOD, the owner of the whole world. We all must be very thankfull to HIM."
            .lbl3.Caption = "Also Thanks to Taxmaster for background."
        '----------For First Step--------------------------
        ElseIf GrpOpt = "First Step" Then 'USE 4 FIRST STEP AFTER LOGIN.
            .lblwelcome.Left = 4250: .lblwelcome.Top = 1500: .lblwelcome.Visible = True
            .lblTo.Left = 4800: .lblTo.Top = 2050: .lblTo.Visible = True
            
            .lbl0.Width = 3095: .lbl1.Width = 4795: .lbl0.FontSize = 12: .lbl1.FontSize = 12
            .lbl0.Left = 3700: .lbl0.Top = 2500: .lbl0.Visible = True
            .lbl1.Left = 2800: .lbl1.Top = 2900: .lbl1.Visible = True
            
            .lblContact.Left = 3200: .lblContact.Top = 4700: .lblContact.Visible = True
            .Image2.Visible = False 'Arrow Image
            .lbl2.Caption = "This system completly helps / supports the Address Book Options it will manage the all users's data with securities.  Thanks to every one who helps me in the field of Programming espacialy to Almighty GOD, the owner of the whole world. We all must be very thankfull to HIM."
            .lbl3.Caption = "Also Thanks to Taxmaster for background."
            
            .lblHello.Caption = "HELLO   Mr. " & LoginName: .lblUserName.Caption = LoginName
            .lblHello.Visible = True: .lblUserName.Visible = True: .lbllogout.Visible = True
            .lblDisplayUID = "System Login as " & LoginUID: .lblDisplayUID.Visible = True
            
            .lbl1st.Top = 3500: .lbl1st.Left = 3200: .lbl1st.Visible = True 'use for First Running System.

            .ImgSubmit.Visible = False: .ImgSubmit_Dis.Visible = False
            .ImgCancel.Visible = False: .ImgCancel_Dis.Visible = False
        '----------For Contact--------------------------
        ElseIf GrpOpt = "Contact" Then 'USE 4 CONTACT ME.
            .lblMyCaption.Left = 2250: .lblMyCaption.Top = 1400: .lblMyCaption.Visible = True
            .ImgMyPic.Left = 2250: .ImgMyPic.Top = 1900: .ImgMyPic.Visible = True
            
            .ImgButter.Left = 7200: .ImgButter.Top = 1600: .ImgButter.Visible = True
            
            .lblMyName.Left = 3600: .lblMyName.Top = 1900: .lblMyName.Visible = True
            .lblMyDesignation.Left = 3600: .lblMyDesignation.Top = 2200: .lblMyDesignation.Visible = True
            .lblMyContact.Left = 4000: .lblMyContact.Top = 3000: .lblMyContact.Visible = True
            .lblMyComment.Left = 2250: .lblMyComment.Top = 3900: .lblMyComment.Visible = True
            
            .Image2.Visible = False 'Arrow Image
            .lblDisplayUID = "System Login as " & LoginUID: .lblDisplayUID.Visible = True
            
            .lblHello.Caption = "HELLO   Mr. " & LoginName: .lblUserName.Caption = LoginName
            .lblHello.Visible = True: .lblUserName.Visible = True: .lbllogout.Visible = True

            .ImgSubmit.Visible = False: .ImgSubmit_Dis.Visible = False
            .ImgCancel.Visible = False: .ImgCancel_Dis.Visible = False
        End If
    End With: Call FillCombo 'use 4 filling the Grid Boxes
End Sub

Public Sub InitGrid(Grd As MSHFlexGrid)
    With Grd
        .Clear: .ClearStructure
        .Rows = 2
        .FixedRows = 1: .FixedCols = 2
        
        If ViewOpt = "New" Then
            .Cols = 7: .ColSel = 6
            
            'Initialize the column size
            .ColWidth(0) = 315: .ColWidth(1) = 315
            .ColWidth(2) = 600: .ColWidth(3) = 1000
            .ColWidth(4) = 1500: .ColWidth(5) = 1500
            .ColWidth(6) = 2500

            'Initialize the column name
            .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = ""
            .TextMatrix(0, 2) = "CITY #": .TextMatrix(0, 3) = "COUNTRY #"
            .TextMatrix(0, 4) = "LAND LINE #": .TextMatrix(0, 5) = "CELL #"
            .TextMatrix(0, 6) = "E-MAIL"
            
            'Set the column alignment
            .ColAlignment(0) = vbLeftJustify: .ColAlignment(1) = vbLeftJustify
            .ColAlignment(2) = vbCenter: .ColAlignment(3) = vbLeftJustify
            .ColAlignment(4) = vbCenter: .ColAlignment(5) = vbLeftJustify
            .ColAlignment(6) = vbCenter
            
        ElseIf ViewOpt = "Country" Then
            .Cols = 4: .ColSel = 2
            
            'Initialize the column size
            .ColWidth(0) = 315: .ColWidth(1) = 315:
            .ColWidth(2) = 1500: .ColWidth(3) = 3550
    
            'Initialize the column name
            .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = ""
            .TextMatrix(0, 2) = "CODE #": .TextMatrix(0, 3) = "NAME"
            
            'Set the column alignment
            .ColAlignment(0) = vbLeftJustify: .ColAlignment(1) = vbLeftJustify
            .ColAlignment(2) = vbCenter: .ColAlignment(3) = vbLeftJustify
            
        ElseIf ViewOpt = "City" Then
            .Cols = 5: .ColSel = 4
            
            'Initialize the column size
            .ColWidth(0) = 315: .ColWidth(1) = 315:
            .ColWidth(2) = 1500: .ColWidth(3) = 1500
            .ColWidth(4) = 2150
            
            'Initialize the column name
            .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = ""
            .TextMatrix(0, 2) = "COUNTRY CODE": .TextMatrix(0, 3) = "CITY CODE"
            .TextMatrix(0, 4) = "CITY NAME"
            
            'Set the column alignment
            .ColAlignment(0) = vbLeftJustify: .ColAlignment(1) = vbLeftJustify
            .ColAlignment(2) = vbCenter: .ColAlignment(3) = vbLeftJustify
            .ColAlignment(4) = vbCenter
        End If
    End With
End Sub

Public Sub FillGrid(frm As Form, Grd As MSHFlexGrid)
    IntI = 0
    If ViewOpt = "Country" Then
        With RstContry
            .Close: .Open "SELECT * FROM " & tbl_country
            If .RecordCount > 0 Then 'if Database has Records
                For IntI = 1 To .RecordCount
                    Grd.Rows = Grd.Rows + 1
                    Grd.TextMatrix(IntI, 2) = .Fields(0).Value
                    Grd.TextMatrix(IntI, 3) = .Fields(1).Value
                    Grd.Row = Grd.Rows - 1
                    If .EOF = True Then Exit For
                    If .EOF = False Then .MoveNext
                Next
            End If
        End With
        
    ElseIf ViewOpt = "City" Then
        With RstCity
            .Close: .Open "SELECT * FROM " & tbl_city
            If .RecordCount > 0 Then 'if Database has Records
                For IntI = 1 To .RecordCount
                    Grd.Rows = Grd.Rows + 1
                    Grd.TextMatrix(IntI, 2) = .Fields(1).Value
                    Grd.TextMatrix(IntI, 3) = .Fields(0).Value
                    Grd.TextMatrix(IntI, 4) = .Fields(2).Value
                    Grd.Row = Grd.Rows - 1
                    If .EOF = True Then Exit For
                    If .EOF = False Then .MoveNext
                Next
            End If
        End With
        
    ElseIf ViewOpt = "New" Then
        With frmMain.Grid 'Add to grid
            For IntI = 1 To .Rows - 1
                If ((.Rows > 2) And (.TextMatrix(IntI, 4) = frmMain.txtPhone.Text)) Then
                    Msg_Responce = MsgBox("Current entry already exist in the List." & vbCrLf & _
                        "Do you want to replace it.", vbExclamation + vbYesNo, _
                        "Error! Record Existance"): Exit For
                End If
            Next
            If Msg_Responce = vbYes Then
                .TextMatrix(IntI, 2) = frmMain.lblCntCode
                .TextMatrix(IntI, 3) = frmMain.lblCitCode
                .TextMatrix(IntI, 4) = frmMain.txtPhone.Text
                .TextMatrix(IntI, 5) = frmMain.txtMobile.Text
                .TextMatrix(IntI, 6) = frmMain.txtEmail.Text
                 Msg_Responce = 0
            ElseIf Msg_Responce = vbNo Then
                Msg_Responce = 0: Exit Sub
            End If
            If ((frmMain.txtName.Text <> "") And (frmMain.txtPhone.Text <> "")) Then
            If .Rows = 2 And .TextMatrix(1, 2) = "" Then 'Perform if the grid is empty.
                .TextMatrix(1, 2) = frmMain.lblCntCode
                .TextMatrix(1, 3) = frmMain.lblCitCode
                .TextMatrix(1, 4) = frmMain.txtPhone.Text
                .TextMatrix(1, 5) = frmMain.txtMobile.Text
                .TextMatrix(1, 6) = frmMain.txtEmail.Text
            Else 'if grid has at least record.
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 2) = frmMain.lblCntCode
                .TextMatrix(.Rows - 1, 3) = frmMain.lblCitCode
                .TextMatrix(.Rows - 1, 4) = frmMain.txtPhone.Text
                .TextMatrix(.Rows - 1, 5) = frmMain.txtMobile.Text
                .TextMatrix(.Rows - 1, 6) = frmMain.txtEmail.Text
                .Row = .Rows - 1
            End If
            End If
        End With
    End If
    Call Clk_Grid(frmMain) 'for grid revome button
End Sub

Public Sub Clk_Grid(frm As Form)
    With frm.Grid
        If .Rows = 2 And .TextMatrix(1, 2) = "" Then
            frm.BtnRemove.Visible = False
        Else
            frm.BtnRemove.Visible = True
            frm.BtnRemove.Top = (.CellTop + .Top) - 20
            frm.BtnRemove.Left = .Left + 50
        End If
    End With
End Sub

Public Sub LoginAction(frm As Form, Rst As Recordset)
    With frm
        .CmbUID.Clear
        .CmbUID.AddItem "Select User"
        .CmbUID.AddItem "Administrator"
        If Rst.RecordCount > 0 Then
            For IntI = 1 To Rst.RecordCount
                .CmbUID.AddItem Rst.Fields(0).Value
            Next
        End If
        .CmbUID.Text = "Select User"
        .txtLoginPWD.SetFocus
    End With
End Sub

Public Sub NumOnly(Asc As Integer)
Char = Chr(Asc)
    If Char Like "[0-9,.,__,.,/]" Or Asc = 8 Or Asc = 32 Or Asc = 13 Then
    Else: Asc = 0
    End If
End Sub

Public Sub CharOnly(Asc As Integer)
    Char = Chr(Asc)
    If Char Like "[a-z,.,A-Z,.,__,.,--,.,/]" Or Asc = 8 Or Asc = 32 Or Asc = 13 Then
    Else: Asc = 0
    End If
End Sub

Public Sub Alpha_Char(Asc As Integer, frm As Form, Txt As TextBox)
    If ((Asc <> 13) And (Asc <> 8)) Then
        If Len(Txt) = 0 Then Chrtxt = Asc: Asc = 0: _
           Txt = UCase(Chr(Chrtxt)): SendKeys "{End}": Chrtxt = ""
        If Chrtxt = 32 Then txtVal = Txt: Chrtxt = Asc: Asc = 0: _
           Txt = txtVal & UCase(Chr(Chrtxt)): SendKeys "{End}": Chrtxt = ""
        If ((Asc = 32) And (Asc <> 8)) Then Chrtxt = Asc
    End If
End Sub

Public Sub GetID(Rst As Recordset, tbl As String)
    With Rst
        .Close: .Open "SELECT * FROM " & tbl
        
        If ViewOpt = "City" Then RECID = "CI-" & .RecordCount + 1
        If ViewOpt = "Country" Then RECID = "CN-" & .RecordCount + 1
        If ViewOpt = "New" Then RECID = "AB-" & .RecordCount + 1
    End With
End Sub

Public Sub FillCombo()
    frmMain.cmbcountry.Clear
    frmMain.cmbcity.Clear
    
    With RstContry
        .Close: .Open "SELECT * FROM " & tbl_country
        frmMain.cmbcountry.AddItem "Select Country"
        For IntI = 1 To .RecordCount
            frmMain.cmbcountry.AddItem .Fields(1).Value
            If .EOF = True Then Exit For
            If .EOF = False Then .MoveNext
        Next
    End With: frmMain.cmbcountry.Text = "Select Country"
    
    With RstCity '
        .Close: .Open "SELECT * FROM " & tbl_city
        frmMain.cmbcity.AddItem "Select City"
        For IntI = 1 To .RecordCount
            frmMain.cmbcity.AddItem .Fields(2).Value
            If .EOF = True Then Exit For
            If .EOF = False Then .MoveNext
        Next
    End With: frmMain.cmbcity.Text = "Select City"
End Sub

Public Sub SeekRecord()
    With frmMain
        If .txtSearName.Text <> "" Then
            RstAB.Close: RstAB.Open "SELECT * FROM " & tbl_AB & " WHERE AB_Name='" & .txtSearName.Text & "'"
            RstABD.Close: RstABD.Open "SELECT * FROM " & tbl_ABD & " WHERE AB_ID='" & RstAB.Fields(0).Value & "'"
                        
        ElseIf .txtSearLandLine.Text <> "" Then
            RstABD.Close: RstABD.Open "SELECT * FROM " & tbl_ABD & " WHERE AB_LandLine='" & .txtSearLandLine.Text & "'"
            RstAB.Close: RstAB.Open "SELECT * FROM " & tbl_AB & " WHERE AB_ID='" & RstABD.Fields(0).Value & "'"
            
        ElseIf .txtSearCell.Text <> "" Then
            RstABD.Close: RstABD.Open "SELECT * FROM " & tbl_ABD & " WHERE AB_Mobile='" & .txtSearCell.Text & "'"
            RstAB.Close: RstAB.Open "SELECT * FROM " & tbl_AB & " WHERE AB_ID='" & RstABD.Fields(0).Value & "'"
            
        ElseIf .txtSearName.Text = "" And .txtSearLandLine.Text And .txtSearCell.Text = "" Then
            MsgBox "Invalid Search Criteria." & vbCrLf & _
                   "Please enter atleast entry for search.", vbCritical, "Error! Search Criteria..."
                   SendKeys "{Home}+{End}": .txtSearName.SetFocus
        End If
    End With: Call FillList 'use 2 fill list by seached record
    If RstAB.RecordCount > 0 Then frmMain.lblSearchedName.Visible = True: frmMain.lblSearchedName = "Searched Name :  " & RstAB.Fields(1).Value
    If RstAB.RecordCount <= 0 Then frmMain.lblSearchedName.Visible = False: frmMain.lblSearchedName = ""
End Sub

Public Sub FillList()
    With RstABD
        If .RecordCount > 0 Then
            frmMain.LstSearch.ListItems.Clear 'use to clear the listview.
            For IntI = 1 To RstABD.RecordCount
                Set LItem = frmMain.LstSearch.ListItems.Add(IntI, , IntI, 2, 2)
                    RstContry.Close: RstContry.Open "SELECT * FROM " & tbl_country & " WHERE Country_Code='" & RstABD.Fields(1).Value & "'"
                    RstCity.Close: RstCity.Open "SELECT * FROM " & tbl_city & " WHERE City_Code='" & RstABD.Fields(2).Value & "'"
                    
                    LItem.SubItems(1) = RstCity.Fields(2).Value
                    LItem.SubItems(2) = .Fields(3).Value
                    LItem.SubItems(3) = .Fields(4).Value
                    LItem.SubItems(4) = .Fields(5).Value
                    If RstABD.EOF = True Then Exit For
                    If RstABD.EOF = False Then RstABD.MoveNext
            Next
            frmMain.lblSearAdd = RstAB.Fields(2).Value
            frmMain.lblSearCity = RstCity.Fields(2).Value
            frmMain.lblSearCountry = RstContry.Fields(1).Value
            
            frmMain.Frame_Contacts.Visible = True 'use 2 view the frame to display the record.
            frmMain.Frame_Postal.Visible = True 'use 2 view the frame to display the record.
            frmMain.lblSearAddCap.Visible = True: frmMain.lblSearAdd.Visible = True
            frmMain.lblSearCity.Visible = True: frmMain.lblSearCountry.Visible = True
            
        ElseIf .RecordCount <= 0 Then
            frmMain.Frame_Postal.Visible = False 'use 2 Hide the frame if no record found.
            frmMain.lblSearAddCap.Visible = False: frmMain.lblSearAdd.Visible = False
            frmMain.lblSearCity.Visible = False: frmMain.lblSearCountry.Visible = False
            
            MsgBox "Record not matched due to..." & vbCrLf & _
                   "1: Differ Spell." & vbCrLf & "2: Non Existance in Database.", vbExclamation, "Error! Record Not Exist..."
                   SendKeys "{Home}+{End}": frmMain.txtSearName.SetFocus
        End If
    End With
    RstContry.Close: RstContry.Open "SELECT * FROM " & tbl_country
    RstCity.Close: RstCity.Open "SELECT * FROM " & tbl_city
End Sub

Public Sub set_menubtn_logout()
    With frmMain
        .ImgHome.Visible = False: .ImgAbout.Visible = False
        .ImgCountry.Visible = False: .ImgCity.Visible = False
        .ImgNewRec.Visible = False: .ImgSearch.Visible = False
        .ImgContact.Visible = False

        .ImgSubmit.Visible = False: .ImgCancel.Visible = False
        .ImgSubmit_Dis.Visible = False: .ImgCancel_Dis.Visible = False
    End With
End Sub

'use 4 getting the specific record 4rm table.
Public Sub Get_Combo_Rec(Rs As ADODB.Recordset, tbl As String, SearFld As String, FldNo As Integer, cmb As ComboBox)
    cmb.Clear
    Rs.Close: Rs.Open "SELECT * FROM " & tbl & " WHERE " & SearFld & "='" & frmMain.lblCntCode & "'"
    If cmb.List(0) <> "Selection" Then cmb.AddItem "Selection"
    With Rs
        If Rs.RecordCount > 0 Then
            .MoveFirst ': Cmb.Clear
            Do While Not .EOF
                If IsNull(.Fields(FldNo).Value) = False Then _
                    cmb.AddItem .Fields(FldNo).Value
                .MoveNext
            Loop
        ElseIf Rs.RecordCount <= 0 Then
            cmb.AddItem "Record not Exist"
        End If
    End With: cmb.Text = "Selection"
End Sub

