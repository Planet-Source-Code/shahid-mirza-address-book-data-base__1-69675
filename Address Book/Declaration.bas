Attribute VB_Name = "Declaration"
Global DB_Con As New ADODB.Connection 'use for database connection.
Global RstSQL As String 'use for SQL Statment.
Global LItem As ListItem 'use 4 List
Global Chrtxt  'For Alpha Character.
Global msgRes 'For msgBox
Global oCtrl As Control: Global IntI As Integer: Global IntTimer As Integer
Global ViewOpt As String: Global LoginCounter As Integer
Global LoginUID As String: Global LoginName As String: Global ChkPWD As String
Global RECID As String  'use 4 Auto ID Generates
Global CntCODE As String '4 Country Code
Global CitCODE As String '4 City Code
Global GetLstRec As String '4 Get Record From List View.

Public Const tbl_login = "tbl_Login"
Public Const pwd = "absystem"
Public Const tbl_NUser = "tbl_Create_Login"
Public Const tbl_country = "tbl_Country"
Public Const tbl_city = "tbl_City"

Public Const tbl_AB = "tbl_AddBook"
Public Const tbl_ABD = "tbl_AddBook_Detail"

