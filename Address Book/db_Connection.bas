Attribute VB_Name = "db_Connection"
Public Sub Main()
    DB_Con.ConnectionString = "Provider=Microsoft.Jet.oledb.4.0"
    DB_Con.Open App.Path & ("\db_addressbook.mdb")
    Call GetDBase 'use to initialize the Data Base Tables.
    Load frmMain: frmMain.Show vbModal
End Sub

