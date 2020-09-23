VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   7560
   ClientLeft      =   4845
   ClientTop       =   1710
   ClientWidth     =   9300
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7591.424
   ScaleMode       =   0  'User
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame_Contacts 
      BackColor       =   &H8000000D&
      Caption         =   "Contact Numbers"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2040
      Left            =   1800
      TabIndex        =   64
      Top             =   3255
      Width           =   6735
      Begin MSComctlLib.ListView LstSearch 
         Height          =   1740
         Left            =   75
         TabIndex        =   65
         Top             =   225
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   3069
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   65535
         BackColor       =   -2147483635
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SR #"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "AREA #"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "LAND LINE"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "MOBILE"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "E-MAIL"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.Frame Frame_Postal 
      BackColor       =   &H8000000D&
      Caption         =   "Postal Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Left            =   1800
      TabIndex        =   59
      Top             =   2160
      Width           =   6735
      Begin VB.Label lblSearAddCap 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address : "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblSearCountry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblSearCountry"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   62
         Top             =   760
         Width           =   2295
      End
      Begin VB.Label lblSearCity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblSearCity"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   61
         Top             =   760
         Width           =   2295
      End
      Begin VB.Label lblSearAdd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblSearAdd"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1680
         TabIndex        =   60
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.TextBox txtSearName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   6120
      TabIndex        =   54
      Text            =   "txtSearName"
      Top             =   8520
      Width           =   1815
   End
   Begin VB.TextBox txtSearLandLine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   6120
      TabIndex        =   53
      Text            =   "txtSearLandLine"
      Top             =   8520
      Width           =   1815
   End
   Begin VB.TextBox txtSearCell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   6120
      TabIndex        =   52
      Text            =   "txtSearCell"
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   0
      Top             =   6600
   End
   Begin VB.TextBox txtFullName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   9960
      TabIndex        =   35
      Text            =   "txtFullName"
      Top             =   720
      Width           =   2175
   End
   Begin VB.ComboBox CmbUID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   9960
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtLoginPWD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   9960
      PasswordChar    =   "*"
      TabIndex        =   33
      Text            =   "txtLoginPWD"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtConPWD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   9960
      PasswordChar    =   "*"
      TabIndex        =   32
      Text            =   "txtConPWD"
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtPWD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   9960
      PasswordChar    =   "*"
      TabIndex        =   31
      Text            =   "txtPWD"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtNewUser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   9960
      TabIndex        =   30
      Text            =   "txtNewUser"
      Top             =   1200
      Width           =   2175
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   120
      Top             =   8280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":164A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":496E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6300
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C92
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9624
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AFB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C948
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E2DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FC6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1094A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1122A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11F06
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12BE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":138BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1459A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15276
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton CmdEnd 
      Height          =   229
      Left            =   8600
      TabIndex        =   24
      Top             =   199
      Width           =   230
      _ExtentX        =   397
      _ExtentY        =   397
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   15591915
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   99
      MICON           =   "frmMain.frx":15B52
      ALIGN           =   1
      IMGLST          =   "itb32x32"
      IMGICON         =   "7"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.TextBox txtCountryCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   6120
      TabIndex        =   22
      Text            =   "txtCountryCode"
      Top             =   8160
      Width           =   1455
   End
   Begin VB.TextBox txtCityCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   6120
      TabIndex        =   21
      Text            =   "txtCityCode"
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox txtcity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   6120
      TabIndex        =   20
      Text            =   "txtcity"
      Top             =   7080
      Width           =   2535
   End
   Begin VB.ComboBox cmbcity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2370
      Width           =   1880
   End
   Begin VB.ComboBox cmbcountry 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1920
      Width           =   1880
   End
   Begin VB.TextBox txtCountry 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   6120
      TabIndex        =   17
      Text            =   "txtCountry"
      Top             =   7440
      Width           =   2535
   End
   Begin VB.CommandButton BtnRemove 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   8280
      Picture         =   "frmMain.frx":15E6C
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Remove"
      Top             =   7920
      Visible         =   0   'False
      Width           =   275
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   1713
      Left            =   2805
      TabIndex        =   15
      Top             =   3525
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3016
      _Version        =   393216
      BackColor       =   -2147483645
      ForeColor       =   65280
      Rows            =   0
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      BackColorSel    =   1091552
      BackColorBkg    =   15112595
      BackColorUnpopulated=   16711935
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      GridColorUnpopulated=   -2147483633
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   5
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   2800
      TabIndex        =   0
      Text            =   "txtName"
      Top             =   1494
      Width           =   2775
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   780
      Left            =   2800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmMain.frx":1601E
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txtMobile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   4620
      TabIndex        =   3
      Text            =   "txtMobile"
      Top             =   3187
      Width           =   1695
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   6455
      TabIndex        =   4
      Text            =   "txtEmail"
      Top             =   3187
      Width           =   2260
   End
   Begin VB.TextBox txtPhone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   2800
      TabIndex        =   2
      Text            =   "txtPhone"
      Top             =   3187
      Width           =   1695
   End
   Begin VB.TextBox txtID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   4920
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "txtID"
      Top             =   960
      Width           =   1695
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   7680
      Top             =   7800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16029
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16A3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1744D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":177E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17B81
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17F1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":182B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18CC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":196D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A0EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AAFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B50F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BF21
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C933
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CECF
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D46B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24644
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":280AB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image ImgButter 
      Height          =   960
      Left            =   10080
      Picture         =   "frmMain.frx":2B726
      Top             =   6600
      Width           =   960
   End
   Begin VB.Label lblMyComment 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":2CD70
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9720
      TabIndex        =   71
      Top             =   4800
      Width           =   4935
   End
   Begin VB.Image ImgMyPic 
      Height          =   1635
      Left            =   9720
      Picture         =   "frmMain.frx":2CE6F
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1200
   End
   Begin VB.Label lblMyContact 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "methoomirza@hotmail.com methoomirza@yahoo.com  Contact # : +923006410758"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   11520
      TabIndex        =   70
      Top             =   3960
      Width           =   2745
   End
   Begin VB.Label lblMyDesignation 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Senior Dataprocessor and System Developer."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   11040
      TabIndex        =   69
      Top             =   3360
      Width           =   3345
   End
   Begin VB.Label lblMyCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Developed and Designed"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10050
      TabIndex        =   68
      Top             =   2760
      Width           =   3960
   End
   Begin VB.Label lblMyName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mr. Muhammad Shahid Mughal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   11040
      TabIndex        =   67
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Image ImgContact 
      Height          =   255
      Left            =   360
      MouseIcon       =   "frmMain.frx":39655
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":3995F
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Image ImgAbout 
      Height          =   255
      Left            =   360
      MouseIcon       =   "frmMain.frx":3CB6C
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":3CE76
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Image ImgHome 
      Height          =   255
      Left            =   360
      MouseIcon       =   "frmMain.frx":3FFC0
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":402CA
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblSearchedName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblSearchedName"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   6720
      TabIndex        =   66
      Top             =   8520
      Width           =   1905
   End
   Begin VB.Image ImgFind 
      Height          =   255
      Left            =   7080
      MouseIcon       =   "frmMain.frx":433F4
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":436FE
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label lblclick 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click Submit to Save Record"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   480
      TabIndex        =   58
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label lblByLandLine 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Land Line #"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   6120
      TabIndex        =   57
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Label lblByMobile 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile #"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   6120
      TabIndex        =   56
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Label lblbyName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   6120
      TabIndex        =   55
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Image ImgSubmit 
      Height          =   255
      Left            =   360
      MouseIcon       =   "frmMain.frx":46CFF
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":47009
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Image ImgCancel 
      Height          =   255
      Left            =   1515
      MouseIcon       =   "frmMain.frx":4A32F
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":4A639
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label lblCitCode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblCitCode"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   51
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Label lblCntCode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblCntCode"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   50
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Label lbl1st 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lbl1st"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4320
      TabIndex        =   49
      Top             =   8880
      Width           =   4095
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblUserName"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   4320
      TabIndex        =   48
      Top             =   660
      Width           =   2295
   End
   Begin VB.Label lblHello 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblWelcome2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   47
      Top             =   660
      Width           =   1725
   End
   Begin VB.Label lbllogout 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   7680
      MouseIcon       =   "frmMain.frx":4D994
      MousePointer    =   99  'Custom
      TabIndex        =   46
      Top             =   199
      Width           =   765
   End
   Begin VB.Label lblHelp1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblHelp1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1080
      TabIndex        =   45
      Top             =   9720
      Width           =   4815
   End
   Begin VB.Label lblContact 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact : methoomirza@hotmail.com                 (+923006410758)."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Left            =   1440
      TabIndex        =   44
      Top             =   9360
      Width           =   3615
   End
   Begin VB.Label lblDisplayUID 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblDisplayUID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   43
      Top             =   239
      Width           =   3975
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Also Thanks to Taxmaster for background."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   960
      TabIndex        =   42
      Top             =   9000
      Width           =   5055
   End
   Begin VB.Label lblFullName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10080
      TabIndex        =   36
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The Telephone and Address maintain System"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   41
      Top             =   7440
      Width           =   4095
   End
   Begin VB.Label lbl0 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Address Book System - (ABS)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   40
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Label lblTo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   39
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label lblwelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   38
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":4DC9E
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   840
      TabIndex        =   37
      Top             =   7800
      Width           =   6615
   End
   Begin VB.Label lblNewUser 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter New User : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   9960
      TabIndex        =   27
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblConPWD 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   9960
      TabIndex        =   29
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblPWD 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   9960
      TabIndex        =   28
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblOR2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8280
      TabIndex        =   26
      Top             =   8160
      Width           =   315
   End
   Begin VB.Label lblOR1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8280
      TabIndex        =   25
      Top             =   8160
      Width           =   315
   End
   Begin VB.Image ImgCancel_Dis 
      Enabled         =   0   'False
      Height          =   255
      Left            =   1520
      MouseIcon       =   "frmMain.frx":4DDB2
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":4E0BC
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Image ImgSubmit_Dis 
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      MouseIcon       =   "frmMain.frx":51500
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":5180A
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label lblHead 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblHead"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   7680
      TabIndex        =   23
      Top             =   600
      Width           =   675
   End
   Begin VB.Image ImgSearch 
      Height          =   255
      Left            =   360
      MouseIcon       =   "frmMain.frx":54C50
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":54F5A
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Image ImgNewRec 
      Height          =   255
      Left            =   360
      MouseIcon       =   "frmMain.frx":5849D
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":587A7
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image ImgCity 
      Height          =   255
      Left            =   360
      MouseIcon       =   "frmMain.frx":5BC8C
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":5BF96
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Image ImgCountry 
      Height          =   255
      Left            =   360
      MouseIcon       =   "frmMain.frx":5F41E
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":5F728
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   7800
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Label lblLandline 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Land Line"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2800
      TabIndex        =   14
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1680
      TabIndex        =   13
      Top             =   1524
      Width           =   1095
   End
   Begin VB.Label lblCell 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cell "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4620
      TabIndex        =   11
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6455
      TabIndex        =   10
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblNumber 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Numbers "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1550
      TabIndex        =   9
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblCountry 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Country : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5760
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblCity 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5760
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblAddress 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1680
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblID 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Record ID # : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3480
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   5895
      Left            =   60
      Picture         =   "frmMain.frx":62BDD
      Stretch         =   -1  'True
      Top             =   60
      Width           =   9015
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderStyle     =   2  'Dash
      Height          =   6005
      Left            =   15
      Top             =   15
      Width           =   9120
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   3735
      Left            =   720
      Top             =   6480
      Width           =   8055
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   2055
      Left            =   9840
      Top             =   600
      Width           =   2415
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   3255
      Left            =   9600
      Top             =   2760
      Width           =   5055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbcity_Click()
    RstCity.Close: RstCity.Open "SELECT * FROM " & tbl_city & " WHERE City_Name='" & cmbcity.Text & "'"
    If RstCity.RecordCount > 0 Then lblCitCode = RstCity.Fields(0).Value
    If RstCity.RecordCount <= 0 Then lblCitCode = ""
    RstCity.Close: RstCity.Open "SELECT * FROM " & tbl_city
End Sub

Private Sub cmbcity_KeyPress(KeyAscii As Integer)
    Dim CntErr As String
    If KeyAscii = 13 Then
        If cmbcity.Text <> "Select City" Then txtPhone.SetFocus
        If cmbcity.Text = "Record not Exist" Then
            MsgBox "Please verify / add the City information" & vbCrLf & _
                   "For Smoothly Execution....", vbExclamation, "City Records"
                   CntErr = cmbcountry
                   Call Setting_Grp(frmMain, "City")
                   cmbcountry.Text = CntErr
        End If
    End If
End Sub

Private Sub cmbcountry_Click()
    RstContry.Close: RstContry.Open "SELECT * FROM " & tbl_country & " WHERE Country_Name='" & cmbcountry.Text & "'"
    If RstContry.RecordCount > 0 Then lblCntCode = RstContry.Fields(0).Value: RstCity.Close: RstCity.Open "SELECT * FROM " & tbl_city & " WHERE Country_Code='" & lblCntCode & "'"
    If RstContry.RecordCount <= 0 Then lblCntCode = ""
    RstContry.Close: RstContry.Open "SELECT * FROM " & tbl_country
End Sub

Private Sub cmbcountry_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbcountry.Text <> "Select Country" Then
            If ViewOpt = "City" Then
                txtCityCode.SetFocus
            ElseIf ViewOpt = "New" Then
                Call Get_Combo_Rec(RstCity, tbl_city, "Country_Code", 2, cmbcity) 'use 2 get the specific city for selected country.
                cmbcity.SetFocus
            End If
        End If
    End If
End Sub

Private Sub cmbcountry_LostFocus()
    If cmbcountry.Text <> "Select Country" Then
        If ViewOpt = "City" Then
            txtCityCode.SetFocus
        ElseIf ViewOpt = "New" Then
            Call Get_Combo_Rec(RstCity, tbl_city, "Country_Code", 2, cmbcity) 'use 2 get the specific city for selected country.
            cmbcity.SetFocus
        End If
    End If
End Sub

Private Sub CmbUID_Click()
    If CmbUID.Text = "Administrator" Then
        lblHelp1.Caption = " Enter Default Administrator Password. If not Please contact at methoomirza@hotmail.com OR +923006410758."
    ElseIf CmbUID.Text <> "Administrator" Then
        lblHelp1.Caption = " Enter Password of selected User ID..."
    End If: LoginCounter = 0 'Initialize the Login Counter whenever select the New User for Login.

    If CmbUID.Text = "Select User" Then _
        lblHelp1.Caption = " Select valid User ID and enter Password of selected User ID for Successfull Login....."

End Sub

Private Sub CmbUID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmbUID.Text <> "Select User" Then txtLoginPWD.SetFocus
        If CmbUID.Text = "Select User" Then MsgBox "Please select the User ID." & vbCrLf & _
            "You can't login without valid User ID", vbCritical, "Err0r! User Selection...": CmbUID.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call frmSetting(frmMain) 'for initializing the form.
    Call set_menubtn(frmMain): Call Setting_Grp(frmMain, "Welcome")
End Sub

Private Sub Grid_Click()
    Call Clk_Grid(frmMain) 'use form Grid remove button.
End Sub

Private Sub ImgAbout_Click()
    Call Setting_Grp(frmMain, "Welcome")
    lblHello.Visible = True: lbllogout.Visible = True
    lblHello.Caption = "HELLO   Mr. " & LoginName: lblUserName.Caption = LoginName: lblUserName.Visible = True
    lblDisplayUID = "System Login as " & LoginUID: lblDisplayUID.Visible = True
End Sub

Private Sub ImgCancel_Click()
    Call frmSetting(frmMain)
End Sub

Private Sub ImgCity_Click()
    Call Setting_Grp(frmMain, "City")
End Sub

Private Sub ImgContact_Click()
    Call Setting_Grp(frmMain, "Contact")
End Sub

Private Sub ImgCountry_Click()
    Call Setting_Grp(frmMain, "Country")
End Sub

Private Sub ImgFind_Click()
    lblclick.Visible = False
    Call SeekRecord 'use 2 find the record.
    If LstSearch.Visible = True Then LstSearch.SetFocus
End Sub

Private Sub ImgHome_Click()
    Call Setting_Grp(frmMain, "First Step")
End Sub

Private Sub ImgNewRec_Click()
    Call Setting_Grp(frmMain, "New")
End Sub

Private Sub ImgSearch_Click()
    Call Setting_Grp(frmMain, "Search")
End Sub

Private Sub CmdEnd_Click()
    CloseAll 'use 4 close all database table.
    End
End Sub

Private Sub ImgSubmit_Click()
    Call SaveData   'use 4 saving the Data
    Call FillCombo 'use to fill the combo boxes
End Sub

Private Sub lbllogout_Click()
    Call Setting_Grp(frmMain, "Login") 'use to initialize the form for login.
    Call set_menubtn_logout 'use to Hide the Menu Buttons.
    LoginName = "": LoginUID = ""
End Sub

Private Sub LstSearch_Click()
    For IntI = 1 To LstSearch.ListItems.Count
        Set LItem = LstSearch.ListItems.Item(IntI)
            GetLstRec = LItem.SubItems(2): Exit For
    Next
    MsgBox GetLstRec
End Sub

Private Sub Timer1_Timer()
    For IntTimer = 1 To 10
        If IntTimer = 10 Then Call Setting_Grp(frmMain, "Login") 'Move to Login system
        If IntTimer = 10 Then Timer1.Interval = 0
    Next
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    Call Alpha_Char(KeyAscii, frmMain, txtAddress) 'use for Capital Alpha Character.
    If KeyAscii = 13 Then
        If txtAddress.Text <> "" Then cmbcountry.SetFocus: KeyAscii = 0
    End If
End Sub

Private Sub txtcity_KeyPress(KeyAscii As Integer)
    Call CharOnly(KeyAscii) ' 4 Character Values Only.
    Call Alpha_Char(KeyAscii, frmMain, txtcity) 'for Capital Alpha Characters.
    If KeyAscii = 13 Then
        If txtcity.Text <> "" Then Call ImgSubmit_Click
    End If
End Sub

Private Sub txtCityCode_KeyPress(KeyAscii As Integer)
    Call NumOnly(KeyAscii) 'use only 4 Numeric character.
    If KeyAscii = 13 Then
        If txtCityCode.Text <> "" Then
            txtcity.SetFocus: CitCODE = txtCityCode.Text
            ImgSubmit_Dis.Visible = False: ImgSubmit.Visible = True
            ImgCancel_Dis.Visible = False: ImgCancel.Visible = True
            Call GetID(RstCity, tbl_city): txtID.Text = RECID
        End If
    End If
End Sub

Private Sub txtConPWD_GotFocus()
    lblHelp1.Top = 3370: lblHelp1.Left = 5700: lblHelp1.Caption = "Enter Confirm Password match with Password"
End Sub

Private Sub txtConPWD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtConPWD.Text <> "" Then
            If txtConPWD.Text <> txtPWD.Text Then
                MsgBox "Confirm Password not match." & vbCrLf & _
                       "Please enter valid confirm Password.", vbCritical, "Error! Confirm Password...."
                       SendKeys "{Home}+{End}": txtConPWD.SetFocus
            ElseIf txtConPWD.Text = txtPWD.Text Then
                Call SaveNewUser(RstNUser, RstLogin, frmMain) 'call 4 Creating New User
            End If
        ElseIf txtConPWD.Text = "" Then
            MsgBox "Empty Confirm Password field is not supported." & vbCrLf & _
                   "Please enter the Valid Entry.", vbExclamation, "Error! Empty Confirm Password Field"
                   SendKeys "{Home}+{End}": txtConPWD.SetFocus
        End If
    End If
End Sub

Private Sub txtCountry_KeyPress(KeyAscii As Integer)
    Call CharOnly(KeyAscii) 'use only 4 character values
    Call Alpha_Char(KeyAscii, frmMain, txtCountry) 'for Capital Alpha Characters.
    Call CharOnly(KeyAscii) ' 4 Character Values Only.
    If KeyAscii = 13 Then
        If txtCountry.Text <> "" Then Call ImgSubmit_Click
    End If
End Sub

Private Sub txtCountryCode_KeyPress(KeyAscii As Integer)
    Call NumOnly(KeyAscii) 'use only 4 Numeric character.
    If KeyAscii = 13 Then
        If txtCountryCode.Text <> "" Then
            CntCODE = txtCountryCode.Text: txtCountry.SetFocus
            ImgSubmit_Dis.Visible = False: ImgSubmit.Visible = True
            ImgCancel_Dis.Visible = False: ImgCancel.Visible = True
            Call GetID(RstContry, tbl_country): txtID.Text = RECID
        End If
    End If
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtEmail.Text <> "" Then
            If txtPhone.Text <> "Nil" Or txtMobile.Text <> "Nil" Or txtEmail.Text <> "Nil" Then
                Call FillGrid(frmMain, Grid) 'use 2 fill the Grid.
                txtPhone.Text = "": txtMobile.Text = "": txtEmail.Text = ""
                cmbcity.SetFocus
                If Grid.Rows >= 2 Then
                    ImgSubmit_Dis.Visible = False: ImgSubmit.Visible = True
                    ImgCancel_Dis.Visible = False: ImgCancel.Visible = True
                    lblclick.Caption = "Click Submit to Save Record"
                    lblclick.Visible = True
                End If
            ElseIf txtPhone.Text = "Nil" And txtMobile.Text = "Nil" And txtEmail.Text = "Nil" Then
                MsgBox "Not a valid Contacts." & vbCrLf & "Please atleast one contact number.", vbCritical, "Error! Contacts..."
                txtPhone.SetFocus
            End If
        ElseIf txtEmail.Text = "" Then
            txtEmail.Text = "Nil"
        End If
    End If
End Sub

Private Sub txtFullName_GotFocus()
    lblHelp1.Top = 1928: lblHelp1.Left = 5700: lblHelp1.Caption = "Full Name of New User Account"
End Sub

Private Sub txtFullName_KeyPress(KeyAscii As Integer)
    Call CharOnly(KeyAscii) ' 4 Character Values Only.
    Call Alpha_Char(KeyAscii, frmMain, txtFullName) 'use to First Capital Letter.
    If KeyAscii = 13 Then
        If txtFullName.Text = "" Then txtFullName.Text = "UN-DEFINE"
        txtNewUser.SetFocus
    End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtLoginPWD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LoginUID = CmbUID.Text: LoginName = "DEFAULT" '4 User Name
        If Len(txtLoginPWD.Text) < 6 Then 'use to verify the password length.
            MsgBox "UnAthurise Password length." & vbCrLf & _
                   "Password not less than 6 characters", vbCritical, "Error! In Password...."
                   SendKeys "{Home}+{End}": txtLoginPWD.SetFocus
        ElseIf Len(txtLoginPWD.Text) >= 6 Then 'use to verify the password length if okay.
            If RstLogin.RecordCount <= 0 Then 'if Login User not In Database.
                If CmbUID.Text = "Administrator" Then 'if data not in User Login Database. and UserID is Administrator.
                    If LCase(txtLoginPWD.Text) = LCase(pwd) Then 'If Login Is Administrator Login User Zero(0).
                            Call Setting_Grp(frmMain, "New User")
                    ElseIf LCase(txtLoginPWD.Text) <> LCase(pwd) Then 'if Password not matched.
                        MsgBox "In-correct " & CmbUID.Text & " Password." & vbCrLf & _
                               "Please enter correct Password." & vbCrLf & _
                               "As 'abssytem'", vbExclamation, "InCorrect " & CmbUID.Text & " Password"
                               SendKeys "{Home}+{End}": txtLoginPWD.SetFocus
                    End If
                End If
                
            ElseIf RstLogin.RecordCount > 0 Then 'if User present in Login Database.
                If CmbUID.Text = "Administrator" Then 'if User present in Database and User ID is Administrator.
                        If txtLoginPWD.Text = LCase(pwd) Then
                            lbllogout.Visible = True: Call set_menubtn(frmMain) 'use 4 enable the menu buttons.
                            Call Setting_Grp(frmMain, "First Step") 'use 4 First Screen after Login.
                        ElseIf txtLoginPWD.Text <> LCase(pwd) Then
                            MsgBox "It is not a valid Password." & vbCrLf & _
                                   "Please enter the correct Password." & vbCrLf & _
                                   "As 'absystem'", vbCritical, "Error! In Password"
                                   SendKeys "{Home}+{End}": txtLoginPWD.SetFocus
                        End If
                ElseIf CmbUID.Text <> "Administrator" Then
                    RstLogin.Close: RstLogin.Open "SELECT * FROM " & tbl_login & " WHERE UID='" & LoginUID & "'"
                        ChkPWD = RstLogin.Fields(1).Value 'Assign Password to Variable 4 Comparison/Checking.
                            'Following only use 4 User Name
                            RstNUser.Close: RstNUser.Open "SELECT * FROM " & tbl_NUser & " WHERE UID='" & LoginUID & "'"
                            If RstNUser.RecordCount > 0 Then
                                LoginName = RstNUser.Fields(0).Value '4 User Name
                            ElseIf RstNUser.RecordCount <= 0 Then
                                LoginName = "Un-Registered" '4 User Name
                            End If
                            RstNUser.Close: RstNUser.Open "SELECT * FROM " & tbl_NUser
                            '---------------------------------------------------------
                        If txtLoginPWD.Text = ChkPWD Then
                            lbllogout.Visible = True: Call set_menubtn(frmMain) 'use 4 enable the menu buttons.
                            Call Setting_Grp(frmMain, "First Step") 'use 4 First Screen after Login.
                        ElseIf txtLoginPWD.Text <> ChkPWD Then
                            MsgBox "It is not a valid Password." & vbCrLf & _
                                   "Please enter the correct Password.", vbCritical, "Error! In Password"
                                   SendKeys "{Home}+{End}": txtLoginPWD.SetFocus
                        End If

                End If
            End If
            If RstContry.RecordCount <= 0 Then lbl1st.Caption = "Please first enter countries and related cities before use the system proper and correctly."
            If RstContry.RecordCount > 0 Then lbl1st.Caption = "Now you are ready to maintain the Adress and Telephone records."
        End If
    End If
    RstLogin.Close: RstLogin.Open "SELECT * FROM " & tbl_login
End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)
    Call NumOnly(KeyAscii) 'Only 4 Numeric Character.
    If KeyAscii = 13 Then
        If txtMobile.Text <> "" Then txtEmail.SetFocus
        If txtMobile.Text = "" Then txtMobile.Text = "Nil": txtEmail.SetFocus
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    Call CharOnly(KeyAscii) ' 4 Character Values Only.
    Call Alpha_Char(KeyAscii, frmMain, txtName) '4 First Capital Letter
    If KeyAscii = 13 Then
        If txtName.Text <> "" Then txtAddress.SetFocus: Call GetID(RstAB, tbl_AB): txtID.Text = RECID
    End If
End Sub

Private Sub txtNewUser_GotFocus()
    lblHelp1.Top = 2400: lblHelp1.Left = 5700: lblHelp1.Caption = "Enter New UserID for Login"
End Sub

Private Sub txtNewUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtNewUser.Text <> "" Then
            txtPWD.SetFocus
        ElseIf txtNewUser.Text = "" Then
            MsgBox "Empty field is not valid." & vbCrLf & "Please enter the New User for Login.", _
            vbCritical, "Error! Empty User ID": SendKeys "{Home}+{End}": txtNewUser.SetFocus
        End If
    End If
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
    Call NumOnly(KeyAscii) 'Only 4 Numeric Character.
    If KeyAscii = 13 Then
        If txtPhone.Text <> "" Then txtMobile.SetFocus
        If txtPhone.Text = "" Then txtPhone.Text = "Nil": txtMobile.SetFocus
    End If
End Sub

Private Sub txtPWD_GotFocus()
    lblHelp1.Top = 2890: lblHelp1.Left = 5700: lblHelp1.Caption = "Enter Password 4 Login, min 6-Char"
End Sub

Private Sub txtPWD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPWD.Text <> "" Then
            If Len(txtPWD.Text) < 6 Then 'use to verify the password length.
                MsgBox "Un-Athurised Password length." & vbCrLf & _
                       "Password not less than 6 characters", vbCritical, "Error! In Password...."
                       SendKeys "{Home}+{End}": txtPWD.SetFocus
            ElseIf Len(txtPWD.Text) >= 6 Then
                txtConPWD.SetFocus
            End If
        ElseIf txtPWD.Text = "" Then
            MsgBox "Password empty field is not supported." & vbCrLf & _
                   "Please enter the Valid Password.", vbExclamation, "Error! Empty Password Field"
                   SendKeys "{Home}+{End}": txtPWD.SetFocus
        End If
    End If
End Sub

Private Sub txtSearCell_KeyPress(KeyAscii As Integer)
    Call NumOnly(KeyAscii) 'use only 4 numeric values.
    If KeyAscii = 13 Then
        If txtSearName.Text <> "" Or txtSearLandLine.Text <> "" Or txtSearCell.Text <> "" Then
            Call ImgFind_Click
        ElseIf txtSearName.Text <> "" And txtSearLandLine.Text <> "" And txtSearCell.Text <> "" Then
            MsgBox "Invalid Search Values....." & vbCrLf & "Please atleast one values must for Search", _
                    vbInformation, "Error! Search Values...."
                    txtSearName.SetFocus
        End If
    End If
End Sub

Private Sub txtSearLandLine_KeyPress(KeyAscii As Integer)
    Call NumOnly(KeyAscii) 'use only 4 numeric values.
    If KeyAscii = 13 Then
        If txtSearLandLine.Text <> "" Then lblclick.Caption = "Click Start Search to Find": lblclick.Visible = True
        txtSearCell.SetFocus
    End If
End Sub

Private Sub txtSearName_KeyPress(KeyAscii As Integer)
    Call CharOnly(KeyAscii) 'use only 4 character values
    Call Alpha_Char(KeyAscii, frmMain, txtSearName) 'use only for first capital letter.
    If KeyAscii = 13 Then
        If txtSearName.Text <> "" Then lblclick.Caption = "Click Start Search to Find": lblclick.Visible = True
        txtSearLandLine.SetFocus
    End If
End Sub
