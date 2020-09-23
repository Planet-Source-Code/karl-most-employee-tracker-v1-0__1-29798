VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Employee Tracker 1.0"
   ClientHeight    =   7770
   ClientLeft      =   1770
   ClientTop       =   1485
   ClientWidth     =   11025
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   11025
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   630
      Top             =   7650
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontName        =   "Arial"
   End
   Begin VB.Data EMPData 
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\Carl Weis\My Documents\Employee Tracker\EmployeeDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1185
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Employee Records"
      Top             =   7770
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next>"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9540
      TabIndex        =   50
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< Back"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9540
      TabIndex        =   49
      Top             =   3525
      Width           =   1080
   End
   Begin VB.TextBox txtDate 
      DataField       =   "Date"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   8400
      TabIndex        =   47
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtNotes 
      DataField       =   "Notes"
      DataSource      =   "EMPData"
      Height          =   1935
      Left            =   5385
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   5685
      Width           =   5175
   End
   Begin VB.Frame Frame2 
      Height          =   1425
      Left            =   600
      TabIndex        =   42
      Top             =   5610
      Width           =   4695
      Begin VB.TextBox txtEmail 
         DataField       =   "E-mail"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label21 
         Caption         =   "E-Mail Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   585
      TabIndex        =   32
      Top             =   2760
      Width           =   10305
      Begin VB.TextBox txtMobile 
         DataField       =   "Mobile"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   6720
         TabIndex        =   20
         Top             =   2130
         Width           =   1815
      End
      Begin VB.TextBox txtFax 
         DataField       =   "Fax"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   6735
         TabIndex        =   19
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtExt 
         DataField       =   "Ext"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   6720
         TabIndex        =   18
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtWorkPhone 
         DataField       =   "Work Phone"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   6720
         TabIndex        =   17
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtEveningPhone 
         DataField       =   "Eveing Phone"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   6720
         TabIndex        =   16
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtDayTimePhone 
         DataField       =   "Daytime Phone"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   6720
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtZipCode 
         DataField       =   "Postal Code"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox CboState 
         DataField       =   "State"
         DataSource      =   "EMPData"
         Height          =   315
         ItemData        =   "frmMain.frx":0442
         Left            =   1080
         List            =   "frmMain.frx":0444
         TabIndex        =   13
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtCity 
         DataField       =   "City"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtLine2 
         DataField       =   "Line 2"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   705
         Width           =   4095
      End
      Begin VB.TextBox txtAddress 
         DataField       =   "Address "
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label20 
         Caption         =   "Mobile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   44
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   43
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   " Ext"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   41
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "Work Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   40
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Eveining Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5340
         TabIndex        =   39
         Top             =   735
         Width           =   1440
      End
      Begin VB.Label Label15 
         Caption         =   "Daytime Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   38
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "Zip Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   35
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Line 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   34
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox txtEmployeeID 
      DataField       =   "Employee #"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   8400
      TabIndex        =   9
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtSupervisorID 
      DataField       =   "Supervisor ID"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   4800
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtSupervisor 
      DataField       =   "Supervisor Name"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtSocialSecurity 
      DataField       =   "Social Security #"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   8400
      TabIndex        =   6
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtSalary 
      DataField       =   "Salary"
      DataSource      =   "DatPrimaryRS"
      Height          =   285
      Left            =   4800
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtDateHired 
      DataField       =   "Date Hired"
      DataSource      =   "DatPrimaryRS"
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtPosition 
      DataField       =   "Position"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   8400
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtLastName 
      DataField       =   "Last Name"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   4815
      TabIndex        =   2
      Top             =   1545
      Width           =   2175
   End
   Begin VB.TextBox txtFirstName 
      DataField       =   "First Name"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   1545
      Width           =   2175
   End
   Begin MSComctlLib.ImageList ImgLstToolbar1 
      Left            =   0
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0446
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2120
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":77AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9488
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B162
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CE3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EB16
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":107F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":124CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   1852
      ButtonWidth     =   1482
      ButtonHeight    =   1799
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgLstToolbar1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Logoff"
            Key             =   "Logoff"
            Object.ToolTipText     =   "Logoff"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            Key             =   "Add"
            Object.ToolTipText     =   "Add a new employee."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Key             =   "Edit"
            Object.ToolTipText     =   "Edit the current employee."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete the selected employee."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
            Object.ToolTipText     =   "Save any changes that were made."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "Print"
            Object.ToolTipText     =   "Print selected employees information."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Calculator"
            Key             =   "Calculator"
            Object.ToolTipText     =   "Open the calculator."
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "CD Player"
            Key             =   "CD Player"
            Object.ToolTipText     =   "Open the built in CD player."
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Date/Time"
            Key             =   "DateTime"
            Object.ToolTipText     =   "Insert the current Date and Time."
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "Help"
            Object.ToolTipText     =   "Open the help file."
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin VB.Label Label23 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7860
      TabIndex        =   48
      Top             =   1230
      Width           =   495
   End
   Begin VB.Label Label22 
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5430
      TabIndex        =   46
      Top             =   5445
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Employee ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   31
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Supervisor ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   30
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Supervisor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Social Security"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   28
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   27
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Date Hired"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7605
      TabIndex        =   25
      Top             =   1575
      Width           =   750
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3840
      TabIndex        =   24
      Top             =   1560
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   23
      Top             =   1560
      Width           =   945
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_cmd_Logoff 
         Caption         =   "&Logoff"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_cmd_Add 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnu_cmd_Edit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnu_cmd_Delete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnu_cmd_Save 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnu_cmd_Print 
         Caption         =   "&Print"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_cmd_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu_Accessories 
      Caption         =   "A&ccessories"
      Begin VB.Menu mnu_cmd_Calculator 
         Caption         =   "Calculat&or"
      End
      Begin VB.Menu mnu_cmd_CD_Player 
         Caption         =   "CD Pla&yer"
      End
   End
   Begin VB.Menu mnu_Format 
      Caption         =   "For&mat"
      Begin VB.Menu mnu_cmd_Font 
         Caption         =   "Fo&nt"
      End
      Begin VB.Menu mnu_cmd_DateTime 
         Caption         =   "Insert Date and Time"
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "&Help"
      Begin VB.Menu mnu_cmd_How_To 
         Caption         =   "&How To"
      End
      Begin VB.Menu mnu_cmd_About 
         Caption         =   "A&bout"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
'API Function Declaration
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Selected As Integer
Dim Focused As Boolean

Private Sub cmdBack_Click()
On Error Resume Next
  EMPData.Recordset.MovePrevious
  If EMPData.Recordset.BOF Then
      MsgBox "This is the first record in the database" _
    , vbInformation + vbOKOnly, "First Record"
      EMPData.Recordset.MoveFirst
  End If
End Sub

Private Sub cmdNext_Click()
On Error Resume Next
 EMPData.Recordset.MoveNext
  If EMPData.Recordset.EOF Then
      MsgBox "This is the last record in the database" _
    , vbInformation + vbOKOnly, "Last Record"
      EMPData.Recordset.MoveLast
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = vbDefault
End Sub

Private Sub mnu_cmd_About_Click()
frmAbout.Show
End Sub

Private Sub mnu_cmd_Add_Click()
'Code to add a new record
On Error GoTo AddErr
  EMPData.Recordset.AddNew
Exit Sub
AddErr:
    MsgBox Err.Description
End Sub

Private Sub mnu_cmd_Calculator_Click()
'Open the windows calculator
        X = Shell("C:\Windows.0\System32\calc.exe", 3)
End Sub

Private Sub mnu_cmd_CD_Player_Click()
frmCDPlayer.Show
End Sub

Private Sub mnu_cmd_DateTime_Click()
 currentDateTime = Now
     txtDate.Text = Now
End Sub

Private Sub mnu_cmd_Delete_Click()
'Code to delete the selected record
    On Error GoTo DeleteErr
With EMPData.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
 Exit Sub
DeleteErr:
    MsgBox Err.Description
End Sub

Private Sub mnu_cmd_Edit_Click()
'Code to edit the selected entry
    On Error GoTo EditErr
EMPData.Recordset.Edit
 Exit Sub
EditErr:
    MsgBox Err.Description
End Sub

Private Sub mnu_cmd_Exit_Click()
    Unload frmMain
   End
End Sub

Private Sub mnu_cmd_Font_Click()
CD1.ShowFont

txtFirstName.Text = CD1.FontBold
txtFirstName.Text = CD1.FontItalic
txtFirstName.Text = CD1.FontName
txtFirstName.Text = CD1.FontSize
txtFirstName.Text = CD1.FontStrikethru
txtFirstName.Text = CD1.FontUnderline

txtLastName.Text = CD1.FontBold
txtLastName.Text = CD1.FontItalic
txtLastName.Text = CD1.FontName
txtLastName.Text = CD1.FontSize
txtLastName.Text = CD1.FontStrikethru
txtLastName.Text = CD1.FontUnderline

txtPosition.Text = CD1.FontBold
txtPosition.Text = CD1.FontItalic
txtPosition.Text = CD1.FontName
txtPosition.Text = CD1.FontSize
txtPosition.Text = CD1.FontStrikethru
txtPosition.Text = CD1.FontUnderline

txtDateHired.Text = CD1.FontBold
txtDateHired.Text = CD1.FontItalic
txtDateHired.Text = CD1.FontName
txtDateHired.Text = CD1.FontSize
txtDateHired.Text = CD1.FontStrikethru
txtDateHired.Text = CD1.FontUnderline

txtSalary.Text = CD1.FontBold
txtSalary.Text = CD1.FontItalic
txtSalary.Text = CD1.FontName
txtSalary.Text = CD1.FontSize
txtSalary.Text = CD1.FontStrikethru
txtSalary.Text = CD1.FontUnderline

txtSocialSecurity.Text = CD1.FontBold
txtSocialSecurity.Text = CD1.FontItalic
txtSocialSecurity.Text = CD1.FontName
txtSocialSecurity.Text = CD1.FontSize
txtSocialSecurity.Text = CD1.FontStrikethru
txtSocialSecurity.Text = CD1.FontUnderline

txtSupervisor.Text = CD1.FontBold
txtSupervisor.Text = CD1.FontItalic
txtSupervisor.Text = CD1.FontName
txtSupervisor.Text = CD1.FontSize
txtSupervisor.Text = CD1.FontStrikethru
txtSupervisor.Text = CD1.FontUnderline

txtSupervisorID.Text = CD1.FontBold
txtSupervisorID.Text = CD1.FontItalic
txtSupervisorID.Text = CD1.FontName
txtSupervisorID.Text = CD1.FontSize
txtSupervisorID.Text = CD1.FontStrikethru
txtSupervisorID.Text = CD1.FontUnderline

txtEmployeeID.Text = CD1.FontBold
txtEmployeeID.Text = CD1.FontItalic
txtEmployeeID.Text = CD1.FontName
txtEmployeeID.Text = CD1.FontSize
txtEmployeeID.Text = CD1.FontStrikethru
txtEmployeeID.Text = CD1.FontUnderline

txtAddress.Text = CD1.FontBold
txtAddress.Text = CD1.FontItalic
txtAddress.Text = CD1.FontName
txtAddress.Text = CD1.FontSize
txtAddress.Text = CD1.FontStrikethru
txtAddress.Text = CD1.FontUnderline

txtLine2.Text = CD1.FontBold
txtLine2.Text = CD1.FontItalic
txtLine2.Text = CD1.FontName
txtLine2.Text = CD1.FontSize
txtLine2.Text = CD1.FontStrikethru
txtLine2.Text = CD1.FontUnderline

txtCity.Text = CD1.FontBold
txtCity.Text = CD1.FontItalic
txtCity.Text = CD1.FontName
txtCity.Text = CD1.FontSize
txtCity.Text = CD1.FontStrikethru
txtCity.Text = CD1.FontUnderline

txtZipCode.Text = CD1.FontBold
txtZipCode.Text = CD1.FontItalic
txtZipCode.Text = CD1.FontName
txtZipCode.Text = CD1.FontSize
txtZipCode.Text = CD1.FontStrikethru
txtZipCode.Text = CD1.FontUnderline

txtDayTimePhone.Text = CD1.FontBold
txtDayTimePhone.Text = CD1.FontItalic
txtDayTimePhone.Text = CD1.FontName
txtDayTimePhone.Text = CD1.FontSize
txtDayTimePhone.Text = CD1.FontStrikethru
txtDayTimePhone.Text = CD1.FontUnderline

txtEveningPhone.Text = CD1.FontBold
txtEveningPhone.Text = CD1.FontItalic
txtEveningPhone.Text = CD1.FontName
txtEveningPhone.Text = CD1.FontSize
txtEveningPhone.Text = CD1.FontStrikethru
txtEveningPhone.Text = CD1.FontUnderline

txtMobile.Text = CD1.FontBold
txtMobile.Text = CD1.FontItalic
txtMobile.Text = CD1.FontName
txtMobile.Text = CD1.FontSize
txtMobile.Text = CD1.FontStrikethru
txtMobile.Text = CD1.FontUnderline

txtFax.Text = CD1.FontBold
txtFax.Text = CD1.FontItalic
txtFax.Text = CD1.FontName
txtFax.Text = CD1.FontSize
txtFax.Text = CD1.FontStrikethru
txtFax.Text = CD1.FontUnderline

txtWorkPhone.Text = CD1.FontBold
txtWorkPhone.Text = CD1.FontItalic
txtWorkPhone.Text = CD1.FontName
txtWorkPhone.Text = CD1.FontSize
txtWorkPhone.Text = CD1.FontStrikethru
txtWorkPhone.Text = CD1.FontUnderline

txtExt.Text = CD1.FontBold
txtExt.Text = CD1.FontItalic
txtExt.Text = CD1.FontName
txtExt.Text = CD1.FontSize
txtExt.Text = CD1.FontStrikethru
txtExt.Text = CD1.FontUnderline

txtEmail.Text = CD1.FontBold
txtEmail.Text = CD1.FontItalic
txtEmail.Text = CD1.FontName
txtEmail.Text = CD1.FontSize
txtEmail.Text = CD1.FontStrikethru
txtEmail.Text = CD1.FontUnderline

txtNotes.Text = CD1.FontBold
txtNotes.Text = CD1.FontItalic
txtNotes.Text = CD1.FontName
txtNotes.Text = CD1.FontSize
txtNotes.Text = CD1.FontStrikethru
txtNotes.Text = CD1.FontUnderline
End Sub

Private Sub mnu_cmd_How_To_Click()
frmHelp.Show
End Sub

Private Sub mnu_cmd_Logoff_Click()
Unload frmMain
frmLogin.Show
End Sub

Private Sub mnu_cmd_Print_Click()
  Beep
    MsgBox ("This function has not yet been implemented."), vbOKOnly, "Comming Soon"
End Sub

Private Sub mnu_cmd_Save_Click()
'code to save the selected record.
    On Error GoTo SaveErr
 EMPData.Recordset.Update
 Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
        
    Case "Logoff"
    'Logoff and return to the login screen
        Unload frmMain
        frmLogin.Show
        
    Case "Add"
    'Code to add a new record
        On Error GoTo AddErr
  EMPData.Recordset.AddNew
AddErr:
    MsgBox Err.Source
    Case "Edit"
    'Code to edit a selected record.
        On Error GoTo EditErr
    EMPData.Recordset.Edit
EditErr:
    MsgBox Err.Description
    Case "Delete"
    'Code to delete the selected record.
        On Error GoTo DeleteErr
  With EMPData.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
DeleteErr:
    MsgBox Err.Description
    Case "Save"
    'Code to save the selected record.
        On Error GoTo SaveErr
  EMPData.Recordset.Update
SaveErr:
    MsgBox Err.Description
    Case "Print"
            Beep
    MsgBox ("This function has not yet been implemented."), vbOKOnly, "Comming Soon"
    
    Case "Calculator"
    'Open the windows calculator
        X = Shell("C:\Windows.0\System32\calc.exe", 3)
        
    Case "CD Player"
        'Show my cd player
        frmCDPlayer.Show
        
    Case "DateTime"
    'Code to insert the current date and time to the date field.
     currentDateTime = Now
     txtDate.Text = Now

    Case "Help"
    'Show the help file
        frmHelp.Show
    End Select
End Sub

