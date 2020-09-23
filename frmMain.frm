VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5895
   ClientLeft      =   15
   ClientTop       =   -45
   ClientWidth     =   7860
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   7440
      Top             =   5460
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.CheckBox chkIconize 
      BackColor       =   &H00000000&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Left            =   6900
      Style           =   1  'Graphical
      TabIndex        =   63
      TabStop         =   0   'False
      ToolTipText     =   "Iconize"
      Top             =   90
      Width           =   285
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   210
      Locked          =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      Text            =   "Ready..."
      Top             =   5610
      Width           =   7515
   End
   Begin VB.CheckBox chkMinimize 
      BackColor       =   &H00000000&
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Minimize"
      Top             =   90
      Width           =   285
   End
   Begin VB.CheckBox chkExit 
      BackColor       =   &H00000000&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   51
      TabStop         =   0   'False
      ToolTipText     =   "Close"
      Top             =   90
      Width           =   285
   End
   Begin VB.Frame fraReport 
      BackColor       =   &H00000000&
      Caption         =   "Report Builder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2520
      Left            =   420
      TabIndex        =   34
      Top             =   3030
      Width           =   7305
      Begin VB.CheckBox chkRecords 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Record Counts"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1545
      End
      Begin VB.CheckBox chkBrief 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Brief Format"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5280
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   510
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkEverything 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Everything Else"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   2130
         Width           =   1545
      End
      Begin VB.CheckBox chkRCode 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Response Code"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1590
         Width           =   1545
      End
      Begin VB.CheckBox chkTC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Truncated"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1050
         Value           =   1  'Checked
         Width           =   1545
      End
      Begin VB.CheckBox chkAA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Authoritative"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   780
         Value           =   1  'Checked
         Width           =   1545
      End
      Begin VB.CheckBox chkOpCode 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Operation Code"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   510
         Value           =   1  'Checked
         Width           =   1545
      End
      Begin VB.CheckBox chkRA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Recursion Avail."
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1545
      End
      Begin VB.CheckBox chkField 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Cache Time To Live"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   5280
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1590
         Width           =   1875
      End
      Begin VB.CheckBox chkField 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Domain Name"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   5280
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   780
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkField 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Record Type"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   5280
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1050
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkField 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Record Class"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   5280
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1875
      End
      Begin VB.CheckBox chkField 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Data Length"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   5280
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1875
      End
      Begin VB.CheckBox chkField 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Record Data"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   5280
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   2130
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkRecord 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Answer Records (AN)"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   2370
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1350
         Value           =   1  'Checked
         Width           =   2265
      End
      Begin VB.CheckBox chkRecord 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Authoritative Records (NS)"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   2370
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1620
         Value           =   1  'Checked
         Width           =   2265
      End
      Begin VB.CheckBox chkRecord 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Additional Records (AR)"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   2370
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1890
         Value           =   1  'Checked
         Width           =   2265
      End
      Begin VB.CheckBox chkQName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Query Records"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2370
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   540
         Width           =   2265
      End
      Begin VB.CheckBox chkQType 
         BackColor       =   &H00000000&
         Caption         =   "Query Type"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2610
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   810
         Width           =   1185
      End
      Begin VB.CheckBox chkQClass 
         BackColor       =   &H00000000&
         Caption         =   "Query Class"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2610
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Header Fields      "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   13
         Left            =   210
         TabIndex        =   49
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Record Fields           "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   11
         Left            =   5310
         TabIndex        =   48
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Records                         "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   12
         Left            =   2400
         TabIndex        =   47
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame fraHeader 
      BackColor       =   &H00000000&
      Caption         =   "Header Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2605
      Left            =   5670
      TabIndex        =   20
      Top             =   360
      Width           =   2055
      Begin VB.CheckBox chkHeaderRA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Recursion Available"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   810
         Width           =   1755
      End
      Begin VB.CheckBox chkHeaderAA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Authoritative"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   540
         Width           =   1755
      End
      Begin VB.CheckBox chkHeaderTC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Truncated"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   270
         Width           =   1755
      End
      Begin VB.Label lblAdditional 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   33
         Top             =   2190
         Width           =   555
      End
      Begin VB.Label lblAuthority 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   32
         Top             =   1950
         Width           =   555
      End
      Begin VB.Label lblAnswer 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         Top             =   1710
         Width           =   555
      End
      Begin VB.Label lblQuestion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   30
         Top             =   1470
         Width           =   555
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Record Counts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   150
         TabIndex        =   29
         Top             =   1170
         Width           =   1335
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Additional"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   270
         TabIndex        =   28
         Top             =   2190
         Width           =   765
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Authority"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   270
         TabIndex        =   27
         Top             =   1950
         Width           =   765
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Answer"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   270
         TabIndex        =   26
         Top             =   1710
         Width           =   765
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Question"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   270
         TabIndex        =   25
         Top             =   1470
         Width           =   765
      End
   End
   Begin VB.Frame fraQuery 
      BackColor       =   &H00000000&
      Caption         =   "Query Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1590
      Left            =   420
      TabIndex        =   16
      Top             =   1380
      Width           =   5110
      Begin VB.CheckBox chkRecursion 
         BackColor       =   &H00000000&
         Caption         =   "Local Recursion"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1920
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   540
         Value           =   1  'Checked
         Width           =   1545
      End
      Begin VB.CommandButton cmdReverse 
         Caption         =   "&Lookup IP"
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   540
         Width           =   1275
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load Request"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   1275
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Request"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   1275
      End
      Begin VB.CommandButton cmdReport 
         Caption         =   "&Build Report"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1290
         TabIndex        =   4
         Top             =   240
         Width           =   1065
      End
      Begin VB.CheckBox chkRD 
         BackColor       =   &H00000000&
         Caption         =   "Recursion Desired"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   540
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Send &Query"
         Enabled         =   0   'False
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   1065
      End
      Begin VB.ComboBox cmbQClass 
         Height          =   315
         ItemData        =   "frmMain.frx":0442
         Left            =   3720
         List            =   "frmMain.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1140
         Width           =   1275
      End
      Begin VB.ComboBox cmbQType 
         Height          =   315
         ItemData        =   "frmMain.frx":0446
         Left            =   2400
         List            =   "frmMain.frx":0448
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1140
         Width           =   1275
      End
      Begin VB.TextBox txtQName 
         Height          =   315
         Left            =   150
         MaxLength       =   254
         TabIndex        =   0
         Top             =   1140
         Width           =   2175
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   3720
         TabIndex        =   19
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   18
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   150
         TabIndex        =   17
         Top             =   900
         Width           =   2205
      End
   End
   Begin VB.Frame fraDNSServer 
      BackColor       =   &H00000000&
      Caption         =   "DNS Server Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   420
      TabIndex        =   13
      Top             =   360
      Width           =   5115
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   285
         Left            =   4350
         TabIndex        =   12
         ToolTipText     =   "Edit Selected Entry"
         Top             =   540
         Width           =   645
      End
      Begin VB.CommandButton cmdAddDel 
         Caption         =   "&Del"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   11
         ToolTipText     =   "Add/Remove Item"
         Top             =   540
         Width           =   645
      End
      Begin VB.ComboBox cmbServer 
         Height          =   315
         ItemData        =   "frmMain.frx":044A
         Left            =   150
         List            =   "frmMain.frx":044C
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   510
         Width           =   3525
      End
      Begin VB.OptionButton optProtocol 
         BackColor       =   &H00000000&
         Caption         =   "UDP"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   3690
         TabIndex        =   8
         Top             =   210
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton optProtocol 
         BackColor       =   &H00000000&
         Caption         =   "TCP"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   4380
         TabIndex        =   9
         Top             =   210
         Width           =   615
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Server"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   5
         Left            =   150
         TabIndex        =   15
         Top             =   270
         Width           =   525
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Protocol"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   14
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   60
      MousePointer    =   15  'Size All
      TabIndex        =   53
      Top             =   60
      Width           =   7725
   End
   Begin VB.Line Line 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   3
      X1              =   300
      X2              =   300
      Y1              =   300
      Y2              =   5550
   End
   Begin VB.Line Line 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   5760
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   " DNS Lookup v2.03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   360
      TabIndex        =   50
      Top             =   90
      Width           =   2100
   End
   Begin VB.Line Line 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   285
      X2              =   7200
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Shape Border 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   5880
      Index           =   1
      Left            =   30
      Top             =   30
      Width           =   7845
   End
   Begin VB.Shape Border 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      Height          =   5880
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   7845
   End
   Begin VB.Line Line 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   105
      X2              =   7200
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Changes 2.04
    'Fixed time interval fields to show days.  The hours was previously truncated

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const WM_SETTEXT = &HC

Dim WithEvents TrayIcon As SysTray
Attribute TrayIcon.VB_VarHelpID = -1
Dim SkipChange As Boolean, ServerName As String
Dim RawData() As Byte, DNSData As DNSPacket
Private Sub Form_Load()
On Error Resume Next
    Set TrayIcon = New SysTray
    SendMessage Me.hwnd, WM_SETTEXT, 0, ByVal "DNS Lookup"

    Open App.Path & "\servers.txt" For Input As #1
    If Err.Number = 0 Then
        While Not EOF(1)
            Line Input #1, Data
            If Trim(Data) <> "" Then cmbServer.AddItem Trim(Data), cmbServer.ListCount
        Wend
        Close #1
        If cmbServer.ListCount > 0 Then cmbServer.ListIndex = 0
    Else
        Err.Clear
    End If

    On Error GoTo ERR_INI
    Keys = GetKeys("TypeConstants")
    For X = 0 To UBound(Keys)
        cmbQType.AddItem Keys(X)
    Next
    cmbQType.ListIndex = 0

    Keys = GetKeys("ClassConstants")
    For X = 0 To UBound(Keys)
        cmbQClass.AddItem Keys(X)
    Next
    cmbQClass.ListIndex = 0
Exit Sub
ERR_INI:
    MsgBox "strings.ini is not found or is corrupt.", vbCritical, "Fatal Error!"
    End
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Y = 0 Then TrayIcon.HandleEvent X
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If cmbServer.ListIndex > -1 Then
        TopItem = cmbServer.List(cmbServer.ListIndex)
        cmbServer.RemoveItem cmbServer.ListIndex
        cmbServer.AddItem TopItem, 0
    End If
    Open App.Path & "\servers.txt" For Output As #1
        For X = 0 To cmbServer.ListCount - 1
            Print #1, cmbServer.List(X)
        Next
    Close #1
End Sub
Private Sub lblHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub
Private Sub chkIconize_Click()
    TrayIcon.Create Me.hwnd, Me.Icon.Handle, "DNS Lookup"
    Me.Visible = False
    chkIconize.Value = 0
End Sub
Private Sub chkMinimize_Click()
    Me.WindowState = vbMinimized
    chkMinimize.Value = 0
End Sub
Private Sub chkExit_Click()
    Winsock.Close
    DoEvents
    Unload Me
End Sub

'Tray Icon events

Private Sub TrayIcon_LButtonDown()
    frmMain.Visible = True
    TrayIcon.Remove
End Sub
Private Sub TrayIcon_RButtonDown()
    TrayIcon_LButtonDown
End Sub

'Custom-Disable the Checkboxes in the Header Information frame

Private Sub chkHeaderTC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    chkHeaderTC.Tag = chkHeaderTC.Value
End Sub
Private Sub chkHeaderTC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    chkHeaderTC.Value = chkHeaderTC.Tag
End Sub
Private Sub chkHeaderAA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    chkHeaderAA.Tag = chkHeaderAA.Value
End Sub
Private Sub chkHeaderAA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    chkHeaderAA.Value = chkHeaderAA.Tag
End Sub
Private Sub chkHeaderRA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    chkHeaderRA.Tag = chkHeaderRA.Value
End Sub
Private Sub chkHeaderRA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    chkHeaderRA.Value = chkHeaderRA.Tag
End Sub

'Control Events in the "DNS Server Setup" frame

Private Sub cmbServer_Click()
    cmdAddDel.Enabled = True
    cmdAddDel.Caption = "&Del"
    If txtQName.Text <> "" Then cmdSearch.Enabled = True
End Sub
Private Sub cmbServer_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 8
        SkipChange = True
    Case 13
        If cmdAddDel.Enabled Then cmdAddDel_Click
        KeyAscii = 0
    End Select
End Sub
Private Sub cmbServer_Change()
    Enable = cmbServer.Text <> ""
    cmdAddDel.Caption = "&Add"
    cmdAddDel.Enabled = Enable
    cmdReverse.Enabled = IIf(txtQName.Text = "", False, Enable)
    cmdSearch.Enabled = IIf(txtQName.Text = "", False, Enable)
'Auto-Complete Code
    If SkipChange Then
        SkipChange = False
    ElseIf Enable Then
        Length = Len(cmbServer.Text)
        For X = 0 To cmbServer.ListCount - 1
            If Left(cmbServer.List(X), Length) = cmbServer.Text Then
                cmdAddDel.Caption = "&Del"
                If cmbServer.List(X) = cmbServer.Text Then Exit Sub

                cmbServer.Text = cmbServer.List(X)
                cmbServer.SelStart = Length
                cmbServer.SelLength = Len(cmbServer.Text) - Length
                Exit Sub
            End If
        Next
    End If
End Sub
Private Sub cmdAddDel_Click()
Dim X As Long
    If cmdAddDel.Caption = "&Del" Then
        For X = 0 To cmbServer.ListCount - 1
            If cmbServer.List(X) = cmbServer.Text Then
                cmbServer.RemoveItem X
                If cmbServer.ListCount > 0 Then
                    cmbServer.ListIndex = 0
                Else
                    cmbServer.Text = ""
                    cmdAddDel.Enabled = False
                    cmbServer.SetFocus
                End If
                Exit Sub
            End If
        Next
    Else
        cmdAddDel.Caption = "&Del"
        cmbServer.AddItem cmbServer.Text
    End If
End Sub
Private Sub cmdEdit_Click()
    For X = 0 To cmbServer.ListCount - 1
        If cmbServer.List(X) = cmbServer.Text Then
            Server = InputBox("Enter the value to replace the currently displayed value", "Edit DNS Server Entry", cmbServer.Text)
            If Server <> "" Then cmbServer.List(cmbServer.ListIndex) = Server
            Exit Sub
        End If
    Next
    MsgBox "Please select a DNS Server entry from the list", vbInformation, "No Entry Selected"
End Sub

'Control Events in the "Query Setup" frame

Private Sub txtQName_GotFocus()
    txtQName.SelStart = 0
    txtQName.SelLength = Len(txtQName.Text)
End Sub
Private Sub txtQName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSearch.SetFocus
End Sub
Private Sub txtQName_Change()
    cmdSearch.Enabled = txtQName.Text <> "" And cmbServer.Text <> ""
End Sub
Private Sub cmdSearch_Click()
    If cmdSearch.Caption = "Send &Query" Then
        txtQName.SetFocus
        'Start the DNS Lookup
        LockForm True
        'reset the local recursion to 'first lookup'
        Winsock.Tag = ""
        Server = Split(cmbServer.Text, " ")
        If InStr(Server(0), ":") > 0 Then
            InitiateQuery Left(Server(0), InStrRev(Server(0), ":") - 1), cmbServer.Text, Val(Mid(Server(0), InStrRev(Server(0), ":") + 1))
        Else
            InitiateQuery Server(0), cmbServer.Text, 53
        End If
    Else
        'Cancel the Search
        Winsock.Tag = ""
        LockForm False
        txtQName.SetFocus
        txtStatus.Text = "Cancelled Search"
    End If
End Sub
Private Sub cmdReport_Click()
Dim Record As DNSRecord
    If txtQName.Enabled = True Then txtQName.SetFocus

    If chkEverything.Value Then
        Header = Header & "ID:" & vbTab & vbTab & vbTab & DNSData.Header.ID & vbCrLf
        Header = Header & "Query/Response:" & vbTab & vbTab & IIf(DNSData.Header.QR, "Response", "Query") & vbCrLf
    End If
    If chkOpCode.Value Then Header = Header & "Operation Code:" & vbTab & vbTab & GetStr("OperationCodes", CStr(DNSData.Header.OpCode)) & vbCrLf
    If chkAA.Value Then Header = Header & "Authoritative:" & vbTab & vbTab & IIf(DNSData.Header.AA, "Yes", "No") & vbCrLf
    If chkTC.Value Then Header = Header & "Truncated:" & vbTab & vbTab & IIf(DNSData.Header.TC, "Yes", "No") & vbCrLf
    If chkEverything.Value Then Header = Header & "Recursion Desired:" & vbTab & IIf(DNSData.Header.RD, "Yes", "No") & vbCrLf
    If chkRA.Value Then Header = Header & "Recursion Available:" & vbTab & IIf(DNSData.Header.RA, "Yes", "No") & vbCrLf
    If chkRCode.Value Then Header = Header & "Response Code:" & vbTab & vbTab & GetStr("ResponseCodes", CStr(DNSData.Header.RCode)) & vbCrLf
    If chkEverything.Value Then Header = Header & "Z (Reserved):" & vbTab & vbTab & DNSData.Header.Z & vbCrLf
    If chkRecords.Value Then
        Header = Header & vbCrLf & "Record Counts" & vbCrLf
        Header = Header & "  Question:" & vbTab & DNSData.Header.QDCount & vbCrLf
        Header = Header & "  Answer:" & vbTab & DNSData.Header.ANCount & vbCrLf
        Header = Header & "  Authority:" & vbTab & DNSData.Header.NSCount & vbCrLf
        Header = Header & "  Additional:" & vbTab & DNSData.Header.ARCount & vbCrLf
    End If
    If Header <> "" Then Header = "Response By: " & ServerName & vbCrLf & vbCrLf & "--- Header Information ---" & vbCrLf & vbCrLf & Header & vbCrLf

    If chkQName.Value Then
        Questions = "--- Query Record(s) ---" & vbCrLf & vbCrLf
        For X = 1 To DNSData.Header.QDCount
            Questions = Questions & "Query: " & DNSData.Question(X).QName & vbCrLf
            If chkQType.Value Then Questions = Questions & "Type : " & GetStr("TypeStrings", CStr(DNSData.Question(X).QType)) & vbCrLf
            If chkQClass.Value Then Questions = Questions & "Class: " & GetStr("ClassStrings", CStr(DNSData.Question(X).QClass)) & vbCrLf
            Questions = Questions & vbCrLf
        Next
        Questions = Questions & vbCrLf
    End If

    If chkBrief.Value = 0 Then
        For X = 1 To 3
            Select Case X
            Case 1
                RCount = DNSData.Header.ANCount
                RecName = "Answer"
            Case 2
                RCount = DNSData.Header.NSCount
                RecName = "Authority"
            Case 3
                RCount = DNSData.Header.ARCount
                RecName = "Additional"
            End Select
            If chkRecord(X - 1).Value Then
                Records = Records & "--- " & RecName & " Record(s) ---" & vbCrLf & vbCrLf
                For Y = 1 To RCount
                    Select Case X
                    Case 1: Record = DNSData.Answer(Y)
                    Case 2: Record = DNSData.Authority(Y)
                    Case 3: Record = DNSData.Additional(Y)
                    End Select
                    If chkField(0).Value Then Records = Records & "  Name  : " & Record.RName & vbCrLf
                    If chkField(1).Value Then Records = Records & "  Type  : " & GetStr("TypeStrings", CStr(Record.RType)) & vbCrLf
                    If chkField(2).Value Then Records = Records & "  Class : " & GetStr("ClassStrings", CStr(Record.RClass)) & vbCrLf
                    If chkField(3).Value Then Records = Records & "  TTL   : " & FormatSeconds(Record.TTL) & vbCrLf
                    If chkField(4).Value Then Records = Records & "  Length: " & Record.RDLength & vbCrLf
                    If chkField(5).Value Then
                        Data = GetData(Record)
                        If InStr(Data, vbCrLf) Then
                            Records = Records & Data
                        Else
                            Records = Records & "  Data  : " & Data & vbCrLf
                        End If
                    End If
                    Records = Records & vbCrLf
                Next
                Records = Records & vbCrLf
            End If
        Next
    Else
        If chkRecord(0).Value Then
            Records = Records & "--- Answer Record(s) ---" & vbCrLf & vbCrLf
            For X = 1 To DNSData.Header.ANCount
                Select Case DNSData.Answer(X).RType
                Case 1
                    A4 = A4 & "  " & DNSData.Answer(X).RData & vbCrLf
                Case 2
                    NS = NS & "  " & DNSData.Answer(X).RData & vbCrLf
                Case 15
                    MX = MX & "  " & GetData(DNSData.Answer(X)) & vbCrLf
                Case 16
                    TXT = TXT & "  " & DNSData.Answer(X).RData & vbCrLf
                Case 28
                    A6 = A6 & "  " & DNSData.Answer(X).RData & vbCrLf
                Case Else
                    RR = RR & ExpRecord(DNSData.Answer(X)) & vbCrLf
                End Select
            Next
            If A4 <> "" Then Records = Records & "IP4 Address List" & vbCrLf & A4 & vbCrLf
            If NS <> "" Then Records = Records & "DNS Server List" & vbCrLf & NS & vbCrLf
            If MX <> "" Then Records = Records & "Mail Exchange Server List" & vbCrLf & MX & vbCrLf
            If TXT <> "" Then Records = Records & "Text Records" & vbCrLf & TXT & vbCrLf
            If A6 <> "" Then Records = Records & "IP6 Address List" & vbCrLf & A6 & vbCrLf
            If RR <> "" Then Records = Records & RR & vbCrLf
        End If
        If chkRecord(1).Value Then
            NS = ""
            RR = ""
            Records = Records & "--- Authority Record(s) ---" & vbCrLf & vbCrLf
            For X = 1 To DNSData.Header.NSCount
                Select Case DNSData.Authority(X).RType
                Case 2
                    NS = NS & "  " & DNSData.Authority(X).RData & vbCrLf
                Case Else
                    RR = RR & ExpRecord(DNSData.Authority(X)) & vbCrLf
                End Select
            Next
            If NS <> "" Then Records = Records & "DNS Server List" & vbCrLf & NS & vbCrLf
            Records = Records & RR
        End If
        If chkRecord(2).Value Then
            A4 = ""
            A6 = ""
            RR = ""
            Records = Records & "--- Additional Record(s) ---" & vbCrLf & vbCrLf
            For X = 1 To DNSData.Header.ARCount
                Select Case DNSData.Additional(X).RType
                Case 1
                    Dim Address As String * 16
                    Address = DNSData.Additional(X).RData & Space(9)
                    A4 = A4 & "  " & Address & DNSData.Additional(X).RName & vbCrLf
                Case 28
                    IP6Address = DNSData.Additional(X).RData
                    A6 = A6 & "  " & IP6Address & " " & DNSData.Additional(X).RName & vbCrLf
                Case Else
                    RR = RR & ExpRecord(DNSData.Additional(X)) & vbCrLf
                End Select
            Next
            If A4 <> "" Then Records = Records & "IP4 Address List" & vbCrLf & A4 & vbCrLf
            If A6 <> "" Then Records = Records & "IP6 Address List" & vbCrLf & A6 & vbCrLf
            Records = Records & RR
        End If
    End If

    Open App.Path & "\" & DNSData.Question(1).QName & ".nsr" For Output As #1
        Print #1, Header & Questions & Records
    Close #1
    Shell "Notepad """ & App.Path & "\" & DNSData.Question(1).QName & ".nsr", vbNormalFocus
End Sub
Private Sub cmdReverse_Click()
    IPString = InputBox("Enter the IP address to resolve", "Auto Reverse Lookup")
    If IPString <> "" Then
        IPArray = Split(IPString, ".")
        For X = UBound(IPArray) To 0 Step -1
            QName = QName & IPArray(X) & "."
        Next
        txtQName.Text = QName & "IN-ADDR.ARPA"
        cmbQType.ListIndex = 11
        cmdSearch_Click
    End If
End Sub

'Control Events in the Report Builder Frame

Private Sub chkBrief_Click()
    For X = 0 To 5
        chkField(X).Enabled = chkBrief.Value = 0
    Next
End Sub

'Winsock Events

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    txtStatus.Text = "ERROR: " & Description
    LockForm False
End Sub
Private Sub Winsock_Connect()
    SendQuery
End Sub
Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    If Winsock.Protocol = sckTCPProtocol Then
        Winsock.PeekData Data, vbArray + vbByte, 2
        If bytesTotal <> Data(0) * 256 + Data(1) + 2 Then
            Exit Sub
        Else
            Winsock.GetData Data, vbArray + vbByte, 2
        End If
    End If
    On Error Resume Next
    Winsock.GetData RawData, vbArray + vbByte
    If Err.Number Then
        Winsock_Error Err.Number, Err.Description, 0, "", "", 0, False
        Exit Sub
    End If
    On Error GoTo 0

    DNSData = GetDNSInfo(RawData)
    If DNSData.Header.AA = False And Val(Winsock.Tag) < 5 And chkRecursion.Value = 1 Then
        For X = 1 To DNSData.Header.NSCount
            If DNSData.Authority(X).RType = 2 Then
                Winsock.Tag = Val(Winsock.Tag) + 1
                Winsock.Close
                For Y = 1 To DNSData.Header.ARCount
                    If DNSData.Authority(X).RData = DNSData.Additional(Y).RName Then
                        InitiateQuery DNSData.Authority(X).RData, DNSData.Additional(Y).RData, 53
                        Exit Sub
                    End If
                Next
                InitiateQuery DNSData.Authority(X).RData, DNSData.Authority(X).RData, 53
                Exit Sub
            End If
        Next
    End If
    Winsock.Tag = ""
    LockForm False
    cmdReport.Enabled = True

    If Winsock.Protocol = sckTCPProtocol Then
        chkHeaderTC.FontBold = True
        chkHeaderTC.Enabled = False
        chkHeaderTC.Value = 0
    Else
        chkHeaderTC.FontBold = False
        chkHeaderTC.Enabled = True
        chkHeaderTC.Value = Abs(DNSData.Header.TC)
    End If
    chkHeaderAA.Value = Abs(DNSData.Header.AA)
    chkHeaderRA.Value = Abs(DNSData.Header.RA)
    lblQuestion.Caption = DNSData.Header.QDCount
    lblAnswer.Caption = DNSData.Header.ANCount
    lblAuthority.Caption = DNSData.Header.NSCount
    lblAdditional.Caption = DNSData.Header.ARCount

    Select Case DNSData.Header.RCode
    Case 0
        If DNSData.Question(1).QType = 255 Then
            txtStatus.Text = "See Report"
            Exit Sub
        End If
        For X = 1 To DNSData.Header.ANCount
            If DNSData.Answer(X).RType = DNSData.Question(1).QType Then
                Select Case DNSData.Question(1).QType
                Case 1: txtStatus.Text = "Address: " & DNSData.Answer(X).RData
                Case 2 To 5, 7 To 10, 12, 15, 16, 19, 21
                    txtStatus.Text = ExpRecord(DNSData.Answer(X))
                Case 6
                    Dim SOAData As SOA
                    P = DNSData.Answer(X).RData
                    SOAData = GetSOA(RawData, P)
                    txtStatus.Text = "Owner MailBox: " & SOAData.RName
                Case 11: txtStatus.Text = "WKS Record Found! - Sweet! notify SilentRage!"
                Case 13
                    Dim HINFOData As HINFO
                    P = DNSData.Answer(X).RData
                    HINFOData = GetHINFO(Data, P)
                    txtStatus.Text = "OS: " & HINFOData.OS & " | CPU: " & HINFOData.CPU
                Case 17
                    Dim RPData As RP
                    P = DNSData.Answer(X).RData
                    RPData = GetRP(RawData, P)
                    txtStatus.Text = "Owner MailBox: " & RPData.MBox_DName & " | TXT Domain: " & RPData.TXT_DName
                Case 18
                    Dim AFSDBData As AFSDB
                    P = DNSData.Answer(X).RData
                    AFSDBData = GetAFSDB(RawData, P)
                    txtStatus.Text = "Database HostName: " & AFSDBData.HostName
                Case 20
                    Dim ISDNData As ISDN
                    P = DNSData.Answer(X).RData
                    ISDNData = GetISDN(RawData, P)
                    txtStatus.Text = "ISDN-Address: " & ISDNData.Address
                Case 28
                    txtStatus.Text = "IP6 Address: " & DNSData.Answer(X).RData
                Case 29
                    Dim LOCData As LOC
                    P = DNSData.Answer(X).RData
                    LOCData = GetLOC(RawData, P)

                    Lat = Abs(LOCData.Latitude - 2147483648#)
                    D = Int(Lat / 3600000)
                    M = Int((Lat Mod 3600000) / 60000)
                    If (Lat Mod 60000) / 1000 >= 30 Then M = M + 1
                    txtStatus.Text = "lat/lon = " & D & "." & M & IIf(LOCData.Latitude - 2147483648# > 0, "n", "s") & ", "

                    Lat = Abs(LOCData.Longitude - 2147483648#)
                    D = Int(Lat / 3600000)
                    M = Int((Lat Mod 3600000) / 60000)
                    If (Lat Mod 60000) / 1000 >= 30 Then M = M + 1
                    txtStatus.Text = txtStatus.Text & D & "." & M & IIf(LOCData.Longitude - 2147483648# > 0, "e", "w")
                Case Else
                    txtStatus.Text = "Requested Record Available"
                End Select

                Exit Sub
            End If
        Next
        txtStatus.Text = "Record Unavailable"
    Case Else
        txtStatus.Text = "Response Code: " & GetStr("ResponseCodes", CStr(DNSData.Header.RCode))
    End Select
End Sub
Private Sub Winsock_Close()
    Winsock_Error 0, "Disconnected by Peer", 0, "", "", 0, False
End Sub

'Other Procedures

Private Function FormatSeconds(ByVal Seconds As Long) As String
Dim D As String, H As String, M As String, S As String
    D = Right("0" & Seconds \ 86400, 2) & "'"
    Seconds = Seconds Mod 86400
    H = Right("0" & Seconds \ 3600, 2) & "'"
    Seconds = Seconds Mod 3600
    M = Right("0" & Seconds \ 60, 2) & "'"
    Seconds = Seconds Mod 60
    S = Right("0" & Seconds, 2)

    FormatSeconds = D & H & M & S
End Function
Private Function LockForm(DoLock As Boolean)
    Winsock.Close
    cmbServer.Enabled = Not DoLock
    cmdSearch.Caption = IIf(DoLock, "&Cancel", "Send &Query")
    cmdReverse.Enabled = Not DoLock
    txtQName.Enabled = Not DoLock
    cmbQType.Enabled = Not DoLock
    cmbQClass.Enabled = Not DoLock
    chkRD.Enabled = Not DoLock
    chkRecursion.Enabled = Not DoLock
End Function
Private Sub InitiateQuery(ByVal HostName As String, ByVal HostIP As String, ByVal Port As Long)
    ServerName = HostName
    Winsock.Tag = Val(Winsock.Tag) + 1
    Winsock.RemoteHost = HostIP
    Winsock.RemotePort = Port
    txtStatus.Text = "Looking up " & txtQName.Text & "@" & HostIP & ":" & Port & " via " & IIf(optProtocol(0).Value, "UDP", "TCP")
    If optProtocol(0).Value Then
        Winsock.Protocol = sckUDPProtocol
        Winsock.Bind
        SendQuery
    Else
        Winsock.Protocol = sckTCPProtocol
        Winsock.Connect
    End If
End Sub
Private Function SendQuery()
  'Set Recursion Desired to the user selection
    RD = chkRD.Value
  'Format the Query Name
    Labels = Split(txtQName.Text, ".")
    For X = 0 To UBound(Labels)
        Length = Len(Labels(X))
        If Length > 63 Then
            MsgBox "Label '" & Labels(X) & "' is too long (max 63 characters)", vbExclamation, "Syntax Error!"
            txtQName.SetFocus
        Else
            Domain = Domain & Chr(Length) & Labels(X)
        End If
    Next
    QName = Domain & Chr(0)
  'Set Query Type to the user selection
    QType = Abs(GetInt("TypeConstants", cmbQType.Text))
    If QType > 65535 Then QType = 65535
    QType = Chr(Int(QType / 256)) & Chr(QType Mod 256)
  'Set the Query Class to the user selection
    QClass = Abs(GetInt("ClassConstants", cmbQClass.Text))
    If QClass > 65535 Then QClass = 65535
    QClass = Chr(Int(QClass / 256)) & Chr(QClass Mod 256)

    Query = Chr(0) & Chr(1) & Chr(RD) & Chr(0) & Chr(0) & Chr(1) & String(6, Chr(0)) & QName & QType & QClass
    If Winsock.Protocol = TCPProtocol Then Query = Chr(Int(Len(Query) / 256)) & Chr(Len(Query) Mod 256) & Query
On Error Resume Next
    Winsock.SendData Query
    If Err.Number Then Winsock_Error 0, "Invalid DNS server", 0, "", "", 0, False
End Function
Private Function ExpRecord(Record As DNSRecord) As String
    Select Case Record.RType
    Case 1
        Dim Address As String * 16
        Address = Record.RData & Space(9)
        ExpRecord = "Address: " & Address & Record.RName
    Case 2:  ExpRecord = "DNS Server: " & Record.RData
    Case 3:  ExpRecord = "Mail Dest: " * Record.RData
    Case 4:  ExpRecord = "Mail Forward: " & Record.RData
    Case 5:  ExpRecord = "Canonical Name: " & Record.RData
    Case 6
        ExpRecord = ExpRecord & "Zone of Authority" & vbCrLf
        ExpRecord = ExpRecord & GetData(Record)
    Case 7:  ExpRecord = "MailBox: " & Record.RData
    Case 8:  ExpRecord = "Mail Group: " & Record.RData
    Case 9:  ExpRecord = "Mail Rename: " & Record.RData
    Case 10: ExpRecord = "Null Record: " & GetData(Record)
    Case 11
        ExpRecord = ExpRecord & "Well Known Services" & vbCrLf
        ExpRecord = ExpRecord & GetData(Record)
    Case 12: ExpRecord = "Pointer: " & Record.RData
    Case 13
        ExpRecord = ExpRecord & "Host Information" & vbCrLf
        ExpRecord = ExpRecord & GetData(Record)
    Case 14
        ExpRecord = ExpRecord & "Mail Information" & vbCrLf
        ExpRecord = ExpRecord & GetData(Record)
    Case 15: ExpRecord = "Mail Exchange: " & GetData(Record)
    Case 16: ExpRecord = "Text: " & Record.RData
    Case 17
        ExpRecord = ExpRecord & "Responsible Person" & vbCrLf
        ExpRecord = ExpRecord & GetData(Record)
    Case 18
        ExpRecord = ExpRecord & "Andrew File System Database" & vbCrLf
        ExpRecord = ExpRecord & GetData(Record)
    Case 19: ExpRecord = "PSDN-Address: " & Record.RData
    Case 20
        ExpRecord = ExpRecord & "Integrated Service Digital Network" & vbCrLf
        ExpRecord = ExpRecord & GetData(Record)
    Case 21: ExpRecord = "Intermediate Host: " & GetData(Record)
    Case 28: ExpRecord = "IP6 Address: " & Record.RData & " " & Record.RName
    Case 29
        ExpRecord = ExpRecord & "Location Information" & vbCrLf
        ExpRecord = ExpRecord & GetData(Record)
    Case Else: ExpRecord = "Unsupported Record Type '" & GetStr("TypeStrings", CStr(Record.RType)) & "'"
    End Select
End Function
Private Function GetData(Record As DNSRecord)
    Select Case Record.RType
    Case 1, 2 To 5, 7 To 9, 12, 16, 19, 28
        GetData = Record.RData
    Case 6
        Dim SOAData As SOA
        P = Record.RData
        SOAData = GetSOA(RawData, P)
        GetData = GetData & "  Source : " & SOAData.MName & vbCrLf
        GetData = GetData & "  MailBox: " & SOAData.RName & vbCrLf
        GetData = GetData & "  Serial : " & SOAData.Serial & vbCrLf
        GetData = GetData & "  Refresh: " & FormatSeconds(SOAData.Refresh) & vbCrLf
        GetData = GetData & "  Retry  : " & FormatSeconds(SOAData.Retry) & vbCrLf
        GetData = GetData & "  Expire : " & FormatSeconds(SOAData.Expire) & vbCrLf
        GetData = GetData & "  Minimum: " & FormatSeconds(SOAData.Minimum) & vbCrLf
    Case 10 'NULL Record
        GetData = "Notify SilentRage of this lookup"
    Case 11
        Dim WKSData As WKS
        P = Record.RData
        WKSData = GetWKS(RawData, Record.RDLength, P)
        For X = 0 To UBound(WKSData.PortMap)
            If WKSData.PortMap(X) Then PortList = PortList & GetStr("Services", CStr(X)) & ","
        Next
        GetData = GetData & "  Address : " & WKSData.Address & vbCrLf
        GetData = GetData & "  Protocol: " & GetStr("Protocols", CStr(WKSData.Protocol)) & vbCrLf
        GetData = GetData & "  Services: " & Left(PortList, Len(PortList) - 1) & vbCrLf
    Case 13
        Dim HINFOData As HINFO
        P = Record.RData
        HINFOData = GetHINFO(RawData, P)
        GetData = GetData & "  OS : " & HINFOData.OS & vbCrLf
        GetData = GetData & "  CPU: " & HINFOData.CPU & vbCrLf
    Case 14
        Dim MINFOData As MINFO
        P = Record.RData
        MINFOData = GetMINFO(RawData, P)
        GetData = GetData & "  Responsible MailBox: " & MINFOData.RMailBX & vbCrLf
        GetData = GetData & "  Error MailBox      : " & MINFOData.EMailBX & vbCrLf
    Case 15
        Dim MXData As MX
        P = Record.RData
        MXData = GetMX(RawData, P)
        GetData = Right("0" & MXData.Preference, 2) & " - " & MXData.Exchange
    Case 17
        Dim RPData As RP
        P = Record.RData
        RPData = GetRP(RawData, P)
        GetData = GetData & "  MailBox    : " & RPData.MBox_DName & vbCrLf
        GetData = GetData & "  Text Domain: " & RPData.TXT_DName & vbCrLf
    Case 18
        Dim AFSDBData As AFSDB
        P = Record.RData
        AFSDBData = GetAFSDB(RawData, P)
        GetData = GetData & "  Sub-Type: " & AFSDBData.SubType & vbCrLf
        GetData = GetData & "  HostName: " & AFSDBData.HostName & vbCrLf
    Case 20
        Dim ISDNData As ISDN
        P = Record.RData
        ISDNData = GetISDN(RawData, P)
        GetData = GetData & "  ISDN-Address: " & ISDNData.Address & vbCrLf
        GetData = GetData & "  Sub-Address : " & ISDNData.SA & vbCrLf
    Case 21
        Dim RTData As RT
        P = Record.RData
        RTData = GetRT(RawData, P)
        GetData = Right("0" & RTData.Preference, 2) & " - " & RTData.Intermediate_Host
    Case 29
        Dim LOCData As LOC
        P = Record.RData
        LOCData = GetLOC(RawData, P)
        If LOCData.Version Then
            GetData = "Version " & LOCData.Version & " (Unsupported)"
        Else
            GetData = GetData & "  Version    : " & LOCData.Version & vbCrLf
            GetData = GetData & "  Size       : " & Int(LOCData.Size / 16) * (10 ^ (LOCData.Size Mod 16)) / 100 & " meters" & vbCrLf
            GetData = GetData & "  H Precision: ±" & Int(LOCData.Horiz_Pre / 16) * (10 ^ (LOCData.Horiz_Pre Mod 16)) / 200 & " meters" & vbCrLf
            GetData = GetData & "  V Precision: ±" & Int(LOCData.Vert_Pre / 16) * (10 ^ (LOCData.Vert_Pre Mod 16)) / 200 & " meters" & vbCrLf

            L = Abs(LOCData.Latitude - 2147483648#)
            D = Int(L / 3600000)
            M = Int((L Mod 3600000) / 60000)
            S = Format((L Mod 60000) / 1000, "0.000")
            GetData = GetData & "  Latitude   : " & D & "º " & M & "' " & S & IIf(LOCData.Latitude - 2147483648# > 0, """ N", """ S") & vbCrLf

            L = Abs(LOCData.Longitude - 2147483648#)
            D = Int(L / 3600000)
            M = Int((L Mod 3600000) / 60000)
            S = Format((L Mod 60000) / 1000, "0.000")
            GetData = GetData & "  Longitude  : " & D & "º " & M & "' " & S & IIf(LOCData.Longitude - 2147483648# > 0, """ E", """ W") & vbCrLf
            GetData = GetData & "  Altitude   : " & (LOCData.Altitude - 10000000) / 100 & " meters" & vbCrLf
        End If
    Case Else
        GetData = "Unsupported"
    End Select
End Function
