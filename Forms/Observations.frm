VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ObservationsForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Observations"
   ClientHeight    =   10290
   ClientLeft      =   20040
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "Observations.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10290
   ScaleMode       =   0  'User
   ScaleWidth      =   6020.548
   Begin VB.PictureBox ObsSnapShot 
      Height          =   615
      Index           =   0
      Left            =   7680
      ScaleHeight     =   555
      ScaleWidth      =   795
      TabIndex        =   41
      Top             =   240
      Width           =   855
   End
   Begin MSComctlLib.Toolbar PopupViewToolbar 
      Height          =   4140
      Left            =   1440
      TabIndex        =   23
      Top             =   5460
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   7303
      ButtonWidth     =   2090
      ButtonHeight    =   1799
      ImageList       =   "PopupViewImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Precision Vision"
            Key             =   "PVGPage"
            Object.ToolTipText     =   "Precision Vision Graph"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Observations"
            Key             =   "ObsPage"
            Object.ToolTipText     =   "Observations"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pipeline Details"
            Key             =   "PipeDetailsPage"
            Object.ToolTipText     =   "Pipeline Details"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            Key             =   "OptionsPage"
            Object.ToolTipText     =   "Options"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar PopupReportsToolbar 
      Height          =   5160
      Left            =   2280
      TabIndex        =   39
      Top             =   4440
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   9102
      ButtonWidth     =   2064
      ButtonHeight    =   1799
      ImageList       =   "ReportsImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "PV Graph"
            Key             =   "PVGraph"
            Object.ToolTipText     =   "PV Graph"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Analysis Graph"
            Key             =   "PVLandscape"
            Object.ToolTipText     =   "Analysis Graph"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "4 in 1"
            Key             =   "4_in_1"
            Object.ToolTipText     =   "4 PVGraphs on 1 Page"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Profile"
            Key             =   "PVPortrait"
            Object.ToolTipText     =   "Profile Snap-shot"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Multi Profile"
            Key             =   "MultiProfile"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Snap Shot"
            Key             =   "SnapShot"
            Object.ToolTipText     =   "Snap Shot Landscape"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ReportsImageList1 
      Left            =   4290
      Top             =   1215
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":2064
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":3D3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":5A18
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList PVImageList1 
      Left            =   4290
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":76F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":93CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":B0A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":CD80
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":EA5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":10734
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":1240E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":140E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":15DC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":17A9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":19776
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":1C7C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":1E4A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar PipeDetailsToolbar 
      Height          =   1680
      Left            =   0
      TabIndex        =   24
      Top             =   8580
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   2963
      ButtonWidth     =   1455
      ButtonHeight    =   1429
      ImageList       =   "PVImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OpenFile"
            Object.ToolTipText     =   "Open File"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveToFile"
            Object.ToolTipText     =   "Save To File"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ViewSelection"
            Object.ToolTipText     =   "View Selection"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reports"
            Object.ToolTipText     =   "Reports"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Information"
            Object.ToolTipText     =   "Information and Help"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList PopupViewImageList 
      Left            =   4290
      Top             =   615
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":2017C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":21E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":23B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Observations.frx":2580A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Recommendations_lbl 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1365
      Left            =   120
      TabIndex        =   22
      Top             =   7065
      Width           =   4065
      Begin VB.TextBox Recommendations 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1050
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   255
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   3015
      Left            =   4470
      TabIndex        =   14
      Top             =   5595
      Width           =   4065
      Begin MSComctlLib.StatusBar StatusBar1_old 
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   2640
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Style           =   1
         SimpleText      =   "No observations"
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   1
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Observation_old 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   990
         Width           =   3810
      End
      Begin VB.TextBox Distance_old 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2790
         TabIndex        =   2
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Observation_lbl_old 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Observation:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   3825
      End
      Begin VB.Label Distance_units_old 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   3645
         TabIndex        =   16
         Top             =   360
         Width           =   225
      End
      Begin VB.Label Distance_lbl 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Distance:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   465
         Left            =   135
         TabIndex        =   15
         Top             =   330
         Width           =   2580
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   2445
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   4065
      Begin VB.TextBox OutsideDiameter 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2400
         TabIndex        =   0
         Top             =   2025
         Width           =   1080
      End
      Begin VB.TextBox InternalDiameterExpected 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2400
         TabIndex        =   4
         Top             =   1530
         Width           =   1080
      End
      Begin VB.TextBox AssetNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2400
         TabIndex        =   8
         Top             =   150
         Width           =   1545
      End
      Begin VB.TextBox SiteID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2400
         TabIndex        =   7
         Top             =   495
         Width           =   1545
      End
      Begin VB.TextBox StartNodeNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2400
         TabIndex        =   6
         Top             =   840
         Width           =   1545
      End
      Begin VB.TextBox FinishNodeNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2400
         TabIndex        =   5
         Top             =   1185
         Width           =   1545
      End
      Begin VB.Label OutsideDiameter_lbl 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Outside Diameter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   480
         Left            =   90
         TabIndex        =   20
         Top             =   2025
         Width           =   2280
      End
      Begin VB.Label OutsideDiameter_units 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   3555
         TabIndex        =   19
         Top             =   2085
         Width           =   360
      End
      Begin VB.Label InternalDiameterExpected_lbl 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Internal Diameter (Expected)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   480
         Left            =   90
         TabIndex        =   18
         Top             =   1575
         Width           =   2280
      End
      Begin VB.Label InternalDiameterExpected_units 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   3555
         TabIndex        =   17
         Top             =   1605
         Width           =   360
      End
      Begin VB.Label FinishNodeNo_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Finish No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   480
         Left            =   90
         TabIndex        =   13
         Top             =   1185
         Width           =   2265
      End
      Begin VB.Label StartNodeNo_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Start No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   480
         Left            =   90
         TabIndex        =   12
         Top             =   840
         Width           =   2265
      End
      Begin VB.Label AssetNo_lbl 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Asset No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   465
         Left            =   90
         TabIndex        =   11
         Top             =   195
         Width           =   2325
      End
      Begin VB.Label SiteID_lbl 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Site ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   465
         Left            =   90
         TabIndex        =   10
         Top             =   540
         Width           =   2235
      End
   End
   Begin VB.Frame Observation_lbl 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4470
      Left            =   120
      TabIndex        =   26
      Top             =   2520
      Width           =   4065
      Begin VB.CommandButton AcceptObs 
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   2745
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Observations.frx":274E4
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3720
         Visible         =   0   'False
         Width           =   400
      End
      Begin VB.CheckBox AddNewFlag 
         Caption         =   "Check1"
         Height          =   210
         Left            =   2430
         TabIndex        =   37
         Top             =   4155
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.CommandButton AddNewObs 
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   2745
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Observations.frx":27BE6
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   4080
         Width           =   400
      End
      Begin VB.CommandButton DeleteObs 
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   3165
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Observations.frx":27EF8
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   4080
         Width           =   400
      End
      Begin VB.TextBox ObsFrameNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   3735
         Width           =   705
      End
      Begin VB.CommandButton ResetFixedPoints 
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Observations.frx":2820A
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   4080
         Width           =   400
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   375
         Left            =   45
         TabIndex        =   31
         Top             =   4035
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
         Style           =   1
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   1
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Observation 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   630
         Left            =   1035
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   3390
         Width           =   2970
      End
      Begin VB.TextBox Distance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   30
         TabIndex        =   28
         Top             =   3405
         Width           =   705
      End
      Begin VB.ListBox ObservationsList 
         Height          =   2985
         Left            =   30
         MultiSelect     =   2  'Extended
         TabIndex        =   27
         Top             =   270
         Width           =   3990
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Fr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   765
         TabIndex        =   34
         Top             =   3810
         Width           =   225
      End
      Begin VB.Label Distance_units 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   765
         TabIndex        =   29
         Top             =   3465
         Width           =   225
      End
   End
   Begin VB.Label Recommendations_lbl_old 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   4830
      TabIndex        =   38
      Top             =   2490
      Width           =   3825
   End
End
Attribute VB_Name = "ObservationsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''Dim DataEntryChange As Boolean 'PCN1768
'''
'''
'''Private Sub AcceptObs_Click()
''''****************************************************************************************
''''Name    : AddNewObs_Click
''''Created : 29 July 2004, PCN2928
''''Updated :
''''Prg By  : Geoff Logan
''''Param   :
''''Desc    : Add a new observation to the PipeObservations
''''Usage   :
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''
'''If AddNewFlag = 1 Then
'''    Call ObservationAddNew
'''Else
'''    'Update current PipeObservation record
'''    Call ObservationUpdatePipeObs(ObservationsList.ListIndex + 1)
'''End If
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox error$
'''End Sub
'''
'''Private Sub AddNewObs_Click()
''''****************************************************************************************
''''Name    : AddNewObs_Click
''''Created : 29 July 2004, PCN2928
''''Updated :
''''Prg By  : Geoff Logan
''''Param   :
''''Desc    : Add a new observation to the PipeObservations
''''Usage   :
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''
'''    Call SetupNewObs 'PCN3576
''''AddNewFlag = 1
''''ObsFrameNo.text = PVFrameNo
''''Distance.text = ""
''''Observation.text = ""
'''
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox error$
'''End Sub
'''
'''
'''Private Sub DeleteObs_Click()
''''****************************************************************************************
''''Name    : DeleteObs_Click
''''Created : 29 July 2004, PCN2928
''''Updated :
''''Prg By  : Geoff Logan
''''Param   :
''''Desc    : Delete the Pipe Observation at current position of the List
''''Usage   :
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''Dim SelectedIndex As Integer
'''
'''Call ObservationDelete(ObservationsList.ListIndex + 1)
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox error$
'''End Sub
'''
'''Private Sub Distance_Validate(Cancel As Boolean)
''''****************************************************************************************
''''Name    : Distance_Validate
''''Created : 29 July 2004, PCN2928
''''Updated :
''''Prg By  : Geoff Logan
''''Param   :
''''Desc    : Update or add a new observation to the Pipe Observations
''''Usage   :
''''****************************************************************************************
''''On Error GoTo Err_Handler
''''
''''If AddNewFlag = 1 Then
''''    Call ObservationAddNew
''''Else
''''    'Update current PipeObservation record
''''    Call ObservationUpdatePipeObs(ObservationsList.ListIndex + 1)
''''End If
''''
''''Exit Sub
''''Err_Handler:
''''    MsgBox error$
'''
'''End Sub
'''Private Sub Form_Activate()
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''''PCN     : PCN3576
''''Name    : Form_Activate
''''Created : 13 July 2005
''''Updated :
''''Prg By  : Antony van Iersel
''''Param   : None
''''Desc    : When the obs form is activated then default setting have to be set up
''''          eg, frame number, and distance if available
''''Usage   :
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''On Error GoTo Err_Handler
'''    If ObservationsList.SelCount <> 0 Then Exit Sub  'If an observation is all ready selected then dont
'''                                                     'populate the Observation textbox
'''
'''    If PVDataNoOfLines < 1 Then Exit Sub
'''
'''
'''    ObsFrameNo.text = PVFrameNo
'''
'''    If DistanceMethod <> "None" Then
'''        Distance.text = PVDistances(PVFrameNo)
'''    Else
'''        Distance.text = ""
'''    End If
'''
''''    Observation.text = ""
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox Err & " - " & error$
'''End Sub
'''
'''
'''Private Sub Form_KeyPress(KeyAscii As Integer) 'PCN1894
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''''Form_KeyPress Sub  Michelle Lindsay michellelindsay@cbsys.co.nz
''''
''''Revision history"
''''   V0.0    Michelle Lindsay,    04/03/2003     Building initial framework
''''
''''Description:
''''Allows the user to press the escape key to close any popups opened in error.
''''
''''Purpose:
''''
''''
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''On Error GoTo Err_Handler
'''
'''Select Case KeyAscii
'''    Case vbKeyEscape
'''     If (PopupReportsToolbar.Visible = True) Then
'''       PopupReportsToolbar.Visible = False
'''     ElseIf (PopupViewToolbar.Visible = True) Then
'''       PopupViewToolbar.Visible = False
'''     End If
'''    Case Else
'''End Select
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox error$
'''
'''End Sub
'''
''''Private Sub Distance_Validate(Cancel As Boolean)
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''''Form_Load Sub  Michelle Lindsay michellelindsay@cbsys.co.nz
''''
''''Revision history
''''   V0.0    Michelle Lindsay,   23/01/2003   Adding functionality
''''
''''Description:
''''
''''Purpose: Validate the distance field to ensure that it is numeric.
''''
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''''On Error GoTo Err_handler
'''
''''If Not IsNumeric(Distance) Then
''''    MsgBox "The distance must be numeric", vbExclamation, "Distance invalid"
''''    Cancel = True
''''Else
''''    PipeObservations.PipeObsDist = Distance
'''
''''End If
''''Exit Sub
'''
''''Err_handler:
''''    MsgBox Error$
'''
''''End Sub
'''
'''Private Sub Form_Load()
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''''Form_Load Sub  Michelle Lindsay michellelindsay@cbsys.co.nz
''''
''''Revision history
''''   V0.0    Michelle Lindsay,   19/12/02    Adding functionality
''''
''''Description:
''''
''''Purpose:
''''
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''On Error GoTo Err_Handler
'''ConvertLanguage Me, Language 'PCN2111
'''PopupViewToolbar.width = PopupViewToolbar.ButtonWidth 'PCN2138
'''PopupReportsToolbar.width = PopupReportsToolbar.ButtonWidth 'PCN2138
'''
'''Observations.Left = PVPageLeft 'PCNGL060103
'''Observations.Top = PVPageTop
'''Observations.width = PVPageWidth
'''Observations.height = PVPageHeight
'''
'''If (IdentifyOperatingSystem = "Windows XP") Then 'ML120203
'''    PipeDetailsToolbar.Top = 8450
'''    Recommendations_lbl.Top = 7000
'''    Observation_lbl.Top = 2520 'PCN2928
'''    Frame1.Left = 0
'''    Recommendations_lbl.Left = 0
'''    Observation_lbl.Left = 0 'PCN2928
'''End If
'''
'''If MeasurementUnits = "mm" Then
'''    InternalDiameterExpected_units.Caption = "mm"
'''    OutsideDiameter_units.Caption = "mm"
'''    Distance_units.Caption = "m"
'''Else
'''    InternalDiameterExpected_units.Caption = "in"
'''    OutsideDiameter_units.Caption = "in"
'''    Distance_units.Caption = "ft"
'''End If
'''
'''AcceptObs.Top = AddNewObs.Top 'PCN3576
'''
'''
''''Load the form data from the PipelineInfo array 'PCNGL130103
'''Call CopyPipeDetailsToObsForm
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox error$
'''End Sub
'''
'''
'''Private Sub Observation_Click()
'''    If ObservationsList.SelCount = 0 And Observation.text = "" And AcceptObs.Visible = False Then
'''        Call SetupNewObs 'PCN3576
'''    End If
'''    If ObservationsList.SelCount <> 0 And AcceptObs.Visible = False Then
'''        Call AddjustCurrentObs
'''    End If
'''End Sub
'''Private Sub Observation_Validate(Cancel As Boolean)
''''****************************************************************************************
''''Name    : Observation_Validate
''''Created : 29 July 2004, PCN2928
''''Updated :
''''Prg By  : Geoff Logan
''''Param   :
''''Desc    : Update or add a new observation to the Pipe Observations
''''Usage   :
''''****************************************************************************************
''''On Error GoTo Err_Handler
''''
''''If AddNewFlag = 1 Then
''''    Call ObservationAddNew
''''Else
''''    'Update current PipeObservation record
''''    Call ObservationUpdatePipeObs(ObservationsList.ListIndex + 1)
''''End If
''''
''''Exit Sub
'''
'''
''''Err_Handler:
''''    MsgBox error$
'''End Sub
'''
'''Private Sub ObservationsList_Click()
''''****************************************************************************************
''''Name    : ObservationsList_Click
''''Created : 29 July 2004, PCN2928
''''Updated :
''''Prg By  : Geoff Logan
''''Param   :
''''Desc    :
''''Usage   :
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''
'''AddNewFlag = 0
'''    Call ShowAddNewObsButton
'''Call ObservationTextUpdate(ObservationsList.ListIndex + 1)
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox error$
'''End Sub
'''
'''
''''Private Sub InternalDiameterExpected_Validate(Cancel As Boolean)
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''''InternalDiameterExpected Sub  Michelle Lindsay michellelindsay@cbsys.co.nz
''''
''''Revision history"
''''   V0.0    Michelle Lindsay,    23/01/2003     Building initial framework
''''
''''Description:
''''
''''Purpose:Validates the InternalDiameterExpected field
''''
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''''On Error GoTo Err_handler
''''If Not IsNumeric(InternalDiameterExpected) Then
''''    MsgBox "The internal diameter must be numeric", vbExclamation, "Internal diameter invalid"
''''    Cancel = True
''''Else
'''
'''
''''End Sub
'''
'''Private Sub PipeDetailsToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''''PipeDetailsToolbar_ButtonClick Sub  Geoff Logan geofflogan@cbsys.co.nz
''''
''''Revision history"
''''   V0.0    Geoff Logan,    20/11/02     Building initial framework
''''
''''Description:
''''
''''Purpose:
''''
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''On Error GoTo Err_Handler
'''
''''Call FisheyeFunctions.LiveFishEyeOFF
''''Action on button press
'''
''''PCN3217
''''PopupReportsToolbar.Visible = False
''''PopupViewToolbar.Visible = False
'''Select Case Button.Key
'''    Case "OpenFile"
'''        Call OpenAnyFile("") 'PCN2133
'''    Case "SaveToFile"
'''        If Registered = False Then 'PCN1956
'''            MsgBox DisplayMessage("Cannot save a .PVD file, please register the software to access this."), vbExclamation 'PCN1956 'PCN2111
'''        Else
'''            Call SaveImageAndOrData 'PCNGL110103
'''        End If
'''    Case "Information"
'''        Dim hwndHelp As Long
'''        If HelpFilename <> "" And Dir(ReadOnlyAppPath & "Language\" & HelpFilename) = "" Then 'Check whether that file exists actually.'PCN2167 7/8/03 by Abe
'''            MsgBox HelpFilename & " " & DisplayMessage("file for the language") & _
'''                "(" & Language & ") " & _
'''                DisplayMessage("does not exist. Create this file first. The default language(English) is loaded for Help file."), , "Clear Line Profiler" 'PCN2167 7/8/03 by Abe, PCN2171
'''        End If
'''        If HelpFilename <> "" And Dir(ReadOnlyAppPath & "Language\" & HelpFilename) <> "" Then  'PCN2167 7/8/03 by Abe ---------v
'''            hwndHelp = HtmlHelp(hwnd, ReadOnlyAppPath & "Language\" & HelpFilename, HH_DISPLAY_TOPIC, 0)
'''        Else '-------------------------------------------------------------------------------------------^
'''            'PCN 1972 LS 8/7/03
'''            'hwndHelp = HtmlHelp(hwnd, App.Path & "\HelpFile.chm", HH_DISPLAY_TOPIC, 0)
'''            hwndHelp = HtmlHelp(hwnd, LocToSave & "HelpFile.chm", HH_DISPLAY_TOPIC, 0)
'''        End If
'''    Case "ViewSelection"
'''        'PCN3217
'''        Call ScreenDrawing.ToggleViewSelectionPopUp(Observations)
'''        'PopupViewToolbar.Visible = True
'''        'PopupReportsToolbar.Visible = False
'''    Case "Reports"
'''        'PCN3217
'''        Call ScreenDrawing.ToggleReportsPopUp(Observations)
'''        'PopupReportsToolbar.Visible = True
'''        'PopupViewToolbar.Visible = False
''''Removed
''''    Case "ImageProcessing" 'PCN2463
''''        AutoTune.Show
''''        AutoTune.ZOrder 0
''''    Case "Distance" 'PCN2463
''''        Load Distance
''''        Distance.ZOrder 0
''''    Case "FishEye" 'FISH-EYE PCN2290
''''        Fisheye.Show
''''        Fisheye.ZOrder 0
'''    Case Else
'''    End Select
'''
'''Exit Sub
'''Err_Handler:
'''Select Case Err
'''    Case 53 'File not found
'''        Close #1
'''        Exit Sub
'''    Case 54 'Bad file mode
'''        Close #1
'''        Exit Sub
'''    Case 62 'past end of file
'''        Close #1
'''        Exit Sub
'''    Case Else
'''        MsgBox Err & " - " & error$
'''End Select
'''End Sub
'''
'''Private Sub PopupReportsToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''''PopupupReportsToolbar_ButtonClick Sub Michelle Lindsay michellelindsay@cbsys.co.nz
''''
''''Revision history
''''   V0.0    Michelle Lindsay,   12/12/02    Building initial framework
''''
''''Description:
''''       Event for showing correct reports when the user click a button on the
''''       pop-up toolbar.
''''
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''
'''On Error GoTo Err_Handler
'''
'''Select Case Button.Key
'''    'Case SnapShot  - Not currently required, ML06.01.03
'''    '    PopupReportsToolbar.Visible = False
'''    '    If isopen("LaserImageReport") Then Unload LaserImageReport
'''    '    LaserImageReport.Show
'''    Case "PVPortrait"
'''        'Call PVPortraitReport(Observations) 'PCN2777
'''        Load PVReportProfile
'''    Case "PVLandscape" 'PCNGL040103
'''        Call PVLandscapeReport(Observations)
'''    Case "PVGraph" 'PCN2777
'''        'Call PVLandscapeSingleReport(Observations) 'PCN2777
'''        Load PVReportSingle
'''    Case "4_in_1"
'''        'This report shows four different PVGraphs (eg Capacity, Ovality,
'''        'Delta or XY and the Flat) on the same page.
'''        Call PV4_In_1Report(Observations)
'''    '^^^^ ****************************************************************
'''    'vvvv PCN3479 **************************************************
'''    Case "MultiProfile"
'''        'This report is a multi profile report and is based on the
'''        'Profiles at all of the Observations points.
'''        Call PVMultiProfileReport(Observations)
'''    '^^^^ **********************************************************
'''    Case Else
'''
'''End Select
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox error$
'''End Sub
'''
'''Private Sub PopupViewToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''''Name    : PopupViewToolbar_ButtonClick
''''Created : 12 November 2002,
''''Updated : 18 November 2003, PCN2402 Tidy up of form layout for XP
''''Prg By  : Geoff Logan
''''Param   :
''''Desc    :
''''Usage   :
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''On Error GoTo Err_Handler
'''
''''Action on button press
'''PopupViewToolbar.Visible = False
'''Select Case Button.Key
'''    Case "PipeDetailsPage"
'''        Load PipelineDetails
'''        PipelineDetails.Show
'''        PipelineDetails.ZOrder 0
'''    Case "ObsPage"
'''    Case "PVGPage"
'''        Load PrecisionVisionGraph
'''        PrecisionVisionGraph.Show
'''        PrecisionVisionGraph.ZOrder 0
''''        Call DrawPVYScale(1) 'PCNLS200203
'''    Case "OptionsPage"
'''        Load OptionsPage
'''        OptionsPage.Show
'''        OptionsPage.ZOrder 0
'''    Case Else
'''End Select
'''
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox error$
'''End Sub
'''
'''
'''Private Sub Recommendations_KeyUp(KeyCode As Integer, Shift As Integer) 'PCN1768
'''On Error GoTo Err_Handler
'''Dim SaveFailed As Boolean
'''
'''DataEntryChange = True
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox Err & error$
'''End Sub
'''
'''Private Sub Recommendations_Validate(Cancel As Boolean) 'PCN1768
'''On Error GoTo Err_Handler
'''Dim SaveFailed As Boolean
'''
'''If DataEntryChange = True And PVDFileName <> "" Then
'''    PipelineInfo.Comments = Me.Recommendations
'''    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed)
'''End If
'''
'''DataEntryChange = False
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox Err & error$
'''End Sub
'''
'''Function ObservationsListLoad()
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''''Name    : ObservationsListLoad
''''Created : 29 July 2004, PCN2928
''''Updated :
''''Prg By  : Geoff Logan
''''Param   :
''''Desc    : Loads the ObservationList with PipeObservation.
''''Usage   :
''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''On Error GoTo Err_Handler
'''Dim ObsIndex As Integer
'''Dim ObsListStr As String
'''
'''ObservationsList.Clear
''''Load all Observations
'''For ObsIndex = 1 To NoOfPipeObservations
'''    If MeasurementUnits = "mm" Then
'''        ObsListStr = "Fr " & PipeObservations(ObsIndex).PipeObsFrameNo & ", " & Format(PipeObservations(ObsIndex).PipeObsDist, "#0.0") & Distance_units.Caption & ", " & PipeObservations(ObsIndex).PipeObs
'''    Else
'''        ObsListStr = "Fr " & PipeObservations(ObsIndex).PipeObsFrameNo & ", " & Format(PipeObservations(ObsIndex).PipeObsDist, "#0") & Distance_units.Caption & ", " & PipeObservations(ObsIndex).PipeObs
'''    End If
'''    ObservationsList.AddItem ObsListStr
'''Next ObsIndex
'''
'''StatusBar1.SimpleText = (ObservationsList.ListCount) & " " & Observation_lbl.Caption
'''
'''Exit Function
'''Err_Handler:
'''    MsgBox Err & error$
'''End Function
'''
'''Private Sub ResetFixedPoints_Click()
''''****************************************************************************************
''''Name    : ResetFixedPoints_Click
''''Created : 29 July 2004, PCN2928
''''Updated :
''''Prg By  : Geoff Logan
''''Param   :
''''Desc    : Deletes all of the DistanceFixedPoints from the PipeObservation array.
''''Usage   :
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''
'''Call DistanceFixedPtReset
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox Err & error$
'''End Sub
'''
'''Function DistanceFixedPtReset()
''''****************************************************************************************
''''Name    : DistanceFixedPtReset
''''Created : 29 July 2004, PCN2928
''''Updated :
''''Prg By  : Geoff Logan
''''Param   :
''''Desc    : Deletes all of the DistanceFixedPoints from the PipeObservation array. It also
''''          reorganizes the other PipeObservations in order to consolidate the array dimension size.
''''Usage   :
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''Dim ObsIndex As Integer 'The Fixed Point index number
''''Dim PipeObsTemp() As PipeObservationType_V50 'Remember to update when V50 is changed
'''Dim PipeObsTemp() As PipeObservationType_V60
'''Dim NewNoOfPipeObservations As Integer
'''Dim FileSaveFail As Boolean
'''
'''NewNoOfPipeObservations = 0
'''
'''For ObsIndex = 1 To NoOfPipeObservations
''''    If InStr(1, PipeObservations(ObsIndex).PipeObs, "<<<--I-->>>") <> 0 Then
'''    If Left(PipeObservations(ObsIndex).PipeObs, 11) <> "<<<--I-->>>" Then
'''        NewNoOfPipeObservations = NewNoOfPipeObservations + 1
'''        ReDim Preserve PipeObsTemp(NewNoOfPipeObservations)
'''        PipeObsTemp(NewNoOfPipeObservations).PipeObsFrameNo = PipeObservations(ObsIndex).PipeObsFrameNo
'''        PipeObsTemp(NewNoOfPipeObservations).PipeObsDist = PipeObservations(ObsIndex).PipeObsDist
'''        PipeObsTemp(NewNoOfPipeObservations).PipeObs = PipeObservations(ObsIndex).PipeObs
'''        PipeObsTemp(NewNoOfPipeObservations).PipeObsSnapshotLength = PipeObservations(ObsIndex).PipeObsSnapshotLength ' PC3576
'''        PipeObsTemp(NewNoOfPipeObservations).PipeObsSnapshotOffset = PipeObservations(ObsIndex).PipeObsSnapshotOffset ' PCN3576
'''    End If
'''Next ObsIndex
'''
''''Reset PipeObservations
'''NoOfPipeObservations = NewNoOfPipeObservations
'''ReDim PipeObservations(NoOfPipeObservations)
''''Reload the Pipe Observations
'''For ObsIndex = 1 To NoOfPipeObservations
'''    PipeObservations(ObsIndex).PipeObsFrameNo = PipeObsTemp(ObsIndex).PipeObsFrameNo
'''    PipeObservations(ObsIndex).PipeObsDist = PipeObsTemp(ObsIndex).PipeObsDist
'''    PipeObservations(ObsIndex).PipeObs = PipeObsTemp(ObsIndex).PipeObs
'''    PipeObservations(ObsIndex).PipeObsSnapshotLength = PipeObsTemp(ObsIndex).PipeObsSnapshotLength ' PCN3576
'''    PipeObservations(ObsIndex).PipeObsSnapshotOffset = PipeObsTemp(ObsIndex).PipeObsSnapshotOffset ' PCN3576
'''Next ObsIndex
'''
''''Write back to the PVD
'''Call SaveToFilePipeObs(FileSaveFail)
'''
''''Reset the ObservationsList
'''Call ObservationsListLoad
'''
''''Reset the Observations text boxes
'''Observations.Observation.text = ""
'''Observations.Distance.text = ""
'''Observations.ObsFrameNo = ""
'''
'''Exit Function
'''Err_Handler:
'''Select Case Err
'''    Case Else
'''        MsgBox Err & "-" & error$
'''End Select
'''End Function
'''
'''Function ObservationTextUpdate(PipeObsIndexNo As Integer)
''''****************************************************************************************
''''Name    : ObservationTextUpdate
''''Created : 29 July 2004, PCN2928
''''Updated :
''''Prg By  : Geoff Logan
''''Param   :
''''Desc    : Updates the Observations text boxes and moves to the PipeObs Frame No.
''''Usage   : Called by the ObservationList click event.
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''
'''
'''
'''If PipeObsIndexNo = 0 Then Exit Function
'''
'''Observations.Distance = PipeObservations(PipeObsIndexNo).PipeObsDist
'''Observations.Observation = Trim(PipeObservations(PipeObsIndexNo).PipeObs)
'''Observations.ObsFrameNo = PipeObservations(PipeObsIndexNo).PipeObsFrameNo
'''
''''PCN3576 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Call ExtractObservationSnapshotAtIndex(PipeObsIndexNo)
'''
'''If CDbl(Observations.Distance) = InvalidData Then Observations.Distance = "" 'PCN3597
''''Move to this PVFrameNo
'''PVFrameNo = PipeObservations(PipeObsIndexNo).PipeObsFrameNo
''''vvvv PCN2930 ***************************
'''If CLPScreenMode = PV And PVFrameNo <> 0 Then
'''    Call ClearLineScreen.GotoPVProfile(PVFrameNo, True)
'''''Else
'''''    Call GotoPVGraphProfile(PVFrameNo) 'Move PV Frame to new position
'''End If
''''^^^^ ***********************************
'''
'''If mediatype = Video Then
'''    If CheckAVIInitialised = True Then 'Check that the AVI is correctly initialised before running the C code
'''        Call ClearLineScreen.GotoAVIFrame(PVDFileName, PVFrameNo, 1)
'''    End If
''''    CurrentTime = PVTimes(PVFrameNo)
''''    Call ClearLineScreen.SeekTime(CurrentTime)
''''    VideoFrameSlider.value = (CurrentTime / AVITime) * VideoFrameSlider.Max 'Calculate new slider value
''''    SliderFrame = VideoFrameSlider.value 'Set Slider Value
''''    Call VideoFrameSliderMove
''''    AVITimeVar = Round((VideoFrameSlider.value / VideoFrameSlider.Max) * AVITime, 1)
''''    Call FormatTime
'''End If
'''
'''
'''
'''
'''
'''Exit Function
'''Err_Handler:
'''    MsgBox error$
'''End Function
'''
'''Function ObservationDelete(PipeObsIndexNo As Integer)
''''****************************************************************************************
''''Name    : ObservationDelete
''''Created : 29 July 2004, PCN2928
''''Updated :
''''Prg By  : Geoff Logan
''''Param   :
''''Desc    : Delete the Pipe Observation at PipeObsIndex
''''Usage   :
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''Dim ObsIndex As Integer 'The Observation index number
''''Dim PipeObsTemp() As PipeObservationType_V50 'Remember to update when V50 is changed
'''Dim PipeObsTemp() As PipeObservationType_V60 ' PCN3576
'''
'''Dim NewNoOfPipeObservations As Integer
'''Dim FileSaveFail As Boolean
'''
'''If PipeObsIndexNo = 0 Then Exit Function
'''If ObservationsList.SelCount = 0 Then Exit Function 'PCN3576
'''
'''ReDim PipeObsTemp(NoOfPipeObservations - 1)
'''NewNoOfPipeObservations = 0
'''
'''For ObsIndex = 1 To NoOfPipeObservations
'''    'Delete all selected records
'''    If ObsIndex <= ObservationsList.ListCount Then
'''        If ObservationsList.Selected(ObsIndex - 1) = False Then
'''    '    If ObsIndex <> PipeObsIndexNo Then
'''            NewNoOfPipeObservations = NewNoOfPipeObservations + 1
'''            PipeObsTemp(NewNoOfPipeObservations).PipeObsFrameNo = PipeObservations(ObsIndex).PipeObsFrameNo
'''            PipeObsTemp(NewNoOfPipeObservations).PipeObsDist = PipeObservations(ObsIndex).PipeObsDist
'''            PipeObsTemp(NewNoOfPipeObservations).PipeObs = PipeObservations(ObsIndex).PipeObs
'''            PipeObsTemp(NewNoOfPipeObservations).PipeObsSnapshotLength = PipeObservations(ObsIndex).PipeObsSnapshotLength  ' PCN3576
'''            PipeObsTemp(NewNoOfPipeObservations).PipeObsSnapshotOffset = PipeObservations(ObsIndex).PipeObsSnapshotOffset  ' PCN3576
'''        End If
'''    End If
'''Next ObsIndex
'''
''''Reset PipeObservations
'''NoOfPipeObservations = NewNoOfPipeObservations
'''ReDim PipeObservations(NoOfPipeObservations)
''''Reload the Pipe Observations
'''For ObsIndex = 1 To NoOfPipeObservations
'''    PipeObservations(ObsIndex).PipeObsFrameNo = PipeObsTemp(ObsIndex).PipeObsFrameNo
'''    PipeObservations(ObsIndex).PipeObsDist = PipeObsTemp(ObsIndex).PipeObsDist
'''    PipeObservations(ObsIndex).PipeObs = PipeObsTemp(ObsIndex).PipeObs
'''    PipeObservations(ObsIndex).PipeObsSnapshotLength = PipeObsTemp(ObsIndex).PipeObsSnapshotLength  ' PCN3576
'''    PipeObservations(ObsIndex).PipeObsSnapshotOffset = PipeObsTemp(ObsIndex).PipeObsSnapshotOffset  ' PCN3576
'''Next ObsIndex
'''
''''Write back to the PVD
'''Call SaveToFilePipeObs(FileSaveFail)
'''
''''Reset the ObservationsList
'''Call ObservationsListLoad
'''
''''Reset the Observations text boxes
'''Observations.Observation.text = ""
'''Observations.Distance.text = ""
'''Observations.ObsFrameNo = ""
'''
'''Exit Function
'''Err_Handler:
'''    MsgBox error$
'''End Function
'''
'''
'''Function ObservationAddNew()
''''****************************************************************************************
''''Name    : ObservationAcceptNew
''''Created : 29 July 2004, PCN2928
''''Updated :
''''Prg By  : Geoff Logan
''''Param   :
''''Desc    : Add a new observation to the Pipe Observations
''''Usage   :
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''Dim CheckDistValue As Double
'''Dim FileSaveFail As Boolean
'''Dim PipeObsSnapshotOffset As Long
'''Dim PipeObsSnapshotLength As Long
'''Dim SnapshotFilename As String
'''
'''
'''If AddNewFlag <> 1 Then Exit Function
''''Ensure NoOfPipeObservations does not exceed the PipeObsBuffer
'''If NoOfPipeObservations >= PipeObsBuffer Then Exit Function
'''
''''PCN3576
'''If Distance.text = "" Then
'''    Call AttentionTextBox(Distance)
'''    Exit Function
'''End If
'''
'''If Observation.text = "" Then
'''    Call AttentionTextBox(Observation)
'''    Exit Function
'''End If
'''
'''
'''On Error GoTo InvalidData
'''CheckDistValue = CDbl(Observations.Distance.text)
'''On Error GoTo Err_Handler
''''Check there is a valid observation
'''If Len(Observations.Observation) = 0 Then Exit Function
''''Store this new setting in the Pipe Observation array
'''NoOfPipeObservations = NoOfPipeObservations + 1
'''ReDim Preserve PipeObservations(NoOfPipeObservations)
'''PipeObservations(NoOfPipeObservations).PipeObs = Observations.Observation
'''PipeObservations(NoOfPipeObservations).PipeObsFrameNo = PVFrameNo
'''PipeObservations(NoOfPipeObservations).PipeObsDist = CheckDistValue
'''
''''PCN3576 '''''''''' Disabling for now ''''''''''''''''''''''''''''''''''
''''If VideoFileName = "" Then
'''    PipeObservations(NoOfPipeObservations).PipeObsSnapshotOffset = 0
'''    PipeObservations(NoOfPipeObservations).PipeObsSnapshotLength = 0
''''Else
''''    SnapshotFilename = LocToSave & "Snapshot.bmp"
''''
''''    If CLPScreenMode = PV Then
''''        SavePicture ClearLineScreen.PVScreen.Image, SnapshotFilename
''''    ElseIf CLPScreenMode = Video And VideoSnapShotMode = Video Then
''''        ClearLineScreen.Snap
''''    ElseIf VideoSnapShotMode = SnapShot Then
''''        SavePicture ClearLineScreen.SnapShotScreen.Image, SnapshotFilename
''''    End If
''''
''''    Call PageFunctions.EmbedFile(SnapshotFilename, PipeObsSnapshotOffset, PipeObsSnapshotLength)
''''    PipeObservations(NoOfPipeObservations).PipeObsSnapshotOffset = PipeObsSnapshotOffset
''''    PipeObservations(NoOfPipeObservations).PipeObsSnapshotLength = PipeObsSnapshotLength
''''End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''
''''Write back to the PVD
'''Call SaveToFilePipeObs(FileSaveFail)
''''Reset the ObservationsList
'''Call ObservationsListLoad
''''Reset AddNewFlag
'''AddNewFlag = 0
'''DoEvents
'''ObservationsList.ListIndex = ObservationsList.ListCount - 1
'''
''''PCN3576 ''''''''''''''''''''
'''Call ShowAddNewObsButton
'''ObservationsList.Selected(ObservationsList.ListCount - 1) = True
''''''''''''''''''''''''''''''''
'''
'''InvalidData:
'''    'Do nothing
'''Exit Function
'''Err_Handler:
'''    MsgBox error$
'''End Function
'''
'''Function ObservationUpdatePipeObs(PipeObsIndexNo As Integer)
''''****************************************************************************************
''''Name    : ObservationUpdatePipeObs
''''Created : 29 July 2004, PCN2928
''''Updated :
''''Prg By  : Geoff Logan
''''Param   :
''''Desc    : If the Observation information has be edited, then update PipeObservations.
''''Usage   :
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''Dim CheckDistValue As Double
'''Dim FileSaveFail As Boolean
'''
'''If PipeObsIndexNo = 0 Then Exit Function
'''
'''On Error GoTo InvalidData
'''CheckDistValue = CDbl(Observations.Distance.text)
'''On Error GoTo Err_Handler
''''Check there is a valid observation
'''If Len(Observations.Observation) = 0 Then Exit Function
''''Update values in Pipe Observations
'''PipeObservations(PipeObsIndexNo).PipeObsDist = Observations.Distance
'''PipeObservations(PipeObsIndexNo).PipeObs = Observations.Observation
''''Write back to the PVD
'''Call SaveToFilePipeObs(FileSaveFail)
''''Reset the ObservationsList
'''Call ObservationsListLoad
'''ObservationsList.ListIndex = (PipeObsIndexNo - 1)
'''
''''PCN3576 ''''''''''''''''''''
'''Call ShowAddNewObsButton
'''ObservationsList.Selected(PipeObsIndexNo - 1) = True
''''''''''''''''''''''''''''''''
'''
'''InvalidData:
'''    'Do Nothing
'''Exit Function
'''Err_Handler:
'''    MsgBox error$
'''End Function
'''
'''Sub RemoveAllDistancesAndObservations()
''''****************************************************************************************
''''PCN     : PCN3338
''''Name    : ObservationUpdatePipeObs
''''Created : 8 July 2005
''''Updated :
''''Prg By  : Antony van Iersel
''''Param   :
''''Desc    : Removes all distances and observation
''''Usage   : When loading a new video for profiling, the old observations and distances needed to be deleted
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''
'''Dim Count As Integer
'''
'''Call DistanceFixedPtReset
'''For Count = 1 To NoOfPipeObservations
'''    ObservationsList.Selected(0) = True
'''    Call ObservationDelete(1)
'''Next Count
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox error$
'''End Sub
'''
'''Sub ToggleAddAndAcceptObs()
''''****************************************************************************************
''''PCN     : PCN3576
''''Name    : ToggleAddAndAcceptObs
''''Created : 13 July 2005
''''Updated :
''''Prg By  : Antony van Iersel
''''Param   :
''''Desc    : Toggle the visiblity of the Add and Accept observation buttons
''''Usage   : When there is an observation is edited in any way then its a accept
''''          otherwise its a add observation
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''
'''    If AcceptObs.Visible = False Then
'''        AcceptObs.Visible = True
'''        AddNewObs.Visible = False
'''    Else
'''        AcceptObs.Visible = False
'''        AddNewObs.Visible = True
'''    End If
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox error$
'''End Sub
'''
'''Sub ShowAddNewObsButton()
''''****************************************************************************************
''''PCN     : PCN3576
''''Name    : ShowAddNewObsButton
''''Created : 13 July 2005
''''Updated :
''''Prg By  : Antony van Iersel
''''Param   :
''''Desc    : Shows the AddObs button regardles if AddNewObs or AcceptObs are visible
''''Usage   : When there is an observation is edited in any way then its a accept
''''          otherwise its a add observation
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''
'''    AcceptObs.Visible = False
'''    AddNewObs.Visible = True
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox error$
'''End Sub
'''
'''Sub ShowAcceptObs()
''''****************************************************************************************
''''PCN     : PCN3576
''''Name    : ShowAcceptObs
''''Created : 13 July 2005
''''Updated :
''''Prg By  : Antony van Iersel
''''Param   :
''''Desc    : Shows the AcceptObs button regardless if the AddNewObs or AcceptObs are visbile
''''Usage   : When there is an observation is edited in any way then its a accept
''''          otherwise its a add observation
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''
'''    AcceptObs.Visible = True
'''    AddNewObs.Visible = False
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox error$
'''End Sub
'''
'''Sub DeselectListBox(TheListBox As ListBox)
'''
''''****************************************************************************************
''''PCN     : PCN3576
''''Name    : DeselectListBox
''''Created : 14 July 2005
''''Updated :
''''Prg By  : Antony van Iersel
''''Param   :
''''Desc    : Completly Deselects listbox
''''Usage   : Only to deselect listboxes
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''
'''Dim NumberOfItems As Integer
'''Dim Count As Integer
'''
'''NumberOfItems = TheListBox.ListCount - 1
'''If NumberOfItems < 0 Then Exit Sub
'''Dim TextList() As String
'''ReDim TextList(NumberOfItems)
'''
'''For Count = 0 To NumberOfItems
'''    TextList(Count) = TheListBox.List(0)
'''    TheListBox.RemoveItem (0)
'''Next Count
'''
'''For Count = 0 To NumberOfItems
'''    TheListBox.AddItem (TextList(Count))
'''Next Count
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox error$
'''End Sub
'''
'''
'''Sub SetupNewObs()
'''
''''****************************************************************************************
''''PCN     : PCN3576
''''Name    : SetupNewObs
''''Created : 14 July 2005
''''Updated :
''''Prg By  : Antony van Iersel
''''Param   :
''''Desc    : Sets up the observations form to accept a new observations
''''Usage   :
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''
'''    AddNewFlag = 1
'''    ObsFrameNo.text = PVFrameNo
'''    Distance.text = ""
'''    Observation.text = ""
'''
'''    ObsFrameNo.text = PVFrameNo
'''    If DistanceMethod <> "None" Then
'''        Distance.text = PVDistances(PVFrameNo)
'''    Else
'''        Distance.text = ""
'''    End If
'''
'''    Call DeselectListBox(ObservationsList)
'''    Call ShowAcceptObs
'''    Observation.SetFocus
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox error$
'''End Sub
'''
'''Sub AddjustCurrentObs()
'''
''''****************************************************************************************
''''PCN     : PCN3576
''''Name    : AdjustCurrentObs
''''Created : 14 July 2005
''''Updated :
''''Prg By  : Antony van Iersel
''''Param   :
''''Desc    : Sets up the observations form to accept a new observations
''''Usage   :
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''
'''    Call ShowAcceptObs
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox error$
'''End Sub
'''
'''Sub AttentionTextBox(ItemToAttention As Variant)
'''
''''****************************************************************************************
''''PCN     : PCN3576
''''Name    : SetupNewObs
''''Created : 14 July 2005
''''Updated :
''''Prg By  : Antony van Iersel
''''Param   :
''''Desc    : Sets up the observations form to accept a new observations
''''Usage   :
''''****************************************************************************************
'''On Error GoTo Err_Handler
'''Dim Count As Integer
'''Dim CurrentColour As Long
'''
'''    CurrentColour = ItemToAttention.BackColor
'''
'''    ItemToAttention.SetFocus
'''    For Count = 0 To 3
'''        ItemToAttention.BackColor = &HFF: DoEvents: Call Sleep(200)
'''        ItemToAttention.BackColor = CurrentColour: DoEvents: Call Sleep(200)
'''    Next Count
'''    ItemToAttention.BackColor = CurrentColour:
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox error$
'''End Sub
'''
'''Function ExtractObservationSnapshotAtIndex(ByVal Index As Integer) As Boolean
'''On Error GoTo Err_Handler
'''
'''Dim SnapshotFilename As String
'''Dim PipeObsSnapShotFileOffset As Long 'PCN3576
'''Dim PipeObsSnapshotFileLength As Long 'PCN3576
'''
'''PipeObsSnapShotFileOffset = PipeObservations(Index).PipeObsSnapshotOffset
'''PipeObsSnapshotFileLength = PipeObservations(Index).PipeObsSnapshotLength
'''If PipeObsSnapShotFileOffset <> 0 And PipeObsSnapshotFileLength <> 0 Then
'''    SnapshotFilename = LocToSave & "Snapshot.bmp"
'''    Call PageFunctions.EmbeddedFileExtract(SnapshotFilename, PipeObsSnapShotFileOffset, PipeObsSnapshotFileLength)
'''    ClearLineScreen.LoadImage (LocToSave & "Snapshot.bmp")
'''    ExtractObservationSnapshotAtIndex = True
'''End If
'''
'''Exit Function
'''Err_Handler:
'''    MsgBox error$
'''End Function
'''
'''Sub GotoProfile(ByVal PVFrameNo As Long)
'''On Error GoTo Err_Handler
'''
'''
''''vvvv PCN2930 ***************************
'''If CLPScreenMode = PV And PVFrameNo <> 0 Then
'''    Call ClearLineScreen.GotoPVProfile(PVFrameNo, True)
'''End If
''''^^^^ ***********************************
'''
'''If mediatype = Video Then
'''    If CheckAVIInitialised = True Then 'Check that the AVI is correctly initialised before running the C code
'''        Call ClearLineScreen.GotoAVIFrame(PVDFileName, PVFrameNo, 1)
'''    End If
'''End If
'''
'''Exit Sub
'''Err_Handler:
'''    MsgBox error$
'''End Sub
'''
