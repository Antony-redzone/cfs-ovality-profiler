VERSION 5.00
Begin VB.Form PVGraphsKeyForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox DragBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00B36A36&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      MouseIcon       =   "PVGraphsKeyForm.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   105
      ScaleWidth      =   3705
      TabIndex        =   32
      Top             =   0
      Width           =   3735
   End
   Begin VB.PictureBox PVGraphsKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6195
      Left            =   0
      ScaleHeight     =   411
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   0
      Top             =   120
      Width           =   2205
      Begin VB.TextBox PVKey_Flat3D_Value0 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   360
         MousePointer    =   3  'I-Beam
         TabIndex        =   30
         Text            =   "-10%"
         Top             =   2100
         Width           =   615
      End
      Begin VB.TextBox PVKey_Flat3D_Value7 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   360
         MousePointer    =   3  'I-Beam
         TabIndex        =   29
         Text            =   "10%"
         Top             =   300
         Width           =   615
      End
      Begin VB.Image PVKey_Inclination_Icon 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   30
         Picture         =   "PVGraphsKeyForm.frx":08CA
         Stretch         =   -1  'True
         ToolTipText     =   "Delta Max"
         Top             =   5160
         Width           =   270
      End
      Begin VB.Label PVKey_Inclination_Value_Unit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1020
         TabIndex        =   45
         Top             =   5205
         Width           =   900
      End
      Begin VB.Label DiameterLabel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ø"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1980
         TabIndex        =   44
         Top             =   75
         Width           =   255
      End
      Begin VB.Label RadiusLabel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   570
         TabIndex        =   43
         Top             =   0
         Width           =   255
      End
      Begin VB.Label PVKey_Flat3D_Value3_Unit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   42
         Top             =   4920
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label PVKey_Flat3D_Value4_Unit 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   1890
         TabIndex        =   41
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label PVKey_Flat3D_Value6_Unit 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   1890
         TabIndex        =   40
         Top             =   600
         Width           =   840
      End
      Begin VB.Label PVKey_Flat3D_Value5_Unit 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   1890
         TabIndex        =   39
         Top             =   900
         Width           =   840
      End
      Begin VB.Label PVKey_Flat3D_Value7_Unit 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   1890
         TabIndex        =   38
         Top             =   300
         Width           =   840
      End
      Begin VB.Label PVKey_Flat3D_Value2_Unit 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   1890
         TabIndex        =   37
         Top             =   1500
         Width           =   840
      End
      Begin VB.Label PVKey_Flat3D_Value1_Unit 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   1890
         TabIndex        =   36
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label PVKey_Flat3D_Value0_Unit 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   1890
         TabIndex        =   35
         Top             =   2145
         Width           =   840
      End
      Begin VB.Image DimensionIcon 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   30
         Picture         =   "PVGraphsKeyForm.frx":0CFC
         Top             =   2760
         Width           =   270
      End
      Begin VB.Label DimensionValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1020
         TabIndex        =   34
         Top             =   2880
         Width           =   900
      End
      Begin VB.Label PVKey_FrameNo_Value 
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1410
         TabIndex        =   33
         Top             =   2445
         Width           =   570
      End
      Begin VB.Image PVKey_FrameNo_Icon 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   1050
         Picture         =   "PVGraphsKeyForm.frx":1286
         Top             =   2400
         Width           =   270
      End
      Begin VB.Label UnitSquare 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1920
         TabIndex        =   31
         Top             =   4200
         Width           =   255
      End
      Begin VB.Label PVKey_Ovality_Value_Unit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1020
         TabIndex        =   28
         Top             =   3240
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label PVKey_YDiameter_Value_Unit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   900
         TabIndex        =   27
         Top             =   3915
         Width           =   1020
      End
      Begin VB.Label PVKey_DeltaMax_Value_Unit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1020
         TabIndex        =   26
         Top             =   5490
         Width           =   900
      End
      Begin VB.Label PVKey_XDiameter_Value_Unit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   900
         TabIndex        =   25
         Top             =   3600
         Width           =   1020
      End
      Begin VB.Label PVKey_DeltaMin_Value_Unit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1020
         TabIndex        =   24
         Top             =   5775
         Width           =   900
      End
      Begin VB.Label PVKey_Capacity_Value_Unit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1020
         TabIndex        =   23
         Top             =   4260
         Width           =   900
      End
      Begin VB.Label PVKey_Flat3D_Value0_Unit 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   1020
         TabIndex        =   22
         Top             =   2145
         Width           =   840
      End
      Begin VB.Label PVKey_Flat3D_Value1_Unit 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   1020
         TabIndex        =   21
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label PVKey_Flat3D_Value2_Unit 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   1020
         TabIndex        =   20
         Top             =   1500
         Width           =   840
      End
      Begin VB.Label PVKey_Flat3D_Value7_Unit 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   1020
         TabIndex        =   19
         Top             =   300
         Width           =   840
      End
      Begin VB.Label PVKey_Flat3D_Value5_Unit 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   1020
         TabIndex        =   18
         Top             =   900
         Width           =   840
      End
      Begin VB.Label PVKey_Flat3D_Value6_Unit 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   1020
         TabIndex        =   17
         Top             =   600
         Width           =   840
      End
      Begin VB.Label PVKey_Flat3D_Value4_Unit 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   1020
         TabIndex        =   16
         Top             =   1200
         Width           =   840
      End
      Begin VB.Shape PVKey_Flat3D_Color7 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   270
         Left            =   30
         Top             =   300
         Width           =   270
      End
      Begin VB.Label PVKey_Capacity_Value 
         BackColor       =   &H00FFFFFF&
         Caption         =   "0%"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   4260
         Width           =   570
      End
      Begin VB.Label PVKey_Distance_Icon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   270
         Left            =   30
         TabIndex        =   14
         Top             =   2400
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label PVKey_Distance_Value 
         BackColor       =   &H00FFFFFF&
         Caption         =   "0m"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   2445
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Image PVKey_DeltaMax_Icon 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   30
         Picture         =   "PVGraphsKeyForm.frx":1568
         ToolTipText     =   "Delta Max"
         Top             =   5445
         Width           =   270
      End
      Begin VB.Label PVKey_YDiameter_Value 
         BackColor       =   &H00C0C0FF&
         Caption         =   "0%"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   3915
         Width           =   570
      End
      Begin VB.Image PVKey_YDiameter_Icon 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   30
         Picture         =   "PVGraphsKeyForm.frx":18AA
         ToolTipText     =   "Y"
         Top             =   3870
         Width           =   270
      End
      Begin VB.Label PVKey_DeltaMax_Value 
         BackColor       =   &H00FFFFFF&
         Caption         =   "0%"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   5505
         Width           =   570
      End
      Begin VB.Label PVKey_XDiameter_Value 
         BackColor       =   &H00FFFFC0&
         Caption         =   "0%"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   3600
         Width           =   570
      End
      Begin VB.Image PVKey_XDiameter_Icon 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   30
         Picture         =   "PVGraphsKeyForm.frx":1BEC
         ToolTipText     =   "X"
         Top             =   3555
         Width           =   270
      End
      Begin VB.Label PVKey_DeltaMin_Value 
         BackColor       =   &H00FFFFFF&
         Caption         =   "0%"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   5775
         Width           =   570
      End
      Begin VB.Image PVKey_DeltaMin_Icon 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   30
         Picture         =   "PVGraphsKeyForm.frx":1F2E
         ToolTipText     =   "Delta Min"
         Top             =   5745
         Width           =   270
      End
      Begin VB.Label PVKey_Ovality_Value 
         BackColor       =   &H00FFFFFF&
         Caption         =   "0%"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   3285
         Width           =   570
      End
      Begin VB.Image PVKey_Ovality_Icon 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   30
         Picture         =   "PVGraphsKeyForm.frx":2270
         Top             =   3240
         Width           =   270
      End
      Begin VB.Image PVKey_Flat3D_Color4 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   30
         Top             =   1200
         Width           =   270
      End
      Begin VB.Label PVKey_Flat3D_Value4 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   570
      End
      Begin VB.Label PVKey_Flat3D_Value6 
         BackStyle       =   0  'Transparent
         Caption         =   "+6.6%"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   570
      End
      Begin VB.Label PVKey_Flat3D_Value5 
         BackStyle       =   0  'Transparent
         Caption         =   "+3.3%"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   900
         Width           =   570
      End
      Begin VB.Image PVKey_Flat3D_Color3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   30
         Top             =   4905
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image PVKey_Capacity_Icon 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   30
         Picture         =   "PVGraphsKeyForm.frx":25B2
         Top             =   4230
         Width           =   270
      End
      Begin VB.Label PVKey_Flat3D_Value1 
         BackStyle       =   0  'Transparent
         Caption         =   "-6.6%"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1800
         Width           =   570
      End
      Begin VB.Label PVKey_Flat3D_Value2 
         BackStyle       =   0  'Transparent
         Caption         =   "-3.3%"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1500
         Width           =   570
      End
      Begin VB.Label PVKey_Flat3D_Value3 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   4950
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Shape PVKey_Flat3D_Color6 
         FillColor       =   &H000096FF&
         FillStyle       =   0  'Solid
         Height          =   270
         Left            =   30
         Top             =   600
         Width           =   270
      End
      Begin VB.Shape PVKey_Flat3D_Color5 
         FillColor       =   &H0014FFFF&
         FillStyle       =   0  'Solid
         Height          =   270
         Left            =   30
         Top             =   900
         Width           =   270
      End
      Begin VB.Shape PVKey_Flat3D_Color0 
         FillColor       =   &H006F4928&
         FillStyle       =   0  'Solid
         Height          =   270
         Left            =   30
         Top             =   2100
         Width           =   270
      End
      Begin VB.Shape PVKey_Flat3D_Color1 
         FillColor       =   &H00CC9B5A&
         FillStyle       =   0  'Solid
         Height          =   270
         Left            =   30
         Top             =   1800
         Width           =   270
      End
      Begin VB.Shape PVKey_Flat3D_Color2 
         FillColor       =   &H00EEE0B5&
         FillStyle       =   0  'Solid
         Height          =   270
         Left            =   30
         Top             =   1500
         Width           =   270
      End
      Begin VB.Label PVKey_Flat3D_Value3_Unit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0mm"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   870
         TabIndex        =   1
         Top             =   4950
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Shape PVGraphsKey_Shade 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  'Opaque
         Height          =   270
         Index           =   7
         Left            =   3240
         Top             =   1755
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.Shape PVGraphsKey_Shade 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         Height          =   270
         Index           =   6
         Left            =   3240
         Top             =   1515
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.Shape PVGraphsKey_Shade 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  'Opaque
         Height          =   270
         Index           =   5
         Left            =   3240
         Top             =   1260
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.Shape PVGraphsKey_Shade 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         Height          =   270
         Index           =   4
         Left            =   3240
         Top             =   1020
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.Shape PVGraphsKey_Shade 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  'Opaque
         Height          =   270
         Index           =   3
         Left            =   3240
         Top             =   765
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.Shape PVGraphsKey_Shade 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         Height          =   270
         Index           =   2
         Left            =   3240
         Top             =   525
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.Shape PVGraphsKey_Shade 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  'Opaque
         Height          =   270
         Index           =   1
         Left            =   3240
         Top             =   270
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.Shape PVGraphsKey_Shade 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         Height          =   270
         Index           =   0
         Left            =   3240
         Top             =   30
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.Label PVKey_Inclination_Value 
         BackColor       =   &H00FFFFFF&
         Caption         =   "0%"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   5220
         Width           =   930
      End
   End
End
Attribute VB_Name = "PVGraphsKeyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MouseDownX As Integer
Private MouseDownY As Integer

Dim LastMouseDownX As Single
Dim LastMouseDownY As Single
Dim Action As String
Dim PVGraphKeyMode As String 'PCN4171

'vvvv PCN4328 ************************************
Public Event MouseLeave()
'Public Event MouseHover()
'^^^^ ********************************************

Dim WithEvents MouseTrackDragBar As clsTrackInfo
Attribute MouseTrackDragBar.VB_VarHelpID = -1


Public Function DisplayPVGraphsKey()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : DisplayPVGraphsKey
'Created : 20 May 2004, PCN2818
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Displays the PVGraphs Key in the PVScreen.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim HideObjTopOffset As Integer
Dim ColourVal(7) As Double
Dim Position1 As Double
Dim Position2 As Double
Dim Position3 As Double
Dim Position4 As Double
Dim Position5 As Double
Dim Position6 As Double
Dim Position7 As Double
Dim Position8 As Double

'Dim Units As String 'PCN4250
Dim FormHeight As Integer
Dim ColourFormHeight
Dim I As Integer


ColourFormHeight = 2700 + Me.DragBar.height 'PCN4920 was 2350
FormHeight = 2115 + Me.DragBar.height 'PCN4185

For I = 0 To 5
    PVGraphsKey_Shade(I).Left = 0
Next I

If Flat3dLimitL = 0 Then Flat3dLimitL = -10
If Flat3dLimitR = 0 Then Flat3dLimitR = 10

Position1 = PVKey_Flat3D_Color7.Top - (18)
Position2 = PVKey_Flat3D_Color6.Top - (18)
Position3 = PVKey_Flat3D_Color5.Top - (18)
Position4 = PVKey_Flat3D_Color4.Top - (18)
Position5 = PVKey_Flat3D_Color2.Top - (18)
Position6 = PVKey_Flat3D_Color1.Top - (18)
Position7 = PVKey_Flat3D_Color0.Top - (18)



PVGraphsKeyForm.width = 2685 'PCN4920 added one more column for diameter measurements, was 2000
Me.DragBar.width = PVGraphsKeyForm.width 'PCN4185


PVGraphsKey.width = 2685 'PCN4920 added one more column for diameter measurements, was 2000


DimensionIcon.Top = Position1: DimensionValue.Top = Position1 + 4
PVKey_Ovality_Icon.Top = Position2: PVKey_Ovality_Value.Top = Position2 + 4: PVKey_Ovality_Value_Unit.Top = Position2 + 4

''PVKey_DeltaMin_Icon.Top = Position4: PVKey_DeltaMin_Value.Top = Position4 + 4: PVKey_DeltaMin_Value_Unit.Top = Position4 + 4
PVKey_XDiameter_Icon.Top = Position3: PVKey_XDiameter_Value.Top = Position3 + 4: PVKey_XDiameter_Value_Unit.Top = Position3 + 4
PVKey_YDiameter_Icon.Top = Position4: PVKey_YDiameter_Value.Top = Position4 + 4: PVKey_YDiameter_Value_Unit.Top = Position4 + 4
PVKey_Capacity_Icon.Top = Position5: PVKey_Capacity_Value.Top = Position5 + 4: PVKey_Capacity_Value_Unit.Top = Position5 + 4
'PCN6458 PVKey_Inclination_Icon.Top = Position6: PVKey_Inclination_Value.Top = Position6 + 4: PVKey_Inclination_Value_Unit.Top = Position6 + 4

UnitSquare.Top = Position5 - 2
PVKey_Distance_Icon.Top = (FormHeight / 15) - 26 - 4: PVKey_Distance_Value.Top = (FormHeight / 15) - 26  'PCN4920
PVKey_FrameNo_Icon.Top = (FormHeight / 15) - 26 - 2: PVKey_FrameNo_Value.Top = (FormHeight / 15) - 26 'PCN4171 'PCN4920

'vvvv PCN4171 ****************************************
If ConfigInfo.DistanceStart >= 0 And DistanceMethod <> "None" Then
    PVKey_Distance_Icon.Visible = True
    PVKey_Distance_Value.Visible = True
    If MeasurementUnits = "mm" Then
        PVKey_Distance_Icon.Caption = "m"
    Else
        PVKey_Distance_Icon.Caption = "ft"
    End If
End If
'^^^^ ************************************************


If ImageGraphState(0).GraphType = "Flat" And PVGraphKeyMode <> "PVGraphValuesKey" Then 'PCN4171
    
''    If MeasurementUnits = "mm" Then 'PCN4250
''        Units = "mm"
''    Else
''        Units = "in"
''    End If
    Me.DiameterLabel.Visible = True
    Me.RadiusLabel.Visible = True
    
    PVGraphsKey.height = ColourFormHeight - Me.DragBar.height: PVGraphsKeyForm.height = ColourFormHeight
    
    DimensionIcon.Visible = False: DimensionValue.Visible = False 'PCN4171
    
    PVKey_Capacity_Icon.Visible = False: PVKey_Capacity_Value.Visible = False: PVKey_Capacity_Value_Unit.Visible = False
    PVKey_Ovality_Icon.Visible = False: PVKey_Ovality_Value.Visible = False: PVKey_Ovality_Value_Unit.Visible = False
'PCN6458     PVKey_Inclination_Icon.Visible = False: PVKey_Inclination_Value.Visible = False: PVKey_Inclination_Value_Unit.Visible = False
    PVKey_DeltaMin_Icon.Visible = False: PVKey_DeltaMin_Value.Visible = False: PVKey_DeltaMin_Value_Unit.Visible = False
    PVKey_XDiameter_Icon.Visible = False: PVKey_XDiameter_Value.Visible = False: PVKey_XDiameter_Value_Unit.Visible = False
    PVKey_YDiameter_Icon.Visible = False: PVKey_YDiameter_Value.Visible = False: PVKey_YDiameter_Value_Unit.Visible = False
    UnitSquare.Visible = False
    'vvvv PCN4171 ****************************************
'    PVKey_Distance_Icon.Visible = False: PVKey_Distance_Value.Visible = False 'PCN4171
    'Ensure the distance is available for view at all times
    'PVKey_Distance_Icon.Top = PVKey_Flat3D_Color0.Top + 19: PVKey_Distance_Value.Top = PVKey_Flat3D_Color0.Top + 19 + 4 'PCN4920
    'PVKey_FrameNo_Icon.Top = PVKey_Flat3D_Color0.Top + 19: PVKey_FrameNo_Value.Top = PVKey_Flat3D_Color0.Top + 19 + 4   'PCN4920
    
    PVKey_Distance_Icon.Top = (ColourFormHeight / 15) - 26 - 4: PVKey_Distance_Value.Top = (ColourFormHeight / 15) - 26  'PCN4920
    PVKey_FrameNo_Icon.Top = (ColourFormHeight / 15) - 26 - 2: PVKey_FrameNo_Value.Top = (ColourFormHeight / 15) - 26 'PCN4171 'PCN4920
    
    PVKey_Distance_Value.Visible = True
    PVKey_Distance_Icon.Visible = True
    '^^^^ ************************************************
        
    PVKey_Flat3D_Color0.Visible = True: PVKey_Flat3D_Value0.Visible = True: PVKey_Flat3D_Value0_Unit(0).Visible = True 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Color1.Visible = True: PVKey_Flat3D_Value1.Visible = True: PVKey_Flat3D_Value1_Unit(0).Visible = True 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Color2.Visible = True: PVKey_Flat3D_Value2.Visible = True: PVKey_Flat3D_Value2_Unit(0).Visible = True 'PCN4920 (0) is percentage measurements
  '  PVKey_Flat3D_Color3.Visible = True: PVKey_Flat3D_Value3.Visible = True: PVKey_Flat3D_Value3_Unit.Visible = True
    PVKey_Flat3D_Color4.Visible = True: PVKey_Flat3D_Value4.Visible = True: PVKey_Flat3D_Value4_Unit(0).Visible = True 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Color5.Visible = True: PVKey_Flat3D_Value5.Visible = True: PVKey_Flat3D_Value5_Unit(0).Visible = True 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Color6.Visible = True: PVKey_Flat3D_Value6.Visible = True: PVKey_Flat3D_Value6_Unit(0).Visible = True 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Color7.Visible = True: PVKey_Flat3D_Value7.Visible = True: PVKey_Flat3D_Value7_Unit(0).Visible = True 'PCN4920 (0) is percentage measurements
    
    'PCN4920 this is all new
    PVKey_Flat3D_Color0.Visible = True: PVKey_Flat3D_Value0.Visible = True: PVKey_Flat3D_Value0_Unit(1).Visible = True 'PCN4920 (1) is diameter percentage
    PVKey_Flat3D_Color1.Visible = True: PVKey_Flat3D_Value1.Visible = True: PVKey_Flat3D_Value1_Unit(1).Visible = True 'PCN4920 (1) is diameter percentage
    PVKey_Flat3D_Color2.Visible = True: PVKey_Flat3D_Value2.Visible = True: PVKey_Flat3D_Value2_Unit(1).Visible = True 'PCN4920 (1) is diameter percentage
  '  PVKey_Flat3D_Color3.Visible = True: PVKey_Flat3D_Value3.Visible = True: PVKey_Flat3D_Value3_Unit.Visible = True
    PVKey_Flat3D_Color4.Visible = True: PVKey_Flat3D_Value4.Visible = True: PVKey_Flat3D_Value4_Unit(1).Visible = True 'PCN4920 (1) is diameter percentage
    PVKey_Flat3D_Color5.Visible = True: PVKey_Flat3D_Value5.Visible = True: PVKey_Flat3D_Value5_Unit(1).Visible = True 'PCN4920 (1) is diameter percentage
    PVKey_Flat3D_Color6.Visible = True: PVKey_Flat3D_Value6.Visible = True: PVKey_Flat3D_Value6_Unit(1).Visible = True 'PCN4920 (1) is diameter percentage
    PVKey_Flat3D_Color7.Visible = True: PVKey_Flat3D_Value7.Visible = True: PVKey_Flat3D_Value7_Unit(1).Visible = True 'PCN4920 (1) is diameter percentage
    
    
    'Update Flat3D percentage values
    Dim DeltaLimitPercent As Single
'    Dim DeltaLimitPerL As Single
'    Dim DeltaLimitPerR As Single
    Call PrecisionVisionGraph.SetLimitLines
'    Call PrecisionVisionGraph.GetPVXLimits_Delta(DeltaLimitPerL, DeltaLimitPerR) 'PCN2680
    
    Call ThreeDColourValues(ColourVal)
    
    Call FlatColourRealValues(ColourVal) 'PCN4250
    
     PVKey_Flat3D_Value0_Unit(1).Left = 126
     PVKey_Flat3D_Value1_Unit(1).Left = 126
     PVKey_Flat3D_Value2_Unit(1).Left = 126
     PVKey_Flat3D_Value4_Unit(1).Left = 126
     PVKey_Flat3D_Value5_Unit(1).Left = 126
     PVKey_Flat3D_Value6_Unit(1).Left = 126
     PVKey_Flat3D_Value7_Unit(1).Left = 126
     Me.DiameterLabel.Left = 132
     Me.PVKey_FrameNo_Icon.Left = 102
     Me.PVKey_FrameNo_Value.Left = 126
    
    If MedianFlat And PVDFileName <> "" Then 'PCN4974
       PVKey_Flat3D_Value0_Unit(0).Visible = False 'PCN4920 (0) is percentage measurements
       PVKey_Flat3D_Value1_Unit(0).Visible = False 'PCN4920 (0) is percentage measurements
       PVKey_Flat3D_Value2_Unit(0).Visible = False 'PCN4920 (0) is percentage measurements
    
       PVKey_Flat3D_Value4_Unit(0).Visible = False 'PCN4920 (0) is percentage measurements
       PVKey_Flat3D_Value5_Unit(0).Visible = False 'PCN4920 (0) is percentage measurements
       PVKey_Flat3D_Value6_Unit(0).Visible = False 'PCN4920 (0) is percentage measurements
       PVKey_Flat3D_Value7_Unit(0).Visible = False 'PCN4920 (0) is percentage measurements
       Me.width = Me.width - (36 * 15)
       Me.PVGraphsKey.width = Me.PVGraphsKey.width - (36 * 15)

       
        Me.PVKey_Flat3D_Value0_Unit(1).Left = 90
        Me.PVKey_Flat3D_Value1_Unit(1).Left = 90
        Me.PVKey_Flat3D_Value2_Unit(1).Left = 90
        Me.PVKey_Flat3D_Value4_Unit(1).Left = 90
        Me.PVKey_Flat3D_Value5_Unit(1).Left = 90
        Me.PVKey_Flat3D_Value6_Unit(1).Left = 90
        Me.PVKey_Flat3D_Value7_Unit(1).Left = 90
        Me.DiameterLabel.Left = 96 '132
        Me.PVKey_FrameNo_Icon.Left = 76 '102
        Me.PVKey_FrameNo_Value.Left = 100 '126
    End If
    
Else
    Me.DiameterLabel.Left = 96 '132
    Me.PVKey_FrameNo_Icon.Left = 76 '102
    Me.PVKey_FrameNo_Value.Left = 100 '126


    Me.DiameterLabel.Visible = False
    Me.RadiusLabel.Visible = False
    PVGraphsKey.width = 2205: Me.width = 2205
    PVGraphsKey.height = FormHeight - Me.DragBar.height: PVGraphsKeyForm.height = FormHeight
    DimensionIcon.Visible = True: DimensionValue.Visible = True 'PCN4171
    PVKey_Capacity_Icon.Visible = True: PVKey_Capacity_Value.Visible = True: PVKey_Capacity_Value_Unit.Visible = True
    PVKey_Ovality_Icon.Visible = True: PVKey_Ovality_Value.Visible = True: 'PVKey_Ovality_Value_Unit.Visible = True
'PCN6458     PVKey_Inclination_Icon.Visible = True: PVKey_Inclination_Value.Visible = True: PVKey_Inclination_Value_Unit.Visible = True 'PCN4171
'    PVKey_DeltaMin_Icon.Visible = True: PVKey_DeltaMin_Value.Visible = True: PVKey_DeltaMin_Value_Unit.Visible = True
    PVKey_XDiameter_Icon.Visible = True: PVKey_XDiameter_Value.Visible = True: PVKey_XDiameter_Value_Unit.Visible = True
    PVKey_YDiameter_Icon.Visible = True: PVKey_YDiameter_Value.Visible = True: PVKey_YDiameter_Value_Unit.Visible = True
    PVKey_Distance_Icon.Visible = True: PVKey_Distance_Value.Visible = True
    UnitSquare.Visible = True
        
    PVKey_Flat3D_Color0.Visible = False: PVKey_Flat3D_Value0.Visible = False: PVKey_Flat3D_Value0_Unit(0).Visible = False 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Color1.Visible = False: PVKey_Flat3D_Value1.Visible = False: PVKey_Flat3D_Value1_Unit(0).Visible = False 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Color2.Visible = False: PVKey_Flat3D_Value2.Visible = False: PVKey_Flat3D_Value2_Unit(0).Visible = False 'PCN4920 (0) is percentage measurements
  '  PVKey_Flat3D_Color3.Visible = False: PVKey_Flat3D_Value3.Visible = False: PVKey_Flat3D_Value3_Unit.Visible = False
    PVKey_Flat3D_Color4.Visible = False: PVKey_Flat3D_Value4.Visible = False: PVKey_Flat3D_Value4_Unit(0).Visible = False 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Color5.Visible = False: PVKey_Flat3D_Value5.Visible = False: PVKey_Flat3D_Value5_Unit(0).Visible = False 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Color6.Visible = False: PVKey_Flat3D_Value6.Visible = False: PVKey_Flat3D_Value6_Unit(0).Visible = False 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Color7.Visible = False: PVKey_Flat3D_Value7.Visible = False: PVKey_Flat3D_Value7_Unit(0).Visible = False 'PCN4920 (0) is percentage measurements
    
    'PCN4920 this is all new
    PVKey_Flat3D_Color0.Visible = False: PVKey_Flat3D_Value0.Visible = False: PVKey_Flat3D_Value0_Unit(1).Visible = False 'PCN4920 (1) percentage of diameter
    PVKey_Flat3D_Color1.Visible = False: PVKey_Flat3D_Value1.Visible = False: PVKey_Flat3D_Value1_Unit(1).Visible = False 'PCN4920 (1) percentage of diameter
    PVKey_Flat3D_Color2.Visible = False: PVKey_Flat3D_Value2.Visible = False: PVKey_Flat3D_Value2_Unit(1).Visible = False 'PCN4920 (1) percentage of diameter
  '  PVKey_Flat3D_Color3.Visible = False: PVKey_Flat3D_Value3.Visible = False: PVKey_Flat3D_Value3_Unit.Visible = False
    PVKey_Flat3D_Color4.Visible = False: PVKey_Flat3D_Value4.Visible = False: PVKey_Flat3D_Value4_Unit(1).Visible = False 'PCN4920 (1) percentage of diameter
    PVKey_Flat3D_Color5.Visible = False: PVKey_Flat3D_Value5.Visible = False: PVKey_Flat3D_Value5_Unit(1).Visible = False 'PCN4920 (1) percentage of diameter
    PVKey_Flat3D_Color6.Visible = False: PVKey_Flat3D_Value6.Visible = False: PVKey_Flat3D_Value6_Unit(1).Visible = False 'PCN4920 (1) percentage of diameter
    PVKey_Flat3D_Color7.Visible = False: PVKey_Flat3D_Value7.Visible = False: PVKey_Flat3D_Value7_Unit(1).Visible = False 'PCN4920 (1) percentage of diameter
    
    'PCN5186
    If MedianFlat And PVDFileName <> "" Then
        PVKey_XDiameter_Value_Unit.Visible = False
        PVKey_YDiameter_Value_Unit.Visible = False
    End If
    
    'vvvv PCN4171 ****************************************
'    If ConfigInfo.DistanceStart >= 0 And DistanceMethod <> "None" Then
'        PVKey_Distance_Icon.Visible = True
'        PVKey_Distance_Value.Visible = True
'        If MeasurementUnits = "mm" Then
'            PVKey_Distance_Icon.Caption = "m"
'        Else
'            PVKey_Distance_Icon.Caption = "ft"
'        End If
'    End If
    '^^^^ ************************************************
End If

PVGraphsKey.Visible = True
PVGraphsKeyForm.Visible = True

'PVGraphsKeyForm.ZOrder 0
Call ScreenDrawing.FormTopMost(PVGraphsKeyForm.hwnd) 'PCN2990

    
Exit Function
Err_Handler:
    MsgBox Err & "-GKF1:" & Error$
    

End Function

Private Sub DragBar_DblClick()
On Error GoTo Err_Handler

If PVGraphKeyMode = "PVGraphValuesKey" Then
    PVGraphKeyMode = "KeyNormal"
Else
    PVGraphKeyMode = "PVGraphValuesKey"
End If
Call DisplayPVGraphsKey
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-GKF2:" & Error$
    End Select
End Sub

Private Sub DragBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
Dim curSelect As StdPicture

Set curSelect = LoadResPicture(109, vbResIcon)
DragBar.MouseIcon = curSelect
LastMouseDownX = X
LastMouseDownY = Y
If Button = 1 Then Action = "Move"
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-GKF3:" & Error$
    End Select
End Sub

Private Sub DragBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.DragBar.BackColor = &HC0C000 'PCN4328

If Action = "Move" Then
    Me.Left = Me.Left + X - LastMouseDownX
    Me.Top = Me.Top + Y - LastMouseDownY
End If
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-GKF4:" & Error$
    End Select
End Sub

Private Sub DragBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
Dim curSelect As StdPicture

Set curSelect = LoadResPicture(108, vbResIcon)
DragBar.MouseIcon = curSelect
Action = ""
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-GKF5:" & Error$
    End Select
End Sub







Private Sub Form_Load()
'****************************************************************************************
'Name    : Form_Load
'Created : Dec 05 2005, PCN3931
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim RetVal As Variant

RetVal = SetWindowPos(Me.hwnd, -1, 100, 100, 519, 433, &H40)

'vvvv PCN4328 ************************************
'Initilise the mouse leave event on the key's drag bar.
Set MouseTrackDragBar = New clsTrackInfo
MouseTrackDragBar.hwnd = DragBar.hwnd

StartTrack MouseTrackDragBar
'^^^^ ********************************************

'vvvv PCN4341 *********************
PVKey_Ovality_Icon.ToolTipText = DisplayMessage("Ovality")
DimensionIcon.ToolTipText = DisplayMessage("Median Diameter")
PVKey_XDiameter_Icon.ToolTipText = DisplayMessage("X Diameter")
PVKey_YDiameter_Icon.ToolTipText = DisplayMessage("Y Diameter")
PVKey_Capacity_Icon.ToolTipText = DisplayMessage("Capacity")

PVKey_FrameNo_Icon.ToolTipText = DisplayMessage("Frame No")
'^^^^ *****************************

'vvvv PCN4341 *********************
PVKey_Flat3D_Value7_Unit(0).ToolTipText = DisplayMessage("Variation in radius") 'PCN4920 (0) is measurement percentage
PVKey_Flat3D_Value6_Unit(0).ToolTipText = DisplayMessage("Variation in radius") 'PCN4920 (0) is measurement percentage
PVKey_Flat3D_Value5_Unit(0).ToolTipText = DisplayMessage("Variation in radius") 'PCN4920 (0) is measurement percentage
PVKey_Flat3D_Value4_Unit(0).ToolTipText = DisplayMessage("Variation in radius") 'PCN4920 (0) is measurement percentage
PVKey_Flat3D_Value3_Unit(0).ToolTipText = DisplayMessage("Variation in radius") 'PCN4920 (0) is measurement percentage
PVKey_Flat3D_Value2_Unit(0).ToolTipText = DisplayMessage("Variation in radius") 'PCN4920 (0) is measurement percentage
PVKey_Flat3D_Value1_Unit(0).ToolTipText = DisplayMessage("Variation in radius") 'PCN4920 (0) is measurement percentage
PVKey_Flat3D_Value0_Unit(0).ToolTipText = DisplayMessage("Variation in radius") 'PCN4920 (0) is measurement percentage

PVKey_Flat3D_Value7_Unit(1).ToolTipText = DisplayMessage("Percentage of diameter") 'PCN4920 (1) percentage of diameter
PVKey_Flat3D_Value6_Unit(1).ToolTipText = DisplayMessage("Percentage of diameter") 'PCN4920 (1) percentage of diameter
PVKey_Flat3D_Value5_Unit(1).ToolTipText = DisplayMessage("Percentage of diameter") 'PCN4920 (1) percentage of diameter
PVKey_Flat3D_Value4_Unit(1).ToolTipText = DisplayMessage("Percentage of diameter") 'PCN4920 (1) percentage of diameter
PVKey_Flat3D_Value3_Unit(1).ToolTipText = DisplayMessage("Percentage of diameter") 'PCN4920 (1) percentage of diameter
PVKey_Flat3D_Value2_Unit(1).ToolTipText = DisplayMessage("Percentage of diameter") 'PCN4920 (1) percentage of diameter
PVKey_Flat3D_Value1_Unit(1).ToolTipText = DisplayMessage("Percentage of diameter") 'PCN4920 (1) percentage of diameter
PVKey_Flat3D_Value0_Unit(1).ToolTipText = DisplayMessage("Percentage of diameter") 'PCN4920 (1) percentage of diameter

PVKey_Flat3D_Value7.ToolTipText = DisplayMessage("Percentage of radius") ' PCN4920 percentage of radius
PVKey_Flat3D_Value6.ToolTipText = DisplayMessage("Percentage of radius") ' PCN4920 percentage of radius
PVKey_Flat3D_Value5.ToolTipText = DisplayMessage("Percentage of radius") ' PCN4920 percentage of radius
PVKey_Flat3D_Value4.ToolTipText = DisplayMessage("Percentage of radius") ' PCN4920 percentage of radius
PVKey_Flat3D_Value3.ToolTipText = DisplayMessage("Percentage of radius") ' PCN4920 percentage of radius
PVKey_Flat3D_Value2.ToolTipText = DisplayMessage("Percentage of radius") ' PCN4920 percentage of radius
PVKey_Flat3D_Value1.ToolTipText = DisplayMessage("Percentage of radius") ' PCN4920 percentage of radius
PVKey_Flat3D_Value0.ToolTipText = DisplayMessage("Percentage of radius") ' PCN4920 percentage of radius
'^^^^ *****************************

Exit Sub
Err_Handler:
    MsgBox Err & "-GKF6:" & Error$
End Sub

Private Sub Form_Terminate() 'PCN4328
On Error GoTo Err_Handler

EndTrack MouseTrackDragBar
Set MouseTrackDragBar = Nothing

Exit Sub
Err_Handler:
    MsgBox Err & "-GKF7:" & Error$
End Sub

Private Sub MouseTrackDragBar_MouseLeave() 'PCN4328
On Error GoTo Err_Handler

Me.DragBar.BackColor = &HB36A36

RaiseEvent MouseLeave
Exit Sub
Err_Handler:
    MsgBox Err & "-GKF8:" & Error$
End Sub

Private Sub PVKey_Flat3D_Value0_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handler
    Select Case KeyAscii
        Case 13: Call PVKey_Flat3D_Value0_LostFocus 'RETURN key
        Case 27: Call PVKey_Flat3D_Value0_LostFocus 'ESC key
    End Select
Exit Sub
Err_Handler:
    MsgBox Err & "-GKF9:" & Error$
End Sub

Private Sub PVKey_Flat3D_Value7_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handler
    Select Case KeyAscii
        Case 13: Call PVKey_Flat3D_Value7_LostFocus 'RETURN key
        Case 27: Call PVKey_Flat3D_Value7_LostFocus 'ESC key
    End Select
Exit Sub
Err_Handler:
    MsgBox Err & "-GKF10:" & Error$
End Sub








Private Sub PVGraphsKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVGraphsKey_MouseDown
'Created : 27 May 2004, PCN2818
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Enables movement of the PVGraphsKey
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

PVGraphsKey.MousePointer = 99
PVGraphsKey.MouseIcon = LoadResPicture(109, vbResIcon) 'Move holding icon
MouseDownY = Y
MouseDownX = X

Exit Sub
Err_Handler:
    MsgBox Err & "-GKF11:" & Error$
End Sub

Private Sub PVGraphsKey_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVGraphsKey_MouseMove
'Created : 27 May 2004, PCN2818
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Enables movement of the PVGraphsKey
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If Button = 1 Then
    PVGraphsKeyForm.Top = PVGraphsKeyForm.Top - MouseDownY + Y
    PVGraphsKeyForm.Left = PVGraphsKeyForm.Left - MouseDownX + X
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-GKF12:" & Error$
End Sub

Private Sub PVGraphsKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVGraphsKey_MouseUp
'Created : 27 May 2004, PCN2818
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Enables movement of the PVGraphsKey
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

PVGraphsKey.MousePointer = 99
PVGraphsKey.MouseIcon = LoadResPicture(108, vbResIcon) 'Move icon

Exit Sub
Err_Handler:
    MsgBox Err & "-GKF13:" & Error$
End Sub

Function PVGraphsKeyUpdate()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVGraphsKeyUpdate
'Created : 26 May 2003, PCN2818
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Updates the PVGraphs Key in the PVScreen.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Dim CapacityVal As Double
Dim OvalityVal As Double
Dim OvalityOrigVal As Double
Dim DeltaMinVal As Double
Dim DeltaMaxVal As Double
Dim XDiameterVal As Double
Dim YDiameterVal As Double
Dim DisplayUnits As String
Dim DiameterMedianVal As Double
Dim DiameterMaxVal As Double
Dim DiameterMinVal As Double
'PCN6458 Dim InclinationVal As Double
'PCN6458 Dim InclinationValDeflection As Double




'If CLPScreenMode = PV Then
If PVDataNoOfLines > 1 And Not PVRecording Then
    If PVGraphsKeyForm.height > 5000 Then Call DisplayPVGraphsKey
    If MeasurementUnits = "mm" Then
        DisplayUnits = "mm"
    Else
        DisplayUnits = "in"
    End If

    PVKey_FrameNo_Value.Caption = PVFrameNo 'PCN4171
    
    If DebrisOn = False Then
        If UBound(GraphInfoContainer(PVCapacitySmooth).DataSingle) = 0 Then
            CapacityVal = PVCapacityFullData(PVFrameNo) + CapacityDataOffset ' / PVCalculationsMultiplier
        Else
            CapacityVal = GraphInfoContainer(PVCapacitySmooth).DataSingle(PVFrameNo) + CapacityDataOffset
        End If
    
    
    
    Else
        CapacityVal = GraphInfoContainer(PVDebris).DataSingle(PVFrameNo) 'PCN4461
    End If
        
        
    If UBound(GraphInfoContainer(PVOvalitySmooth).DataSingle) = 0 Then
        OvalityVal = Abs(GraphInfoContainer(PVOvality).DataSingle(PVFrameNo)) 'PCN3540 / PVCalculationsMultiplier
    Else
        OvalityVal = Abs(GraphInfoContainer(PVOvalitySmooth).DataSingle(PVFrameNo))
    End If

    
    
    'OvalityOrigVal = PVOvalityOrigFullData(PVFrameNo)
    
'    DeltaMinVal = PVDeltaFullMin(PVFrameNo)
'    DeltaMaxVal = PVDeltaFullMax(PVFrameNo)

    'PCN5186 added the following three lines
    If MedianFlat And PVDFileName <> "" Then
        XDiameterVal = GraphInfoContainer(PVDeflectionX).DataSingle(PVFrameNo)
        YDiameterVal = GraphInfoContainer(PVDeflectionY).DataSingle(PVFrameNo)
    
    ElseIf UBound(GraphInfoContainer(PVXDiameterSmooth).DataSingle) = 0 Then 'PCN9999
        XDiameterVal = PVXDiameterFullData(PVFrameNo)
        YDiameterVal = PVYDiameterFullData(PVFrameNo)
    Else
        XDiameterVal = GraphInfoContainer(PVXDiameterSmooth).DataSingle(PVFrameNo)
        YDiameterVal = GraphInfoContainer(PVYDiameterSmooth).DataSingle(PVFrameNo)
    End If
        
    If UBound(GraphInfoContainer(PVMedianDiameterSmooth).DataSingle) = 0 Then
        DiameterMedianVal = PVDiameterMedian(PVFrameNo) + TrueDiameterOffset
    Else
        DiameterMedianVal = GraphInfoContainer(PVMedianDiameterSmooth).DataSingle(PVFrameNo) + TrueDiameterOffset
    End If
    
    If UBound(GraphInfoContainer(PVMaxDiameterSmooth).DataSingle) = 0 Then
        DiameterMaxVal = GraphInfoContainer(PVMaxDiameter).DataDouble(PVFrameNo)
    Else
        DiameterMaxVal = GraphInfoContainer(PVMaxDiameterSmooth).DataSingle(PVFrameNo)
    End If
    
    If UBound(GraphInfoContainer(PVMinDiameterSmooth).DataSingle) = 0 Then
        DiameterMinVal = GraphInfoContainer(PVMinDiameter).DataDouble(PVFrameNo) 'PCN4333
    Else
        DiameterMinVal = GraphInfoContainer(PVMinDiameterSmooth).DataSingle(PVFrameNo) 'PCN4333
    End If
    
    'PCN6128
'PCN6458    If MeasurementUnits = "mm" Then
'PCN6458        If UBound(GraphInfoContainer(PVInclinationSmooth).DataSingle) = 0 Then
'PCN6458            InclinationValDeflection = (GraphInfoContainer(PVInclination).DataSingle(PVFrameNo) - GraphInfoContainer(PVDesignGradient).DataSingle(PVFrameNo)) * 1000
'PCN6458            InclinationVal = (GraphInfoContainer(PVInclination).DataSingle(PVFrameNo)) '- GraphInfoContainer(PVDesignGradient).DataSingle(PVFrameNo)) * 1000
'PCN6458        Else
'PCN6458            InclinationValDeflection = (GraphInfoContainer(PVInclinationSmooth).DataSingle(PVFrameNo) - GraphInfoContainer(PVDesignGradient).DataSingle(PVFrameNo)) * 1000
'PCN6458            InclinationVal = (GraphInfoContainer(PVInclinationSmooth).DataSingle(PVFrameNo)) '- GraphInfoContainer(PVDesignGradient).DataSingle(PVFrameNo)) * 1000
'PCN6458        End If
'PCN6458    Else
'PCN6458        If UBound(GraphInfoContainer(PVInclinationSmooth).DataSingle) = 0 Then
'PCN6458            InclinationValDeflection = (GraphInfoContainer(PVInclination).DataSingle(PVFrameNo) - GraphInfoContainer(PVDesignGradient).DataSingle(PVFrameNo)) * 12
'PCN6458            InclinationVal = (GraphInfoContainer(PVInclination).DataSingle(PVFrameNo)) '- GraphInfoContainer(PVDesignGradient).DataSingle(PVFrameNo)) * 1000
'PCN6458        Else
'PCN6458            InclinationValDeflection = (GraphInfoContainer(PVInclinationSmooth).DataSingle(PVFrameNo) - GraphInfoContainer(PVDesignGradient).DataSingle(PVFrameNo)) * 12
'PCN6458            InclinationVal = (GraphInfoContainer(PVInclinationSmooth).DataSingle(PVFrameNo)) '- GraphInfoContainer(PVDesignGradient).DataSingle(PVFrameNo)) * 1000
'PCN6458        End If
'PCN6458    End If
    
    If DiameterMedianVal > InvalidData Then
        DimensionValue.Caption = Format(DiameterMedianVal, "#0.0") & DisplayUnits
    Else
        DimensionValue.Caption = "%"
    End If
    
    
    If DebrisOn Then
        PVKey_Capacity_Value.Caption = Format(CapacityVal, "#0.00") & "%"
    Else
        If (CapacityVal > InvalidData) Then
            PVKey_Capacity_Value.Caption = Format(CapacityVal, "#0.0") & "%"
        Else
            PVKey_Capacity_Value.Caption = "%"
        End If
    End If
        
    If (OvalityVal > 1000) Or OvalityVal < -1000 Then
        PVKey_Ovality_Value.Caption = "%"
    Else
        PVKey_Ovality_Value.Caption = Format(OvalityVal, "#0.0") & "%"
    End If
    PVKey_DeltaMin_Value.Caption = Format(OvalityOrigVal, "#0.0") & "%"
    'PVKey_DeltaMin_Value.Caption = Format(ConvertRealToPer(DeltaMinVal, "Rad"), "#0.0") & "%"
    PVKey_DeltaMax_Value.Caption = Format(ConvertRealToPer(DeltaMaxVal, "Rad"), "#0.0") & "%"
'PCN6458     PVKey_Inclination_Value.Caption = Format(InclinationValDeflection, "#0.0") & DisplayUnits
    
'    PVKey_Ovality_Value_Unit.Caption = Format(ConvertDistPerToReal(PVOvalityFullData(PVFrameNo) / PVCalculationsMultiplier), "#0.0") & "mm"
    
    If MeasurementUnits = "mm" Then
        PVKey_Capacity_Value_Unit.Caption = Format(ConvertPerToReal(CapacityVal, "Area"), "#0.00") & "cm"
    Else
        PVKey_Capacity_Value_Unit.Caption = Format(ConvertPerToReal(CapacityVal, "Area"), "#0.00") & "in"
    End If
    
    If (CapacityVal <= InvalidData) Then
        PVKey_Capacity_Value_Unit.Caption = ""
    End If
    
    PVKey_DeltaMin_Value_Unit.Caption = Format(DeltaMinVal, "#0.0") & DisplayUnits
    PVKey_DeltaMax_Value_Unit.Caption = Format(DeltaMaxVal, "#0.0") & DisplayUnits
    
'PCN6458     If DisplayUnits = "mm" Then
'PCN6458         PVKey_Inclination_Value_Unit.Caption = Format(InclinationVal, "#0.0##") & "m"
'PCN6458     Else
'PCN6458         PVKey_Inclination_Value_Unit.Caption = Format(InclinationVal, "#0.0##") & "ft"
'PCN6458     End If
    
    'vvvv PCN3123 ***************************************
    If XDiameterVal > InvalidData Then
        If Not MedianFlat Then PVKey_XDiameter_Value.Caption = Format(ConvertRealToPer(XDiameterVal, "Dia"), "#0.0") & "%"
        If MedianFlat And PVDFileName <> "" Then PVKey_XDiameter_Value.Caption = Format(XDiameterVal, "#0.0") & "%"
        PVKey_XDiameter_Value_Unit.Caption = Format(XDiameterVal, "#0.0") & DisplayUnits
    Else
        PVKey_XDiameter_Value.Caption = "%"
        PVKey_XDiameter_Value_Unit.Caption = DisplayUnits
    End If
    If YDiameterVal > InvalidData Then
        If Not MedianFlat Then PVKey_YDiameter_Value.Caption = Format(ConvertRealToPer(YDiameterVal, "Dia"), "#0.0") & "%"
        If MedianFlat And PVDFileName <> "" Then PVKey_YDiameter_Value.Caption = Format(YDiameterVal, "#0.0") & "%"
        PVKey_YDiameter_Value_Unit.Caption = Format(YDiameterVal, "#0.0") & DisplayUnits
    Else
        PVKey_YDiameter_Value.Caption = "%"
        PVKey_YDiameter_Value_Unit.Caption = DisplayUnits
    End If
    '^^^^ ***********************************************
    'Distance
    If ConfigInfo.DistanceStart >= 0 Then
        If MeasurementUnits = "mm" Then
            PVKey_Distance_Value = Format(PVDistances(PVFrameNo), "#0.0") & "m"
        Else
            PVKey_Distance_Value = Format(PVDistances(PVFrameNo), "#0") & "ft"
        End If
    End If
    'Nominal Internal Diameter ?
    
    PVGraphsKeyForm.Visible = True
'PVGraphsKeyForm.ZOrder 0
    Call ScreenDrawing.FormTopMost(PVGraphsKeyForm.hwnd) 'PCN2990
Else
    PVGraphsKeyForm.Visible = False
End If



Exit Function
Err_Handler:
    Select Case Err
        Case 9 'Out of subscript
            Resume Next
        Case Else
            MsgBox Err & "-GKF14:" & Error$
    End Select
End Function

Private Sub PVKey_Flat3D_Value7_LostFocus()
On Error GoTo Err_Handler
Dim ColourVal(7) As Double
Dim NewValue As String
Dim PercentPos As Integer

PercentPos = InStr(PVKey_Flat3D_Value7.text, "%")
If PercentPos = 0 Then
    NewValue = PVKey_Flat3D_Value7.text
Else
    NewValue = Left(PVKey_Flat3D_Value7.text, PercentPos - 1)
End If

If Flat3dLimitR = SafeCDbl(NewValue) Then Exit Sub 'PCN4161

If IsNumeric(NewValue) Then
    Flat3dLimitR = SafeCDbl(NewValue) 'PCN4161
    If Flat3dLimitR < 0 Then Flat3dLimitR = Flat3dLimitR * -1
    Flat3dLimitL = Flat3dLimitR * -1 'PCN4185
End If
Call ThreeDColourValues(ColourVal)
'Call ScreenDrawing.GraphSelect("Flat", 0)
Call ScreenDrawing.PVFlat3DCalcCPP(1, PVDataNoOfLines)
    
Call FlatColourRealValues(ColourVal) 'PCN4250
Call DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True) 'PCN4370

Exit Sub
Err_Handler:
            MsgBox Err & "-GKF15:" & Error$
    
End Sub

Private Sub PVKey_Flat3D_Value0_LostFocus()
On Error GoTo Err_Handler
Dim ColourVal(7) As Double
    
    If IsNumeric(PVKey_Flat3D_Value0.text) Then
        Flat3dLimitL = SafeCDbl(PVKey_Flat3D_Value0.text) 'PCN4161
        If Flat3dLimitL > 0 Then Flat3dLimitL = Flat3dLimitL * -1
        Flat3dLimitR = Flat3dLimitL * -1 'PCN4185
    End If
    Call ThreeDColourValues(ColourVal)
    'Call ScreenDrawing.GraphSelect("Flat", 0)
    Call ScreenDrawing.PVFlat3DCalcCPP(1, PVDataNoOfLines)
    Call DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True) 'PCN4370
    
    
Exit Sub
Err_Handler:
            MsgBox Err & "-GKF16:" & Error$

End Sub

Sub ThreeDColourValues(ByRef ColourVal() As Double)  'PCN4185
On Error GoTo Err_Handler

ColourVal(3) = 0
ColourVal(4) = 0
ColourVal(2) = Flat3dLimitL / 3
ColourVal(5) = Flat3dLimitR / 3
ColourVal(1) = 2 * Flat3dLimitL / 3
ColourVal(6) = 2 * Flat3dLimitR / 3
ColourVal(0) = Flat3dLimitL
ColourVal(7) = Flat3dLimitR

PVKey_Flat3D_Value3.Caption = "0%"
PVKey_Flat3D_Value4.Caption = "0%"
PVKey_Flat3D_Value2.Caption = Format(ColourVal(2), "#0.0") & "%"
PVKey_Flat3D_Value5.Caption = Format(ColourVal(5), "#0.0") & "%"
PVKey_Flat3D_Value1.Caption = Format(ColourVal(1), "#0.0") & "%"
PVKey_Flat3D_Value6.Caption = Format(ColourVal(6), "#0.0") & "%"
PVKey_Flat3D_Value0.text = Format(ColourVal(0), "#0.0") & "%"
PVKey_Flat3D_Value7.text = Format(ColourVal(7), "#0.0") & "%"

    'PCN4920 this is all new
PVKey_Flat3D_Value3_Unit(1).Caption = "0%" 'PCN4920 (1) percentage of diameter
PVKey_Flat3D_Value4_Unit(1).Caption = "0%" 'PCN4920 (1) percentage of diameter
PVKey_Flat3D_Value2_Unit(1).Caption = Format(ColourVal(2) / 2, "#0.0") & "%" 'PCN4920 (1) percentage of diameter
PVKey_Flat3D_Value5_Unit(1).Caption = Format(ColourVal(5) / 2, "#0.0") & "%" 'PCN4920 (1) percentage of diameter
PVKey_Flat3D_Value1_Unit(1).Caption = Format(ColourVal(1) / 2, "#0.0") & "%" 'PCN4920 (1) percentage of diameter
PVKey_Flat3D_Value6_Unit(1).Caption = Format(ColourVal(6) / 2, "#0.0") & "%" 'PCN4920 (1) percentage of diameter
PVKey_Flat3D_Value0_Unit(1).Caption = Format(ColourVal(0) / 2, "#0.0") & "%" 'PCN4920 (1) percentage of diameter
PVKey_Flat3D_Value7_Unit(1).Caption = Format(ColourVal(7) / 2, "#0.0") & "%" 'PCN4920 (1) percentage of diameter



Exit Sub
Err_Handler:
    MsgBox Err & "-GKF17:" & Error$
End Sub

Sub FlatColourRealValues(ByRef ColourVal() As Double) 'PCN4250
On Error GoTo Err_Handler
If MeasurementUnits = "mm" Then
    PVKey_Flat3D_Value3_Unit(0).Caption = Format(ConvertPerToReal(ColourVal(3), "Flat"), "#0") & MeasurementUnits 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Value4_Unit(0).Caption = Format(ConvertPerToReal(ColourVal(4), "Flat"), "#0") & MeasurementUnits 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Value2_Unit(0).Caption = Format(ConvertPerToReal(ColourVal(2), "Flat"), "#0") & MeasurementUnits 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Value5_Unit(0).Caption = Format(ConvertPerToReal(ColourVal(5), "Flat"), "#0") & MeasurementUnits 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Value1_Unit(0).Caption = Format(ConvertPerToReal(ColourVal(1), "Flat"), "#0") & MeasurementUnits 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Value6_Unit(0).Caption = Format(ConvertPerToReal(ColourVal(6), "Flat"), "#0") & MeasurementUnits 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Value0_Unit(0).Caption = Format(ConvertPerToReal(ColourVal(0), "Flat"), "#0") & MeasurementUnits 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Value7_Unit(0).Caption = Format(ConvertPerToReal(ColourVal(7), "Flat"), "#0") & MeasurementUnits 'PCN4920 (0) is percentage measurements
    

Else
    PVKey_Flat3D_Value3_Unit(0).Caption = Format(ConvertPerToReal(ColourVal(3), "Flat"), "#0.0") & MeasurementUnits 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Value4_Unit(0).Caption = Format(ConvertPerToReal(ColourVal(4), "Flat"), "#0.0") & MeasurementUnits 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Value2_Unit(0).Caption = Format(ConvertPerToReal(ColourVal(2), "Flat"), "#0.0") & MeasurementUnits 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Value5_Unit(0).Caption = Format(ConvertPerToReal(ColourVal(5), "Flat"), "#0.0") & MeasurementUnits 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Value1_Unit(0).Caption = Format(ConvertPerToReal(ColourVal(1), "Flat"), "#0.0") & MeasurementUnits 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Value6_Unit(0).Caption = Format(ConvertPerToReal(ColourVal(6), "Flat"), "#0.0") & MeasurementUnits 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Value0_Unit(0).Caption = Format(ConvertPerToReal(ColourVal(0), "Flat"), "#0.0") & MeasurementUnits 'PCN4920 (0) is percentage measurements
    PVKey_Flat3D_Value7_Unit(0).Caption = Format(ConvertPerToReal(ColourVal(7), "Flat"), "#0.0") & MeasurementUnits 'PCN4920 (0) is percentage measurements
    


End If

Exit Sub
Err_Handler:
    MsgBox Err & "-GKF18:" & Error$
End Sub

