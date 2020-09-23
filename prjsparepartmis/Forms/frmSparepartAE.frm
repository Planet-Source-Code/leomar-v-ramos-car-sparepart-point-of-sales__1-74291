VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "NSDataCombo.ocx"
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "XPTab.ocx"
Begin VB.Form frmSparepartAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin prjXTab.XTab XPTab 
      Height          =   4815
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   8493
      TabCount        =   2
      TabCaption(0)   =   "Part"
      TabContCtrlCnt(0)=   22
      Tab(0)ContCtrlCap(1)=   "txtEntry11"
      Tab(0)ContCtrlCap(2)=   "txtEntry9"
      Tab(0)ContCtrlCap(3)=   "txtEntry8"
      Tab(0)ContCtrlCap(4)=   "txtEntry10"
      Tab(0)ContCtrlCap(5)=   "NSCategory"
      Tab(0)ContCtrlCap(6)=   "txtEntry0"
      Tab(0)ContCtrlCap(7)=   "txtEntry1"
      Tab(0)ContCtrlCap(8)=   "txtEntry2"
      Tab(0)ContCtrlCap(9)=   "txtEntry3"
      Tab(0)ContCtrlCap(10)=   "fraPhoto"
      Tab(0)ContCtrlCap(11)=   "Labels4"
      Tab(0)ContCtrlCap(12)=   "Labels3"
      Tab(0)ContCtrlCap(13)=   "Labels2"
      Tab(0)ContCtrlCap(14)=   "Label11"
      Tab(0)ContCtrlCap(15)=   "Label13"
      Tab(0)ContCtrlCap(16)=   "Labels12"
      Tab(0)ContCtrlCap(17)=   "Labels0"
      Tab(0)ContCtrlCap(18)=   "Labels1"
      Tab(0)ContCtrlCap(19)=   "Label1"
      Tab(0)ContCtrlCap(20)=   "Label5"
      Tab(0)ContCtrlCap(21)=   "Shape1"
      Tab(0)ContCtrlCap(22)=   "Shape3"
      TabCaption(1)   =   "Vehicle"
      TabContCtrlCnt(1)=   15
      Tab(1)ContCtrlCap(1)=   "NSCarType"
      Tab(1)ContCtrlCap(2)=   "NSCarMake"
      Tab(1)ContCtrlCap(3)=   "txtEntry4"
      Tab(1)ContCtrlCap(4)=   "txtEntry5"
      Tab(1)ContCtrlCap(5)=   "txtEntry6"
      Tab(1)ContCtrlCap(6)=   "txtEntry7"
      Tab(1)ContCtrlCap(7)=   "cboYear"
      Tab(1)ContCtrlCap(8)=   "Liner3"
      Tab(1)ContCtrlCap(9)=   "Label2"
      Tab(1)ContCtrlCap(10)=   "Label8"
      Tab(1)ContCtrlCap(11)=   "Label4"
      Tab(1)ContCtrlCap(12)=   "Label6"
      Tab(1)ContCtrlCap(13)=   "Label7"
      Tab(1)ContCtrlCap(14)=   "Label9"
      Tab(1)ContCtrlCap(15)=   "Shape2"
      TabStyle        =   1
      TabTheme        =   1
      ActiveTabBackStartColor=   16514555
      ActiveTabBackEndColor=   16514555
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   15397104
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   10198161
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   2880
         MaxLength       =   100
         TabIndex        =   42
         Tag             =   "Name"
         Top             =   2640
         Width           =   705
      End
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   40
         Tag             =   "Name"
         Top             =   1200
         Width           =   2280
      End
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   38
         Tag             =   "Name"
         Top             =   2640
         Width           =   705
      End
      Begin ctrlNSDataCombo.NSDataCombo NSCarType 
         Height          =   330
         Left            =   -73680
         TabIndex        =   31
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ctrlNSDataCombo.NSDataCombo NSCarMake 
         Height          =   330
         Left            =   -73680
         TabIndex        =   30
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   -72300
         MaxLength       =   100
         TabIndex        =   29
         Tag             =   "Name"
         Top             =   1080
         Width           =   1800
      End
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   -72300
         MaxLength       =   100
         TabIndex        =   28
         Tag             =   "Name"
         Top             =   1440
         Width           =   1800
      End
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   -73680
         MaxLength       =   100
         TabIndex        =   27
         Tag             =   "Name"
         Top             =   2040
         Width           =   1320
      End
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   -73680
         MaxLength       =   100
         TabIndex        =   26
         Tag             =   "Name"
         Top             =   2400
         Width           =   1320
      End
      Begin VB.ComboBox cboYear 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2760
         Width           =   1335
      End
      Begin prjcmosxp.Liner Liner3 
         Height          =   30
         Left            =   -74760
         TabIndex        =   24
         Top             =   1920
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   53
      End
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   21
         Tag             =   "Name"
         Top             =   3600
         Width           =   1185
      End
      Begin ctrlNSDataCombo.NSDataCombo NSCategory 
         Height          =   330
         Left            =   1320
         TabIndex        =   15
         Top             =   1920
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00E6FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   840
         Width           =   1365
      End
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   13
         Tag             =   "Name"
         Top             =   1560
         Width           =   3195
      End
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2700
         MaxLength       =   100
         TabIndex        =   12
         Tag             =   "Name"
         Top             =   1920
         Width           =   1780
      End
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   11
         Tag             =   "Name"
         Top             =   2280
         Width           =   1320
      End
      Begin VB.Frame fraPhoto 
         Height          =   4215
         Left            =   4800
         TabIndex        =   9
         Top             =   480
         Width           =   4815
         Begin lvButton.lvButtons_H cmdClear 
            Height          =   300
            Left            =   3915
            TabIndex        =   10
            Top             =   3780
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   529
            Caption         =   "Delete..."
            CapAlign        =   2
            BackStyle       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Image picFile 
            Height          =   3855
            Left            =   120
            Stretch         =   -1  'True
            ToolTipText     =   "Double click here to upload photo..."
            Top             =   240
            Width           =   4575
         End
      End
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ReOrder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   2160
         TabIndex        =   43
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   41
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   39
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   -74685
         TabIndex        =   37
         Top             =   600
         Width           =   2865
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Car Make"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   36
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Car Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   35
         Top             =   1440
         Width           =   660
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacity"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   34
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gear Box"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   33
         Top             =   2400
         Width           =   660
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year Model"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -74760
         TabIndex        =   32
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Price"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   3600
         Width           =   960
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Pricing Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   260
         TabIndex        =   22
         Top             =   3240
         Width           =   2865
      End
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   20
         Top             =   2280
         Width           =   345
      End
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PartID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   1575
         Width           =   795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   315
         TabIndex        =   17
         Top             =   480
         Width           =   2865
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   675
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H80000010&
         FillStyle       =   0  'Solid
         Height          =   240
         Left            =   240
         Top             =   480
         Width           =   4455
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H80000010&
         FillStyle       =   0  'Solid
         Height          =   240
         Left            =   240
         Top             =   3240
         Width           =   4455
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H80000010&
         FillStyle       =   0  'Solid
         Height          =   240
         Left            =   -74760
         Top             =   600
         Width           =   4335
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   9960
      TabIndex        =   5
      Top             =   0
      Width           =   9960
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "SPARE PART DETAILS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   3495
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmSparepartAE.frx":0000
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Note:Please fill all required parameters."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   960
         TabIndex        =   6
         Top             =   480
         Width           =   2895
      End
   End
   Begin prjcmosxp.Liner Liner2 
      Height          =   30
      Left            =   120
      TabIndex        =   3
      Top             =   6000
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   53
   End
   Begin prjcmosxp.Liner Liner1 
      Height          =   30
      Left            =   -240
      TabIndex        =   4
      Top             =   960
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   53
   End
   Begin lvButton.lvButtons_H cmdHistory 
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   6120
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   609
      Caption         =   "&Modification History"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   345
      Left            =   7440
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Caption         =   "&Save"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   345
      Left            =   8715
      TabIndex        =   1
      Top             =   6120
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      Caption         =   "&Cancel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin MSComDlg.CommonDialog CDCM 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSparepartAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                        As FORM_STATE
Public PK                           As String

Dim sCMY                            As Integer
Dim FN                              As String
Dim sSQL                            As String
Dim ImageBytes()                    As Byte

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdClear_Click()
On Error Resume Next
CommandPass "Delete"
End Sub

Private Sub cmdHistory_Click()
On Error Resume Next
    Dim DE As String
    Dim DM As String
    Dim EB As String
    Dim MB As String
    
    DE = Format$(RS_SPAREPART.Fields("DateEncoded"), "MMM-dd-yyyy HH:MM AMPM")
    DM = Format$(RS_SPAREPART.Fields("LastDateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    EB = getValueAt("SELECT * FROM Spare_Parts WHERE PartID = '" & txtEntry(0).Text & "'", "EncodedBy")
    MB = getValueAt("SELECT * FROM Spare_Parts WHERE PartID = '" & txtEntry(0).Text & "'", "ModifiedBy")
    
    MsgBox "Date Encoded: " & DE & vbCrLf & _
           "Encoded By: " & EB & vbCrLf & _
           "" & vbCrLf & _
           "Last Date Modified: " & DM & vbCrLf & _
           "Modified By: " & MB, vbInformation, "Modification History"
           
    DE = vbNullString
    DM = vbNullString
    EB = vbNullString
    MB = vbNullString

End Sub

Private Sub cmdSave_Click()
Dim obj As Control
            For Each obj In Me
            If TypeOf obj Is TextBox Or TypeOf obj Is ComboBox Then
                If obj.Text = "" Then
                    MsgBox obj.Name & " could not be left blank. Please complete the field.", vbExclamation
                    obj.SetFocus
                    Exit Sub
                End If
            End If
            Next obj
            
            If picFile.Picture = LoadPicture("") Then
                MsgBox "Please upload picture file for this car brand/make!", vbExclamation
                Exit Sub
            End If
            
            If State = AddStateMode Then
            
                If isRecordExist("Spare_Parts", "PartID", txtEntry(0).Text, True) = True Then
                    MsgBox "PartID already exist in the database.Please check it!", vbExclamation
                    Exit Sub
                End If
                
                RS_SPAREPART.AddNew
                RS_SPAREPART.Fields("PartID") = txtEntry(0).Text
                RS_SPAREPART.Fields("PartNumber") = txtEntry(9).Text
                RS_SPAREPART.Fields("PartDescription") = txtEntry(1).Text
                RS_SPAREPART.Fields("Price") = toMoney(txtEntry(3).Text)
                RS_SPAREPART.Fields("MakeID") = NSCarMake.Text
                RS_SPAREPART.Fields("MakeName") = txtEntry(4).Text
                RS_SPAREPART.Fields("CarTypeID") = NSCarType.Text
                RS_SPAREPART.Fields("CarTypeName") = txtEntry(5).Text
                RS_SPAREPART.Fields("PCategoryID") = NSCategory.Text
                RS_SPAREPART.Fields("PCategoryName") = txtEntry(2).Text
                RS_SPAREPART.Fields("Year") = cboYear.Text
                RS_SPAREPART.Fields("Capacity") = txtEntry(6).Text
                RS_SPAREPART.Fields("Gearbox") = txtEntry(7).Text
                RS_SPAREPART.Fields("Photo").AppendChunk ImageBytes
                RS_SPAREPART.Fields("Inventory") = toNumber(txtEntry(8).Text)
                RS_SPAREPART.Fields("ReOrder") = toNumber(txtEntry(11).Text)
                RS_SPAREPART.Fields("SupplierPrice") = txtEntry(10).Text
                RS_SPAREPART.Fields("DateEncoded") = Format(Now, "M/d/yyyy")
                RS_SPAREPART.Fields("EncodedBy") = ACTIVE_USER.USERNAME
                RS_SPAREPART.Update
                RS_SPAREPART.UpdateBatch adAffectAllChapters
                
                
                MsgBox "New spare part record has been successfully saved!", vbInformation
                SavePicture picFile.Picture, App.Path & "\Graphics\Spare Parts\" & txtEntry(0).Text & ".img"
                
                Unload Me
            
            ElseIf State = EditStateMode Then
                RS_SPAREPART.Fields("PartID") = txtEntry(0).Text
                RS_SPAREPART.Fields("PartNumber") = txtEntry(9).Text
                RS_SPAREPART.Fields("PartDescription") = txtEntry(1).Text
                RS_SPAREPART.Fields("Price") = toMoney(txtEntry(3).Text)
                RS_SPAREPART.Fields("MakeID") = NSCarMake.Text
                RS_SPAREPART.Fields("MakeName") = txtEntry(4).Text
                RS_SPAREPART.Fields("CarTypeID") = NSCarType.Text
                RS_SPAREPART.Fields("CarTypeName") = txtEntry(5).Text
                RS_SPAREPART.Fields("PCategoryID") = NSCategory.Text
                RS_SPAREPART.Fields("PCategoryName") = txtEntry(2).Text
                RS_SPAREPART.Fields("Year") = cboYear.Text
                RS_SPAREPART.Fields("Capacity") = txtEntry(6).Text
                RS_SPAREPART.Fields("Gearbox") = txtEntry(7).Text
                RS_SPAREPART.Fields("Photo").AppendChunk ImageBytes
                RS_SPAREPART.Fields("Inventory") = toNumber(txtEntry(8).Text)
                RS_SPAREPART.Fields("SupplierPrice") = txtEntry(10).Text
                RS_SPAREPART.Fields("LastDateModified") = Now
                RS_SPAREPART.Fields("ModifiedBy") = ACTIVE_USER.USERNAME
                RS_SPAREPART.Update
                RS_SPAREPART.UpdateBatch adAffectAllChapters
                
                MsgBox "Information saved successfully!", vbInformation
                SavePicture picFile.Picture, App.Path & "\Graphics\Spare Parts\" & txtEntry(0).Text & ".img"
                
                Unload Me
                
            End If

End Sub

Private Sub Form_Activate()
On Error Resume Next
Me.BackColor = MAIN.ACPMenu.BackColor

XPTab.TabStripBackColor = MAIN.ACPMenu.BackColor

fraPhoto.BackColor = MAIN.ACPMenu.BackColor

txtEntry(0).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
ElseIf KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
CenterForm frmSparepartAE
InitializeNSD

For sCMY = 1900 To 2100
    cboYear.AddItem sCMY
Next sCMY

If State = AddStateMode Then
    Me.Caption = "Create New Entry"
    txtEntry(0).Locked = False
    cmdHistory.Enabled = False
    
    Call LoadNullPicFile
    
ElseIf State = EditStateMode Then
    txtEntry(0).Locked = True
    Me.Caption = "Modify Entry"
    cmdHistory.Enabled = True
    
    sSQL = "SELECT Spare_Parts.* " & _
                "FROM Spare_Parts " & _
                "WHERE (((Spare_Parts.PartID)='" & PK & "'))"

    Set RS_SPAREPART = New ADODB.Recordset
    If RS_SPAREPART.State = adStateOpen Then RS_SPAREPART.Close
    RS_SPAREPART.Open sSQL, CN, adOpenDynamic, adLockOptimistic
    
    With RS_SPAREPART
        txtEntry(0).Text = .Fields("PartID")
        txtEntry(9).Text = .Fields("PartNumber")
        txtEntry(1).Text = .Fields("PartDescription")
        NSCategory.Text = .Fields("PCategoryID")
        txtEntry(2).Text = .Fields("PCategoryName")
        txtEntry(3).Text = toMoney(.Fields("Price"))
        NSCarMake.Text = .Fields("MakeID")
        txtEntry(4).Text = .Fields("MakeName")
        NSCarType.Text = .Fields("CarTypeID")
        txtEntry(5).Text = .Fields("CarTypeName")
        cboYear.Text = .Fields("Year")
        txtEntry(6).Text = .Fields("Capacity")
        txtEntry(7).Text = .Fields("Gearbox")
        txtEntry(8).Text = toNumber(.Fields("Inventory"))
        txtEntry(11).Text = toNumber(.Fields("ReOrder"))
        txtEntry(10).Text = toMoney(.Fields("SupplierPrice"))
    End With
    
    Call LoadPicFile
End If

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmSparepartAE = Nothing
Set RS_SPAREPART = Nothing

frmSparepart.CommandPass "Refresh"
End Sub


Private Sub NSCarMake_Change()
On Error Resume Next
    txtEntry(4).Text = NSCarMake.getSelValueAt(2)

    With NSCarType
        .ClearColumn
        .AddColumn "CarTypeID", 1700.882
        .AddColumn "CarTypeName", 1800
        .AddColumn "MakeID", 1800
        .AddColumn "MakeName", 2800
        .AddColumn "Remarks", 3100
        
        .Connection = CN.ConnectionString
        
        .SQLFields = "CarTypeID,CarTypeName,MakeID,MakeName,Remarks"
        .sqlTables = "Car_Types"
        .sqlwCondition = "MakeID='" & NSCarMake.Text & "'"
        .sqlSortOrder = "CarTypeName ASC"
        .BoundField = "CarTypeID"
        .PageBy = 10
        .DisplayCol = 1
        
        .setDropWindowSize 6500, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Select Car Type "
    End With
    
    NSCarType.Text = vbNullString
    txtEntry(5).Text = vbNullString
End Sub

Private Sub NSCarType_Change()
txtEntry(5).Text = NSCarType.getSelValueAt(2)
End Sub

Private Sub NSCategory_Change()
txtEntry(2).Text = NSCategory.getSelValueAt(2)
End Sub


Private Sub picFile_DblClick()
On Error Resume Next
CommandPass "Upload"
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
HLText txtEntry(Index)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 3 Or Index = 8 And Index = 10 Then
    KeyAscii = isNumber(KeyAscii)
End If
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
unHLText txtEntry(Index)

If Index = 3 Or Index = 8 Or Index = 10 Then
    txtEntry(3).Text = toMoney(txtEntry(3).Text)
    txtEntry(8).Text = toNumber(txtEntry(8).Text)
    txtEntry(10).Text = toMoney(txtEntry(10).Text)
End If
End Sub

Private Sub InitializeNSD()
On Error Resume Next
    With NSCarMake
        .ClearColumn
        .AddColumn "MakeID", 1700.882
        .AddColumn "MakeName", 1800
        .AddColumn "Remarks", 3100
        .Connection = CN.ConnectionString
        
        .SQLFields = "MakeID,MakeName,Remarks"
        .sqlTables = "Car_Makes"
        .sqlSortOrder = "MakeName ASC"
        .BoundField = "MakeID"
        .PageBy = 10
        .DisplayCol = 1
        
        .setDropWindowSize 6500, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Select Car Make "
    End With

    With NSCarType
        .ClearColumn
        .AddColumn "CarTypeID", 1700.882
        .AddColumn "CarTypeName", 1800
        .AddColumn "MakeID", 1800
        .AddColumn "MakeName", 2800
        .AddColumn "Remarks", 3100
        
        .Connection = CN.ConnectionString
        
        .SQLFields = "CarTypeID,CarTypeName,MakeID,MakeName,Remarks"
        .sqlTables = "Car_Types"
        .sqlSortOrder = "CarTypeName ASC"
        .BoundField = "CarTypeID"
        .PageBy = 10
        .DisplayCol = 1
        
        .setDropWindowSize 6500, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Select Car Type "
    End With
    
    With NSCategory
        .ClearColumn
        .AddColumn "PCategoryID", 1700.882
        .AddColumn "PCategoryName", 1800
        .AddColumn "Remarks", 3100
        .Connection = CN.ConnectionString
        
        .SQLFields = "PCategoryID,PCategoryName,Remarks"
        .sqlTables = "Part_Categories"
        .sqlSortOrder = "PCategoryName ASC"
        .BoundField = "PCategoryID"
        .PageBy = 10
        .DisplayCol = 1
        
        .setDropWindowSize 6500, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Select Part Category "
    End With
End Sub

Public Sub CommandPass(ByVal srcPerformWhat As String)
On Error GoTo errPerformWhat
Select Case srcPerformWhat
    Case "Upload" 'New
            Call BrowsePicFile

    Case "Delete"
            Call LoadNullPicFile
  
End Select
Exit Sub
errPerformWhat:
     MsgBox "Error Number:" & err.Number & vbNewLine & _
            "Description:" & err.Description, vbExclamation
End Sub

Private Sub LoadPicFile()
On Error Resume Next
    
    FN = App.Path & "\Graphics\Spare Parts\" & txtEntry(0).Text & ".img"
    
    With CDCM
        .CancelError = True
        .FileName = FN
        
        If .FileName <> "" Then
            Me.MousePointer = vbHourglass
            
            picFile.Picture = LoadPicture(.FileName)
            
            ReDim ImageBytes(FileLen(.FileName))
            Open .FileName For Binary As #1
                Get #1, , ImageBytes
            Close #1
        
            Me.MousePointer = vbDefault
        End If
    End With
    
End Sub


Private Sub BrowsePicFile()
On Error Resume Next
With CDCM
    .Filter = "JPG Files(*.jpeg)|*.jpg|JPEG Files(*.jpeg)|*.jpg" & _
    "|GIF Files(*.gif)|*.gif|Bitmap Files(*.bmp)|*.bmp|All Supported Files|*.jpeg;*.jpg;*.gif;*.bmp"
    
    .CancelError = True
    .ShowOpen
    
    If .FileName <> "" Then
        Me.MousePointer = vbHourglass
        
        picFile.Picture = LoadPicture(.FileName)
        
        ReDim ImageBytes(FileLen(.FileName))
        Open .FileName For Binary As #1
            Get #1, , ImageBytes
        Close #1
        
        Me.MousePointer = vbDefault
    Else
        Call LoadNullPicFile
    End If
End With
End Sub

Private Sub LoadNullPicFile()
On Error Resume Next
    
    FN = App.Path & "\Graphics\Spare Parts\" & "Null.img"
    
    With CDCM
        .CancelError = True
        .FileName = FN
        
        If .FileName <> "" Then
            Me.MousePointer = vbHourglass
            
            picFile.Picture = LoadPicture(.FileName)
            
            ReDim ImageBytes(FileLen(.FileName))
            Open .FileName For Binary As #1
                Get #1, , ImageBytes
            Close #1
        
            Me.MousePointer = vbDefault
        End If
    End With
End Sub


