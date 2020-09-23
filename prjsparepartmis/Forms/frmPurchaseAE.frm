VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "NSDataCombo.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPurchaseAE 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12075
   Icon            =   "frmPurchaseAE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   12075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00C0FFFF&
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
      Index           =   7
      Left            =   3720
      MaxLength       =   100
      TabIndex        =   1
      Tag             =   "Name"
      Top             =   600
      Width           =   1515
   End
   Begin prjcmosxp.Liner Liner1 
      Height          =   30
      Left            =   120
      TabIndex        =   53
      Top             =   8475
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   53
   End
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   120
      ScaleHeight     =   630
      ScaleWidth      =   11805
      TabIndex        =   29
      Top             =   3120
      Width           =   11805
      Begin VB.TextBox txtCostPrice 
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
         Left            =   7830
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   240
         Width           =   1185
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00FFFFFF&
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
         Left            =   7095
         TabIndex        =   13
         Text            =   "0"
         Top             =   240
         Width           =   700
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox txtGross 
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
         Left            =   9075
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   240
         Width           =   1590
      End
      Begin ctrlNSDataCombo.NSDataCombo NSPart 
         Height          =   315
         Left            =   0
         TabIndex        =   11
         Top             =   225
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
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
      Begin lvButton.lvButtons_H btnAdd 
         Height          =   345
         Left            =   10800
         TabIndex        =   14
         Top             =   240
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
         Caption         =   "&Add To List"
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
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   17
         Left            =   9135
         TabIndex        =   34
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   7095
         TabIndex        =   33
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Cost Price"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   7860
         TabIndex        =   32
         Top             =   0
         Width           =   1290
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "PartID/Sparepart"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000011D&
         Height          =   240
         Index           =   8
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000011D&
         Height          =   240
         Index           =   2
         Left            =   2160
         TabIndex        =   30
         Top             =   0
         Width           =   1515
      End
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00C0FFFF&
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
      Height          =   285
      Left            =   10425
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6945
      Width           =   1425
   End
   Begin VB.TextBox txtRemarks 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Tag             =   "Remarks"
      Top             =   6720
      Width           =   4110
   End
   Begin VB.TextBox txtNetAmount 
      BackColor       =   &H00C0FFFF&
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
      Height          =   285
      Left            =   10425
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   8055
      Width           =   1425
   End
   Begin VB.TextBox txtTaxBase 
      BackColor       =   &H00E6FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   7305
      Width           =   1425
   End
   Begin VB.TextBox txtVat 
      BackColor       =   &H00E6FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   7680
      Width           =   1425
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
      Height          =   885
      Index           =   1
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Tag             =   "Name"
      Top             =   1320
      Width           =   4515
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   0
      Tag             =   "Name"
      Top             =   600
      Width           =   1515
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
      Left            =   7680
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   7
      Tag             =   "Name"
      Top             =   960
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
      Index           =   4
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   3
      Tag             =   "Name"
      Top             =   960
      Width           =   2955
   End
   Begin VB.TextBox txtGrossAmount 
      BackColor       =   &H00C0FFFF&
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
      Height          =   285
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6600
      Width           =   1425
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   240
      Picture         =   "frmPurchaseAE.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Remove"
      Top             =   3840
      Visible         =   0   'False
      Width           =   275
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
      Height          =   675
      Index           =   5
      Left            =   7680
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Tag             =   "Name"
      Top             =   1320
      Width           =   4275
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
      Left            =   8760
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   10
      Tag             =   "Name"
      Top             =   2040
      Width           =   3180
   End
   Begin VB.ComboBox cmbFP 
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
      ItemData        =   "frmPurchaseAE.frx":711C
      Left            =   4320
      List            =   "frmPurchaseAE.frx":7129
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   6720
      Width           =   1395
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
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   5
      Tag             =   "Name"
      Top             =   2280
      Width           =   2475
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2610
      Left            =   120
      TabIndex        =   35
      Top             =   3765
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   4604
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   275
      ForeColorFixed  =   -2147483640
      BackColorSel    =   1091552
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      GridColorUnpopulated=   -2147483633
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   345
      Left            =   9525
      TabIndex        =   17
      Top             =   8580
      Width           =   1125
      _ExtentX        =   1984
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
   Begin lvButton.lvButtons_H cmdUsrHistory 
      Height          =   345
      Left            =   120
      TabIndex        =   20
      Top             =   8580
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
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   345
      Left            =   10725
      TabIndex        =   18
      Top             =   8580
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
   Begin lvButton.lvButtons_H cmdPrintPreview 
      Height          =   345
      Left            =   7680
      TabIndex        =   19
      Top             =   8580
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      Caption         =   "Print Preview"
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
   Begin ctrlNSDataCombo.NSDataCombo NSSupplierID 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
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
   Begin ctrlNSDataCombo.NSDataCombo NSStatus 
      Height          =   315
      Left            =   7680
      TabIndex        =   9
      Top             =   2040
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
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
   Begin MSComCtl2.DTPicker DTDP 
      Height          =   315
      Left            =   7680
      TabIndex        =   6
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   60227587
      CurrentDate     =   40544
   End
   Begin lvButton.lvButtons_H cmdReceive 
      Height          =   345
      Left            =   6240
      TabIndex        =   54
      Top             =   8580
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      Caption         =   "Receive P.O"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DR No."
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
      Height          =   225
      Left            =   3120
      TabIndex        =   55
      Top             =   600
      Width           =   675
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "StatusCD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6120
      TabIndex        =   52
      Top             =   2040
      Width           =   675
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order Details (Receiving)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   225
      TabIndex        =   51
      Top             =   2820
      Width           =   3990
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Person:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6120
      TabIndex        =   50
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Purchased "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   6165
      TabIndex        =   49
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   48
      Top             =   945
      Width           =   1275
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   47
      Top             =   1305
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   46
      Top             =   585
      Width           =   1275
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   195
      Left            =   9630
      TabIndex        =   45
      Top             =   6945
      Width           =   735
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   165
      TabIndex        =   44
      Top             =   6465
      Width           =   990
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gross"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   195
      Index           =   0
      Left            =   9840
      TabIndex        =   43
      Top             =   6600
      Width           =   480
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   195
      Left            =   9360
      TabIndex        =   42
      Top             =   8085
      Width           =   1005
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TaxBase"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   195
      Left            =   9600
      TabIndex        =   41
      Top             =   7305
      Width           =   720
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VAT(12%)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   195
      Left            =   9480
      TabIndex        =   40
      Top             =   7680
      Width           =   885
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   6120
      TabIndex        =   39
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form of Payment:"
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
      Left            =   4320
      TabIndex        =   38
      Top             =   6480
      Width           =   1290
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel. No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   37
      Top             =   2280
      Width           =   1275
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   225
      TabIndex        =   36
      Top             =   120
      Width           =   1845
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   120
      Top             =   120
      Width           =   11820
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   120
      Top             =   2820
      Width           =   11820
   End
End
Attribute VB_Name = "frmPurchaseAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FORM_STATE
Public PK                   As Long
Public srcText              As TextBox

Dim cIRowCount              As Integer


Dim sSQL                    As String
Dim HaveAction              As Boolean
Dim blnSave                 As Boolean

Dim cCostPrice              As Currency
Dim cIGross                 As Currency 'Gross Amount
Dim cIAmount                As Currency 'Current Amount

Dim i                       As Integer


Dim RS                      As New Recordset
Dim RSPurchase              As New Recordset

Dim RSPurchaseOrder         As New Recordset
Dim RSPartUpdate            As New Recordset



Private Sub btnAdd_Click()
    If is_empty(NSSupplierID, False) = True Then Exit Sub
    If is_empty(NSPart, False) = True Then Exit Sub
    If is_empty(txtDescription, False) = True Then Exit Sub
    If is_empty(txtQty, False) = True Then Exit Sub
    If is_empty(txtCostPrice, False) = True Then Exit Sub
    If is_empty(txtGross, False) = True Then Exit Sub
    
    Dim CurrRow As Integer
    
    CurrRow = getFlexPos(Grid, 1, NSPart.Text)

    With Grid
        If CurrRow < 0 Then

            If .Rows = 2 And .TextMatrix(1, 5) = "" Then
                .TextMatrix(1, 1) = NSPart.Text
                .TextMatrix(1, 2) = txtDescription.Text
                .TextMatrix(1, 3) = toNumber(txtQty.Text)
                .TextMatrix(1, 4) = toMoney(txtCostPrice.Text)
                .TextMatrix(1, 5) = toMoney(txtGross.Text)

            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = NSPart.Text
                .TextMatrix(.Rows - 1, 2) = txtDescription.Text
                .TextMatrix(.Rows - 1, 3) = toNumber(txtQty.Text)
                .TextMatrix(.Rows - 1, 4) = toMoney(txtCostPrice.Text)
                .TextMatrix(.Rows - 1, 5) = toMoney(txtGross.Text)

                
                .Row = .Rows - 1
            End If
            
            cIRowCount = cIRowCount + 1
            
        Else
            If MsgBox("PartID/Sparepart already added. Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                
                cIGross = cIGross - toNumber(Grid.TextMatrix(.RowSel, 5))
                txtGrossAmount.Text = Format$(cIGross, "#,##0.00")
                
                cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 5))
                txtNetAmount.Text = Format$(cIAmount, "#,##0.00")
                
                .TextMatrix(CurrRow, 1) = NSPart.Text
                .TextMatrix(CurrRow, 2) = txtDescription.Text
                .TextMatrix(CurrRow, 3) = toNumber(txtQty.Text)
                .TextMatrix(CurrRow, 4) = toMoney(txtCostPrice.Text)
                .TextMatrix(CurrRow, 5) = toMoney(txtGross.Text)

            Else
                Exit Sub
            End If
        End If

        cIGross = cIGross + toNumber(txtGross.Text)
        txtGrossAmount.Text = Format$(cIGross, "#,##0.00")
        cIAmount = cIAmount + toNumber(txtGross.Text)
        txtNetAmount.Text = Format$(cIAmount, "#,##0.00")
        txtTaxBase.Text = toMoney(txtNetAmount.Text / 1.12)
        txtVat.Text = toMoney(txtNetAmount.Text - txtTaxBase.Text)
        
        'hltext the current row's column
        .ColSel = 5
        
        'Display a remove button
        Grid_Click
        ResetEntry
    End With

End Sub

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        cIGross = cIGross - toNumber(Grid.TextMatrix(.RowSel, 5))
        txtGrossAmount.Text = Format$(cIGross, "#,##0.00")
        
        cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 5))
        txtNetAmount.Text = Format$(cIAmount, "#,##0.00")
        
        txtTaxBase.Text = toMoney(txtNetAmount.Text / 1.12)
        txtVat.Text = toMoney(txtNetAmount.Text - txtTaxBase.Text)
        
        'Update the record count
        cIRowCount = cIRowCount - 1
        
        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.RowSel)
    End With
    
    btnRemove.Visible = False
    Grid_Click
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdPrintPreview_Click()
Dim RSPurchaseOrder As New Recordset

On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM qry_Purchase_Order_Detail WHERE PurchaseOrderID = " & PK

    If cIRowCount < 1 Then
        MsgBox "Please enter item(s) before you can save this record.", vbExclamation
        NSPart.SetFocus
        Exit Sub
    End If

    With rptPurchaseOrderInvoice
        .rptCN.Connect = ";DATABASE=" & App.Path & "\Database\CMOSXP_DB.mdb;PWD=qwerty123;"
        .rptCN.DatabaseName = App.Path & "\Database\CMOSXP_DB.mdb"
        .rptCN.RecordSource = sSQL
        
        .lblDate.Caption = Now
        .lblCompany.Caption = ACTIVE_COMPANY.COMPANYNAME
        .lblAddress.Caption = ACTIVE_COMPANY.ADDRESS
        
        .lblPONo.Caption = frmPurchaseAE.txtEntry(0).Text
        .lblSupplier.Caption = frmPurchaseAE.txtEntry(4).Text
        .lblSAddress.Caption = frmPurchaseAE.txtEntry(1).Text
        .lblTelNo.Caption = frmPurchaseAE.txtEntry(2).Text
        
        .lblPT.Caption = frmPurchaseAE.cmbFP.Text
        .lblSalesPerson.Caption = frmPurchaseAE.txtEntry(3).Text
        .lblPDate.Caption = frmPurchaseAE.DTDP.Value
        
        .txtQty.DataField = "Qty"
        .txtProductID.DataField = "PartID"
        .txtDescription.DataField = "PartDescription"
        .txtCostPrice.DataField = "CostPrice"
        .txtAmount.DataField = "GrossAmount"
        
        .lblGross.Caption = toMoney(frmPurchaseAE.txtGrossAmount.Text)
        .lblDiscount.Caption = toMoney(frmPurchaseAE.txtDesc.Text)
        .lblNetAmount.Caption = toMoney(frmPurchaseAE.txtNetAmount.Text)
        .lblTaxBase.Caption = toMoney(frmPurchaseAE.txtTaxBase.Text)
        .lblVAT.Caption = toMoney(frmPurchaseAE.txtVat.Text)
        
        .show vbModal
    End With


Exit Sub
ErrHandler:
    MsgBox "Error #: " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation
End Sub

Private Sub cmdReceive_Click()
    Dim lngrow          As Long

    On Error GoTo ErrTrack
    
    RSPurchaseOrder.CursorLocation = adUseClient
    If RSPurchaseOrder.State = adStateOpen Then RSPurchaseOrder.Close
    RSPurchaseOrder.Open "SELECT qry_Purchase_Order_Detail.* FROM qry_Purchase_Order_Detail WHERE PurchaseOrderID=" & PK, CN, adOpenDynamic, adLockOptimistic
        
    If RSPurchaseOrder![StatusID] = "RC" And RSPurchaseOrder![StatusDesc] = "RECEIVED" Then
        MsgBox "Purchase Order No." & txtEntry(0).Text & Space(1) & "already received.Please check it!", vbExclamation
        Exit Sub
    End If
        
    If MsgBox("Are you sure you want to receive this purchase order?", vbQuestion + vbYesNo) = vbYes Then
        
        Set RSPurchaseOrder = New ADODB.Recordset
        sSQL_Update "UPDATE Purchase_Order SET StatusID='RC', StatusDesc='RECEIVED' WHERE PurchaseOrderID=" & PK
    
        With Grid
            For lngrow = .FixedRows To .Rows - 1
            
            Set RSPartUpdate = New ADODB.Recordset
            If RSPartUpdate.State = adStateOpen Then RSPartUpdate.Close
            RSPartUpdate.Open "SELECT Spare_Parts.* FROM Spare_Parts WHERE PartID ='" & .TextMatrix(lngrow, 1) & "'", CN, adOpenDynamic, adLockOptimistic
            
            RSPartUpdate![Inventory] = RSPartUpdate![Inventory] + toNumber(.TextMatrix(lngrow, 3))
            RSPartUpdate![LastDateModified] = Now
            RSPartUpdate![ModifiedBy] = ACTIVE_USER.USERNAME
            RSPartUpdate.Update
            
            Next lngrow
        End With
    Else
        Exit Sub
    End If
    
    Unload Me
    
ErrTrack:
    Set RSPurchaseOrder = Nothing
    Set RSPartUpdate = Nothing

End Sub

Private Sub cmdSave_Click()
    If is_empty(cmbFP, False) = True Then Exit Sub
    If is_empty(NSSupplierID, False) = True Then Exit Sub
    If is_empty(NSStatus, False) = True Then Exit Sub
    
    If cIRowCount < 1 Then
        MsgBox "Please enter item(s) before you can save this record.", vbExclamation
        NSPart.SetFocus
        Exit Sub
    End If
 
    If MsgBox("This save the record. Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    CN.BeginTrans
    
    If State = AddStateMode Then
        RS.AddNew
        RS.Fields("PurchaseOrderID") = PK
        RS.Fields("DateEncoded") = Format(Now, "M/d/yyyy")
        RS.Fields("EncodedBy") = ACTIVE_USER.USERNAME
    Else
        RS.Fields("LastDateModified") = Now
        RS.Fields("ModifiedBy") = ACTIVE_USER.USERNAME
    End If
    
    With RS
        .Fields("DRNo") = txtEntry(7).Text
        .Fields("SupplierID") = NSSupplierID.Text
        .Fields("SupplierName") = txtEntry(4).Text
        .Fields("Address") = txtEntry(1).Text
        .Fields("TelNo") = txtEntry(2).Text
        .Fields("Date") = DTDP.Value
        .Fields("Instruction") = txtEntry(5).Text
        .Fields("Salesman") = txtEntry(3).Text
        .Fields("Remarks") = txtRemarks.Text
        .Fields("FOP") = cmbFP.Text
            
        .Fields("Gross") = txtGrossAmount.Text
        .Fields("Discount") = txtDesc.Text
        .Fields("TaxBase") = txtTaxBase.Text
        .Fields("VAT") = txtVat.Text
        .Fields("NetAmount") = txtNetAmount.Text
        .Fields("StatusID") = NSStatus.Text
        .Fields("StatusDesc") = txtEntry(6).Text
        
        .Update
    End With
     
    DeleteItems
    
    Dim c As Integer
    
    With Grid
        
        For c = 1 To cIRowCount
            .Row = c
            If State = AddStateMode Then
AddNew:
                Dim RSPurchaseDetail As New Recordset
                
                Set RSPurchaseDetail = New ADODB.Recordset
                RSPurchaseDetail.CursorLocation = adUseClient
                RSPurchaseDetail.Open "SELECT * FROM Purchase_Order_Details ", CN, adOpenStatic, adLockOptimistic

                RSPurchaseDetail.AddNew

                RSPurchaseDetail![PurchaseOrderID] = PK
                RSPurchaseDetail![SupplierID] = NSSupplierID.Text
                RSPurchaseDetail![PartID] = .TextMatrix(c, 1)
                RSPurchaseDetail![PartDescription] = .TextMatrix(c, 2)
                RSPurchaseDetail![Qty] = toNumber(.TextMatrix(c, 3))
                RSPurchaseDetail![CostPrice] = toMoney(.TextMatrix(c, 4))
                RSPurchaseDetail![GrossAmount] = toMoney(.TextMatrix(c, 5))
    
                RSPurchaseDetail.Update
                
            ElseIf State = EditStateMode Then
            
                Set RSPurchaseDetail = New ADODB.Recordset
                RSPurchaseDetail.CursorLocation = adUseClient
                RSPurchaseDetail.Open "SELECT * FROM Purchase_Order_Details WHERE PurchaseOrderID=" & PK, CN, adOpenDynamic, adLockOptimistic
            
                RSPurchaseDetail.Filter = "PartID = '" & .TextMatrix(c, 1) & "'"
                
                If RSPurchaseDetail.RecordCount = 0 Then GoTo AddNew
                RSPurchaseDetail![PartID] = .TextMatrix(c, 1)
                RSPurchaseDetail![PartDescription] = .TextMatrix(c, 2)
                RSPurchaseDetail![Qty] = toNumber(.TextMatrix(c, 3))
                RSPurchaseDetail![CostPrice] = toMoney(.TextMatrix(c, 4))
                RSPurchaseDetail![GrossAmount] = toMoney(.TextMatrix(c, 5))
    
                RSPurchaseDetail.Update
                      
            End If

        Next c
    End With

    'Clear variables
    c = 0
    
    ResetEntry
    
    CN.CommitTrans
    
    HaveAction = True
    
    If State = AddStateMode Then
        MsgBox "New purchase order received has been successfully saved.", vbInformation
        Unload Me
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If
    
  Exit Sub
  
err:
  CN.RollbackTrans
  MsgBox "Error: " & err.Description, vbExclamation
  Exit Sub
  If err.Number = -2147217887 Then Resume Next
End Sub

Private Sub cmdUsrHistory_Click()
On Error Resume Next
    Dim DE As String
    Dim DM As String
    Dim EB As String
    Dim MB As String
    
    DE = Format$(RS.Fields("DateEncoded"), "MMM-dd-yyyy HH:MM AMPM")
    DM = Format$(RS.Fields("LastDateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    EB = getValueAt("SELECT * FROM Purchase_Order WHERE PurchaseOrderID = " & PK, "EncodedBy")
    MB = getValueAt("SELECT * FROM Purchase_Order WHERE PurchaseOrderID = " & PK, "ModifiedBy")
    
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

Private Sub Form_Activate()
On Error Resume Next
    Me.BackColor = MAIN.ACPMenu.BackColor
    picPurchase.BackColor = MAIN.ACPMenu.BackColor
    
    cmbFP.ListIndex = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
ElseIf KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
    CenterForm frmPurchaseAE
    On Error GoTo ErrHandler
    
    InitGrid
    InitializeNSD
    btnAdd.Enabled = False
    
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM Purchase_Order WHERE PurchaseOrderID = " & PK, CN, adOpenStatic, adLockOptimistic

    RSPurchase.CursorLocation = adUseClient
    RSPurchase.Open "SELECT * FROM qry_Purchase_Order_Detail WHERE PurchaseOrderID = " & PK, CN, adOpenStatic, adLockOptimistic

    If State = AddStateMode Then

        cmdPrintPreview.Visible = False
        cmdReceive.Visible = False
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        
        NSStatus.Text = "RC"
        txtEntry(6).Text = "RECEIVED"
        
        NSStatus.DisableDropdown = True
        DTDP.Value = Format$(Now, "MMM-dd-yyyy")
        txtEntry(3).Text = ACTIVE_USER.FULLNAME
                 
        GeneratePK
        ResetEntry
        txtEntry(0).Text = PK

    Else
    
        Screen.MousePointer = vbHourglass
        cmdReceive.Visible = True
        cmdCancel.Caption = "Close"
        DTDP.Enabled = False

        Caption = "Modify Existing Entry"

        DisplayForEditing
        
        Screen.MousePointer = vbDefault
    End If
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RS = Nothing
    Set RSPurchaseOrder = Nothing
    Set RSPartUpdate = Nothing
    Set RSPurchase = Nothing
    Set frmPurchaseAE = Nothing
    
    cIGross = 0

End Sub

Private Sub Grid_Click()
If State = EditStateMode And NSStatus.Text = "RC" Then Exit Sub

With Grid
    NSPart.Text = .TextMatrix(.RowSel, 1)
    txtDescription.Text = .TextMatrix(.RowSel, 2)
    txtQty.Text = toNumber(.TextMatrix(.RowSel, 3))
    txtCostPrice.Text = toMoney(.TextMatrix(.RowSel, 4))
    txtGross.Text = toMoney(.TextMatrix(.RowSel, 5))
    
    If Grid.Rows = 2 And Grid.TextMatrix(1, 5) = "" Then
        btnRemove.Visible = False
    Else
        btnRemove.Visible = True
        btnRemove.Top = (Grid.CellTop + Grid.Top) - 20
        btnRemove.Left = Grid.Left + 50
    End If
End With
End Sub

Private Sub Grid_Scroll()
    btnRemove.Visible = False
End Sub

Private Sub Grid_SelChange()
    Grid_Click
End Sub

Private Sub nssupplierid_Change()
txtEntry(4).Text = NSSupplierID.getSelValueAt(2)
txtEntry(1).Text = NSSupplierID.getSelValueAt(3)
txtEntry(2).Text = NSSupplierID.getSelValueAt(4)
txtEntry(3).Text = NSSupplierID.getSelValueAt(5)
End Sub

Private Sub nsPart_Change()
txtDescription.Text = NSPart.getSelValueAt(2)
txtCostPrice.Text = toMoney(NSPart.getSelValueAt(3))
End Sub

Private Sub NSStatus_Change()
txtEntry(6).Text = NSStatus.getSelValueAt(2)
End Sub

Private Sub txtDisc_Change()
    txtQty_Change
End Sub

Private Sub txtDesc_GotFocus()
HLText txtDesc
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtDesc_LostFocus()
txtDesc.Text = toMoney(txtDesc.Text)
End Sub

Private Sub txtGrossAmount_GotFocus()
HLText txtGrossAmount
End Sub

Private Sub txtNetAmount_GotFocus()
HLText txtNetAmount
End Sub

Private Sub txtQty_Change()
    If toNumber(txtQty.Text) < 1 Then
        btnAdd.Enabled = False
        Exit Sub
    Else
        btnAdd.Enabled = True
    End If
    
    cCostPrice = toNumber(txtCostPrice.Text)
    txtGross.Text = toMoney((toNumber(txtQty.Text) * cCostPrice))
    
End Sub

Private Sub txtQty_GotFocus()
HLText txtQty
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
txtQty.Text = toNumber(txtQty.Text)
End Sub


Private Sub InitGrid()
    cIRowCount = 0
    With Grid
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 6
        .ColSel = 5
        
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 1800
        .ColWidth(2) = txtDescription.Width
        .ColWidth(3) = txtQty.Width + 50
        .ColWidth(4) = txtCostPrice.Width + 100
        .ColWidth(5) = 4200
        
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "PartID"
        .TextMatrix(0, 2) = "PartDescription"
        .TextMatrix(0, 3) = "Qty"
        .TextMatrix(0, 4) = "CostPrice"
        .TextMatrix(0, 5) = "Gross"

        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbLeftJustify
        .ColAlignment(3) = vbLeftJustify
        .ColAlignment(4) = vbLeftJustify
        .ColAlignment(5) = vbLeftJustify
    End With
End Sub


Private Sub InitializeNSD()
    With NSSupplierID
        .ClearColumn
        .AddColumn "SupplierID", 1500
        .AddColumn "Description", 2800
        .AddColumn "Address", 3500
        .AddColumn "BusinessNo", 1500
        .AddColumn "ContactPerson", 3500
        
        .Connection = CN.ConnectionString
        .SQLFields = "SupplierID,Description,Address,BusinessNo,ContactPerson"
        .sqlTables = "Suppliers"
        .sqlSortOrder = "SupplierID ASC"
        
        .BoundField = "SupplierID"
        .PageBy = 25
        .DisplayCol = 1
        
        .setDropWindowSize 6800, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Select Supplier"
    End With
    
    With NSStatus
        .ClearColumn
        .AddColumn "StatusID", 1200
        .AddColumn "Description", 4500
        
        .Connection = CN.ConnectionString
        .SQLFields = "StatusID,Description"
        .sqlTables = "Purchase_Status"
        .sqlSortOrder = "StatusID ASC"
        
        .BoundField = "StatusID"
        .PageBy = 25
        .DisplayCol = 1
        
        .setDropWindowSize 6800, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Select Status"
    End With
    
    With NSPart
        .ClearColumn
        .AddColumn "PartID", 1800
        .AddColumn "PartDescription", 3800
        .AddColumn "SupplierPrice", 1500
        .AddColumn "Inventory", 1500
        
        .Connection = CN.ConnectionString
        .SQLFields = "PartID,PartDescription,SupplierPrice,Inventory"
        .sqlTables = "Spare_Parts"
        .sqlSortOrder = "PartID ASC"
        
        .BoundField = "PartID"
        .PageBy = 25
        .DisplayCol = 1
        
        .setDropWindowSize 6800, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Select Spare_Parts/Item"
    End With
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim RSProduct As New Recordset
    
    If State = AddStateMode Then Exit Sub
    
    RSProduct.CursorLocation = adUseClient
    RSProduct.Open "SELECT * FROM Purchase_Order_Details WHERE PurchaseOrderID=" & PK, CN, adOpenDynamic, adLockOptimistic
    
    If RSProduct.RecordCount > 0 Then
        RSProduct.MoveFirst
        While Not RSProduct.EOF
            CurrRow = getFlexPos(Grid, 1, RSProduct!PartID)
        
            'Add to grid
            With Grid
                If CurrRow < 0 Then
                   DelRecwSQL "Purchase_Order_Details", "PurchaseOrderDetailID", "", True, RSProduct!SalesOrderDetailID
        
                End If
            End With
            RSProduct.MoveNext
        Wend
    End If
    
    Set RSProduct = Nothing
End Sub

Private Sub GeneratePK()
    PK = getIndex("Purchase_Order")
End Sub

Private Sub DisplayForEditing()
On Error GoTo err
    
    With RS
        txtEntry(0).Text = PK
        NSSupplierID.Text = .Fields("SupplierID")
        txtEntry(4).Text = .Fields("SupplierName")
        txtEntry(1).Text = .Fields("Address")
        txtEntry(2).Text = .Fields("TelNo")
        DTDP.Value = .Fields("Date")
        txtEntry(3).Text = .Fields("Salesman")
        txtEntry(5).Text = .Fields("Instruction")
        txtRemarks.Text = .Fields("Remarks")
        cmbFP.Text = .Fields("FOP")
        NSStatus.Text = .Fields("StatusID")
        txtEntry(6).Text = .Fields("StatusDesc")
        txtGrossAmount.Text = toMoney(.Fields("Gross"))
        txtDesc.Text = toMoney(.Fields("Discount"))
        txtTaxBase.Text = toMoney(.Fields("TaxBase"))
        txtVat.Text = toMoney(.Fields("VAT"))
        txtNetAmount.Text = toMoney(.Fields("NetAmount"))
        
        cIGross = toMoney(.Fields("Gross"))
        cIAmount = toMoney(.Fields("NetAmount"))

    End With

    cIRowCount = 0


    If RSPurchase.RecordCount > 0 Then
        RSPurchase.MoveFirst
        While Not RSPurchase.EOF
        
          cIRowCount = cIRowCount + 1
          
            With Grid
                If .Rows = 2 And .TextMatrix(1, 5) = "" Then
                    .TextMatrix(1, 1) = RSPurchase![PartID]
                    .TextMatrix(1, 2) = RSPurchase![PartDescription]
                    .TextMatrix(1, 3) = toNumber(RSPurchase![Qty])
                    .TextMatrix(1, 4) = toMoney(RSPurchase![CostPrice])
                    .TextMatrix(1, 5) = toMoney(RSPurchase![GrossAmount])

                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSPurchase![PartID]
                    .TextMatrix(.Rows - 1, 2) = RSPurchase![PartDescription]
                    .TextMatrix(.Rows - 1, 3) = toNumber(RSPurchase![Qty])
                    .TextMatrix(.Rows - 1, 4) = toMoney(RSPurchase![CostPrice])
                    .TextMatrix(.Rows - 1, 5) = toMoney(RSPurchase![GrossAmount])
                    
                    
                End If
            End With
            RSPurchase.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 5

        If State = EditStateMode Then
            If NSStatus.Text = "RC" Then
                Grid.FixedRows = Grid.Row: Grid.SelectionMode = flexSelectionFree
                Grid.FixedCols = 2

            Else
                Grid.FixedRows = Grid.Row:
                Grid.FixedCols = 1
                

            End If
        End If

    End If

    RSPurchase.Close
    Set RSPurchase = Nothing
    
    Exit Sub
    
err:
    If err.Number = 94 Then Resume Next
End Sub


Private Sub txtRemarks_GotFocus()
HLText txtRemarks
End Sub

Private Sub txtCostPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub ResetEntry()
    NSPart.Text = ""
    txtDescription.Text = ""
    txtQty.Text = "0"
    txtCostPrice.Text = "0.00"
    txtGross.Text = "0.00"
End Sub



