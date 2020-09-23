VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "NSDataCombo.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSalesAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New Entry"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12405
   Icon            =   "frmSalesAE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows Default
   Begin prjcmosxp.Liner Liner1 
      Height          =   30
      Left            =   120
      TabIndex        =   56
      Top             =   8400
      Width           =   12255
      _ExtentX        =   21616
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
      Index           =   2
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   53
      Tag             =   "Name"
      Top             =   2280
      Width           =   4515
   End
   Begin VB.TextBox txtCA 
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
      Left            =   10905
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   8040
      Width           =   1425
   End
   Begin VB.TextBox txtCT 
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
      Left            =   10905
      TabIndex        =   18
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
      Height          =   315
      Index           =   6
      Left            =   9120
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   14
      Tag             =   "Name"
      Top             =   2040
      Width           =   3180
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
      Left            =   8040
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Tag             =   "Name"
      Top             =   1320
      Width           =   4275
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   240
      Picture         =   "frmSalesAE.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Remove"
      Top             =   3840
      Visible         =   0   'False
      Width           =   275
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
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6600
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
      Height          =   315
      Index           =   4
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   2
      Tag             =   "Name"
      Top             =   960
      Width           =   2955
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
      Left            =   8040
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   4
      Tag             =   "Name"
      Top             =   960
      Width           =   3195
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
      Height          =   885
      Index           =   1
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Tag             =   "Name"
      Top             =   1320
      Width           =   4515
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
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   11280
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
      Left            =   4185
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   11265
      Width           =   1425
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
      Left            =   10905
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   7335
      Width           =   1425
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
      Left            =   10905
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6885
      Width           =   1425
   End
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   120
      ScaleHeight     =   630
      ScaleWidth      =   12165
      TabIndex        =   26
      Top             =   3120
      Width           =   12165
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
         Left            =   9555
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   240
         Width           =   1590
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
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   5175
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
         Left            =   7575
         TabIndex        =   8
         Text            =   "0"
         Top             =   240
         Width           =   700
      End
      Begin VB.TextBox txtUnitPrice 
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
         Left            =   8310
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   240
         Width           =   1185
      End
      Begin ctrlNSDataCombo.NSDataCombo NSPart 
         Height          =   315
         Left            =   0
         TabIndex        =   6
         Top             =   225
         Width           =   2280
         _ExtentX        =   4022
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
         Left            =   11160
         TabIndex        =   11
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
         Left            =   2400
         TabIndex        =   45
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Part ID/Spare ID"
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
         TabIndex        =   30
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
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
         Left            =   8340
         TabIndex        =   29
         Top             =   0
         Width           =   1290
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
         Left            =   7575
         TabIndex        =   28
         Top             =   0
         Width           =   660
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
         Left            =   9615
         TabIndex        =   27
         Top             =   0
         Width           =   1260
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2610
      Left            =   120
      TabIndex        =   31
      Top             =   3765
      Width           =   12195
      _ExtentX        =   21511
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
      Left            =   9960
      TabIndex        =   22
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
      TabIndex        =   25
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
      Left            =   11205
      TabIndex        =   23
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
      Left            =   7560
      TabIndex        =   24
      Top             =   8580
      Width           =   1605
      _ExtentX        =   2831
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
   Begin ctrlNSDataCombo.NSDataCombo NSCustomer 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
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
      Left            =   8040
      TabIndex        =   13
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
      Left            =   8040
      TabIndex        =   52
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
      Format          =   64815107
      CurrentDate     =   40544
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   1890
      Left            =   840
      TabIndex        =   57
      Top             =   10320
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   3334
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
      TabIndex        =   12
      Tag             =   "Remarks"
      Top             =   6735
      Width           =   4110
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
      ItemData        =   "frmSalesAE.frx":711C
      Left            =   4320
      List            =   "frmSalesAE.frx":7126
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   6720
      Width           =   1395
   End
   Begin lvButton.lvButtons_H cmdAdd 
      Height          =   315
      Left            =   6000
      TabIndex        =   58
      ToolTipText     =   "New Customer..."
      Top             =   960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmSalesAE.frx":7137
      cBack           =   -2147483633
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
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
      TabIndex        =   55
      Top             =   120
      Width           =   1845
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No."
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
      TabIndex        =   54
      Top             =   2280
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
      TabIndex        =   51
      Top             =   6480
      Width           =   1290
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Amount"
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
      Left            =   9495
      TabIndex        =   49
      Top             =   8070
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Tendered"
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
      Left            =   9585
      TabIndex        =   48
      Top             =   7710
      Width           =   1260
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Special Instructions"
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
      Left            =   6480
      TabIndex        =   47
      Top             =   1320
      Width           =   1275
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
      Left            =   5640
      TabIndex        =   44
      Top             =   11040
      Width           =   885
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
      Left            =   4200
      TabIndex        =   43
      Top             =   11040
      Width           =   720
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
      Left            =   9840
      TabIndex        =   42
      Top             =   7365
      Width           =   1005
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
      Left            =   10320
      TabIndex        =   41
      Top             =   6600
      Width           =   480
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
      TabIndex        =   40
      Top             =   6465
      Width           =   990
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
      Left            =   10110
      TabIndex        =   39
      Top             =   6915
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales No."
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
      TabIndex        =   38
      Top             =   585
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
      TabIndex        =   37
      Top             =   1305
      Width           =   1275
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer:"
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
      TabIndex        =   36
      Top             =   945
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
      Left            =   6525
      TabIndex        =   35
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Salesman/Agent"
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
      Left            =   6480
      TabIndex        =   34
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order Details"
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
      TabIndex        =   33
      Top             =   2820
      Width           =   1845
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
      Left            =   6480
      TabIndex        =   32
      Top             =   2040
      Width           =   675
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   120
      Top             =   2820
      Width           =   12180
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   120
      Top             =   120
      Width           =   12180
   End
End
Attribute VB_Name = "frmSalesAE"
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

Dim cSalesPrice             As Currency
Dim cIGross                 As Currency
Dim cIAmount                As Currency

Dim i                       As Integer

Dim lngrow1                 As Long
Dim lngrow2                 As Long

Dim RS                      As New Recordset
Dim RSSales                 As New Recordset
Dim RSPartUpdate            As Recordset



Private Sub btnAdd_Click()
    If is_empty(NSCustomer, False) = True Then Exit Sub
    If is_empty(NSPart, False) = True Then Exit Sub
    If is_empty(txtDescription, False) = True Then Exit Sub
    If is_empty(txtQty, False) = True Then Exit Sub
    If is_empty(txtUnitPrice, False) = True Then Exit Sub
    If is_empty(txtGross, False) = True Then Exit Sub
    
    If toNumber(txtQty.Text) < 1 Then
        MsgBox "Quantity should be greater than zero value.Please check it!", vbExclamation
        Exit Sub
    End If
    
    If toNumber(txtQty.Text) > toNumber(NSPart.getSelValueAt(4)) Then
        MsgBox "Quantity should not be greater than to the actual stock inventory.Please check it!", vbExclamation
        Exit Sub
    End If
    
    Dim CurrRow As Integer
    
    CurrRow = getFlexPos(Grid, 1, NSPart.Text)

    With Grid
        If CurrRow < 0 Then

            If .Rows = 2 And .TextMatrix(1, 5) = "" Then
                .TextMatrix(1, 1) = NSPart.Text
                .TextMatrix(1, 2) = txtDescription.Text
                .TextMatrix(1, 3) = toNumber(txtQty.Text)
                .TextMatrix(1, 4) = toMoney(txtUnitPrice.Text)
                .TextMatrix(1, 5) = toMoney(txtGross.Text)

            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = NSPart.Text
                .TextMatrix(.Rows - 1, 2) = txtDescription.Text
                .TextMatrix(.Rows - 1, 3) = toNumber(txtQty.Text)
                .TextMatrix(.Rows - 1, 4) = toMoney(txtUnitPrice.Text)
                .TextMatrix(.Rows - 1, 5) = toMoney(txtGross.Text)

                
                .Row = .Rows - 1
            End If
            
            cIRowCount = cIRowCount + 1
            
        Else
            If MsgBox("PartID/Description already added. Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                
                cIGross = cIGross - toNumber(Grid.TextMatrix(.RowSel, 5))
                txtGrossAmount.Text = Format$(cIGross, "#,##0.00")
                
                cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 5))
                txtNetAmount.Text = Format$(cIAmount, "#,##0.00")
                
                .TextMatrix(CurrRow, 1) = NSPart.Text
                .TextMatrix(CurrRow, 2) = txtDescription.Text
                .TextMatrix(CurrRow, 3) = toNumber(txtQty.Text)
                .TextMatrix(CurrRow, 4) = toMoney(txtUnitPrice.Text)
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
                
        Grid2.TextMatrix(1, 1) = NSPart.Text
        Grid2.TextMatrix(1, 2) = txtDescription.Text
        Grid2.TextMatrix(1, 3) = toNumber(txtQty.Text)
        Grid2.TextMatrix(1, 4) = toMoney(txtUnitPrice.Text)
        Grid2.TextMatrix(1, 5) = toMoney(txtGross.Text)
        
    End With
    
    btnRemove.Visible = False
    
    Grid_Click
    txtDesc_Change
    txtCT_Change
End Sub

Private Sub cmdAdd_Click()
On Error Resume Next

Set RS_CUSTOMER = New ADODB.Recordset
RS_CUSTOMER.CursorLocation = adUseClient
RS_CUSTOMER.Open "SELECT * FROM Customers ", CN, adOpenDynamic, adLockOptimistic

frmCustomer.CommandPass "New"

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdPrintPreview_Click()
Dim RSSalesOrder            As New Recordset
Dim jSQL                    As String ' for Sales Order Detail

On Error GoTo ErrHandler
    
    jSQL = "SELECT * FROM qry_Sales_Order_Detail WHERE SalesOrderID = " & PK

    If cIRowCount < 1 Then
        MsgBox "Please enter item(s) before you can save this record.", vbExclamation
        NSPart.SetFocus
        Exit Sub
    End If
    
    Set RSSalesOrder = New ADODB.Recordset
    RSSalesOrder.CursorLocation = adUseClient
    sSQL_Update "UPDATE Sales_Order SET StatusID='CM', StatusDesc='COMPLETED' WHERE SalesOrderID=" & PK

    With rptSalesInvoice
        .rptCN.Connect = ";DATABASE=" & App.Path & "\Database\CMOSXP_DB.mdb;PWD=qwerty123;"
        .rptCN.DatabaseName = App.Path & "\Database\CMOSXP_DB.mdb"
        .rptCN.RecordSource = jSQL
        
        .lblDate.Caption = Now
        .lblCompany.Caption = ACTIVE_COMPANY.COMPANYNAME
        .lblAddress.Caption = ACTIVE_COMPANY.ADDRESS
        
        .lblSONo.Caption = frmSalesAE.txtEntry(0).Text
        .lblSoldTo.Caption = frmSalesAE.txtEntry(4).Text
        .lblCAddress.Caption = frmSalesAE.txtEntry(1).Text
        .lblTelNo.Caption = frmSalesAE.txtEntry(2).Text
        
        .lblPT.Caption = frmSalesAE.cmbFP.Text
        .lblSalesPerson.Caption = frmSalesAE.txtEntry(3).Text
        .lblSDate.Caption = frmSalesAE.DTDP.Value
        
        .txtQty.DataField = "Qty"
        .txtPartID.DataField = "PartID"
        .txtDescription.DataField = "PartDescription"
        .txtUnitPrice.DataField = "UnitPrice"
        .txtAmount.DataField = "GrossAmount"

        .lblGross.Caption = toMoney(frmSalesAE.txtGrossAmount.Text)
        .lblDiscount.Caption = toMoney(frmSalesAE.txtDesc.Text)
        .lblNetAmount.Caption = toMoney(frmSalesAE.txtNetAmount.Text)
        .lblTaxBase.Caption = toMoney(frmSalesAE.txtTaxBase.Text)
        .lblVAT.Caption = toMoney(frmSalesAE.txtVat.Text)
        .lblCT.Caption = toMoney(frmSalesAE.txtCT.Text)
        .lblCA.Caption = toMoney(frmSalesAE.txtCA.Text)
        
        .show vbModal
    End With


Exit Sub
ErrHandler:
    MsgBox "Error #: " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation
End Sub

Private Sub cmdSave_Click()
    If cIRowCount < 1 Then
        MsgBox "Please enter item(s) before you can save this record.", vbExclamation
        NSPart.SetFocus
        Exit Sub
    End If
    
    If toNumber(txtCT.Text) = 0 Then
        MsgBox "Cash Tendered should greater than zero value.Please check it!", vbExclamation
        Exit Sub
    End If
    
    If toNumber(txtCT.Text) < toNumber(txtNetAmount.Text) Then
        MsgBox "Cash Tendered should not be less than the Net Amount.Please check it!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("This save the record. Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    CN.BeginTrans
    
    If State = AddStateMode Then
        RS.AddNew
        RS.Fields("SalesOrderID") = PK
        RS.Fields("DateEncoded") = Format(Now, "M/d/yyyy")
        RS.Fields("EncodedBy") = ACTIVE_USER.USERNAME

        With Grid
            For lngrow1 = .FixedRows To .Rows - 1
                Set RSPartUpdate = New ADODB.Recordset
                RSPartUpdate.Open "SELECT Spare_Parts.* FROM Spare_Parts WHERE PartID ='" & .TextMatrix(lngrow1, 1) & "'", CN, adOpenDynamic, adLockOptimistic
                RSPartUpdate![Inventory] = RSPartUpdate![Inventory] - toNumber(.TextMatrix(lngrow1, 3))
                RSPartUpdate![LastDateModified] = Now
                RSPartUpdate![ModifiedBy] = ACTIVE_USER.USERNAME
                RSPartUpdate.Update
            Next lngrow1
        End With
    
    Else
    
        RS.Fields("LastDateModified") = Now
        RS.Fields("ModifiedBy") = ACTIVE_USER.USERNAME
        
        If NSStatus.Text = "CL" And txtEntry(6).Text = "CANCELLED" Then
            With Grid
                For lngrow1 = .FixedRows To .Rows - 1
                    Set RSPartUpdate = New ADODB.Recordset
                    RSPartUpdate.Open "SELECT Spare_Parts.* FROM Spare_Parts WHERE PartID ='" & .TextMatrix(lngrow1, 1) & "'", CN, adOpenDynamic, adLockOptimistic
                    RSPartUpdate![Inventory] = RSPartUpdate![Inventory] + toNumber(.TextMatrix(lngrow1, 3))
                    RSPartUpdate![LastDateModified] = Now
                    RSPartUpdate![ModifiedBy] = ACTIVE_USER.USERNAME
                    RSPartUpdate.Update
                Next lngrow1
            End With
        ElseIf NSStatus.Text = "OP" And txtEntry(6).Text = "OPEN" Then
            With Grid
                For lngrow1 = .FixedRows To .Rows - 1
                    Set RSPartUpdate = New ADODB.Recordset
                    RSPartUpdate.Open "SELECT Spare_Parts.* FROM Spare_Parts WHERE PartID ='" & .TextMatrix(lngrow1, 1) & "'", CN, adOpenDynamic, adLockOptimistic
                    RSPartUpdate![Inventory] = RSPartUpdate![Inventory] - toNumber(.TextMatrix(lngrow1, 3))
                    RSPartUpdate![LastDateModified] = Now
                    RSPartUpdate![ModifiedBy] = ACTIVE_USER.USERNAME
                    RSPartUpdate.Update
                Next lngrow1
            End With
        ElseIf NSStatus.Text = "CM" And txtEntry(6).Text = "COMPLETED" Then
            With Grid
                For lngrow1 = .FixedRows To .Rows - 1
                    Set RSPartUpdate = New ADODB.Recordset
                    RSPartUpdate.Open "SELECT Spare_Parts.* FROM Spare_Parts WHERE PartID ='" & .TextMatrix(lngrow1, 1) & "'", CN, adOpenDynamic, adLockOptimistic
                    RSPartUpdate![Inventory] = RSPartUpdate![Inventory] - toNumber(.TextMatrix(lngrow1, 3))
                    RSPartUpdate![LastDateModified] = Now
                    RSPartUpdate![ModifiedBy] = ACTIVE_USER.USERNAME
                    RSPartUpdate.Update
                Next lngrow1
            End With
        End If
    
        With Grid2
            For lngrow2 = .FixedRows To .Rows - 1
                If .TextMatrix(lngrow2, 1) = "" Then
                    'DO NOTHING
                Else
                    Set RSPartUpdate = New ADODB.Recordset
                    RSPartUpdate.Open "SELECT Spare_Parts.* FROM Spare_Parts WHERE PartID ='" & .TextMatrix(lngrow2, 1) & "'", CN, adOpenDynamic, adLockOptimistic
                    RSPartUpdate![Inventory] = RSPartUpdate![Inventory] + toNumber(.TextMatrix(lngrow2, 3))
                    RSPartUpdate![LastDateModified] = Now
                    RSPartUpdate![ModifiedBy] = ACTIVE_USER.USERNAME
                    RSPartUpdate.Update
                End If
            Next lngrow2
        End With
        
    End If
    
    With RS
        .Fields("CustomerID") = NSCustomer.Text
        .Fields("CustomerName") = txtEntry(4).Text
        .Fields("Address") = txtEntry(1).Text
        .Fields("TelNo") = txtEntry(2).Text
        .Fields("Date") = DTDP.Value
        .Fields("SInstruction") = txtEntry(5).Text
        .Fields("Salesman") = ACTIVE_USER.FULLNAME
        .Fields("Remarks") = txtRemarks.Text
        .Fields("FOP") = cmbFP.Text
            
        .Fields("Gross") = txtGrossAmount.Text
        .Fields("Discount") = txtDesc.Text
        .Fields("TaxBase") = txtTaxBase.Text
        .Fields("VAT") = txtVat.Text
        .Fields("NetAmount") = txtNetAmount.Text
        .Fields("CT") = txtCT.Text
        .Fields("CA") = txtCA.Text
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
                Dim RSSalesDetail As New Recordset
                
                Set RSSalesDetail = New ADODB.Recordset
                RSSalesDetail.CursorLocation = adUseClient
                RSSalesDetail.Open "SELECT * FROM Sales_Order_Details ", CN, adOpenStatic, adLockOptimistic

                RSSalesDetail.AddNew

                RSSalesDetail![SalesOrderID] = PK
                RSSalesDetail![CustomerID] = NSCustomer.Text
                RSSalesDetail![PartID] = .TextMatrix(c, 1)
                RSSalesDetail![PartDescription] = .TextMatrix(c, 2)
                RSSalesDetail![Qty] = toNumber(.TextMatrix(c, 3))
                RSSalesDetail![UnitPrice] = toMoney(.TextMatrix(c, 4))
                RSSalesDetail![GrossAmount] = toMoney(.TextMatrix(c, 5))
    
                RSSalesDetail.Update
                
            ElseIf State = EditStateMode Then
            
                Set RSSalesDetail = New ADODB.Recordset
                RSSalesDetail.CursorLocation = adUseClient
                RSSalesDetail.Open "SELECT * FROM Sales_Order_Details WHERE SalesOrderID=" & PK, CN, adOpenDynamic, adLockOptimistic
            
                RSSalesDetail.Filter = "PartID = '" & .TextMatrix(c, 1) & "'"
                
                If RSSalesDetail.RecordCount = 0 Then GoTo AddNew
                RSSalesDetail![PartID] = .TextMatrix(c, 1)
                RSSalesDetail![PartDescription] = .TextMatrix(c, 2)
                RSSalesDetail![Qty] = toNumber(.TextMatrix(c, 3))
                RSSalesDetail![UnitPrice] = toMoney(.TextMatrix(c, 4))
                RSSalesDetail![GrossAmount] = toMoney(.TextMatrix(c, 5))
    
                RSSalesDetail.Update
                      
            End If

        Next c
    End With

    'Clear variables
    c = 0
    
    ResetEntry
    
    CN.CommitTrans
    
    HaveAction = True
    
    If State = AddStateMode Then
        MsgBox "New sales record has been successfully saved.", vbInformation
        
        cmdPrintPreview_Click
        
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
    
    EB = getValueAt("SELECT * FROM Sales_Order WHERE SalesOrderID = " & PK, "EncodedBy")
    MB = getValueAt("SELECT * FROM Sales_Order WHERE SalesOrderID = " & PK, "ModifiedBy")
    
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
    CenterForm frmSalesAE
    On Error GoTo ErrHandler
    
    InitGrid
    InitializeNSD
    btnAdd.Enabled = False
    
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM Sales_Order WHERE SalesOrderID = " & PK, CN, adOpenStatic, adLockOptimistic

    RSSales.CursorLocation = adUseClient
    RSSales.Open "SELECT * FROM qry_Sales_Order_Detail WHERE SalesOrderID = " & PK, CN, adOpenStatic, adLockOptimistic

    If State = AddStateMode Then

        Me.cmdPrintPreview.Visible = False
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        
        NSStatus.Text = "OP"
        txtEntry(6).Text = "OPEN"
        
        NSStatus.DisableDropdown = True
        DTDP.Value = Format$(Now, "MMM-dd-yyyy")
        txtEntry(3).Text = ACTIVE_USER.FULLNAME
                 
        GeneratePK
        ResetEntry
        txtEntry(0).Text = PK

    Else
    
        Screen.MousePointer = vbHourglass
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
    Set RSPartUpdate = Nothing
    Set RSSales = Nothing
    Set frmSalesAE = Nothing
    
    lngrow1 = 0
    lngrow2 = 0
    cIGross = 0
    frmSales.CommandPass "Refresh"
End Sub

Private Sub Grid_Click()
If State = EditStateMode And NSStatus.Text = "CM" Then Exit Sub

    With Grid
        NSPart.Text = .TextMatrix(.RowSel, 1)
        txtDescription.Text = .TextMatrix(.RowSel, 2)
        txtQty.Text = toNumber(.TextMatrix(.RowSel, 3))
        txtUnitPrice.Text = toMoney(.TextMatrix(.RowSel, 4))
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

Private Sub NSCustomer_Change()
txtEntry(4).Text = NSCustomer.getSelValueAt(2)
txtEntry(1).Text = NSCustomer.getSelValueAt(3)
txtEntry(2).Text = NSCustomer.getSelValueAt(4)
End Sub

Private Sub nsPart_Change()
txtDescription.Text = NSPart.getSelValueAt(2)
txtUnitPrice.Text = toMoney(NSPart.getSelValueAt(3))
End Sub

Private Sub NSStatus_Change()
txtEntry(6).Text = NSStatus.getSelValueAt(2)
End Sub

Private Sub txtDisc_Change()
    txtQty_Change
End Sub


Private Sub txtCA_GotFocus()
HLText txtCA
End Sub

Private Sub txtCT_Change()
txtCA.Text = toMoney(toNumber(txtCT.Text) - toNumber(txtNetAmount.Text))
End Sub

Private Sub txtCT_GotFocus()
HLText txtCT
End Sub

Private Sub txtCT_KeyPress(KeyAscii As Integer)
KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtCT_LostFocus()
txtCT.Text = toMoney(txtCT.Text)
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
    
    cSalesPrice = toNumber(txtUnitPrice.Text)
    txtGross.Text = toMoney((toNumber(txtQty.Text) * cSalesPrice))
    
End Sub

Private Sub txtQty_GotFocus()
HLText txtQty
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
txtQty.Text = toNumber(txtQty.Text)
End Sub


Private Sub txtRemarks_GotFocus()
HLText txtRemarks
End Sub

Private Sub txtUnitPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub ResetEntry()
    NSPart.Text = ""
    txtDescription.Text = ""
    txtQty.Text = "0"
    txtUnitPrice.Text = "0.00"
    txtGross.Text = "0.00"
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
        .ColWidth(1) = 2000
        .ColWidth(2) = txtDescription.Width
        .ColWidth(3) = txtQty.Width + 50
        .ColWidth(4) = txtUnitPrice.Width + 100
        .ColWidth(5) = 3500
        
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "PartID"
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "Qty"
        .TextMatrix(0, 4) = "UnitPrice"
        .TextMatrix(0, 5) = "Gross"

        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbLeftJustify
        .ColAlignment(3) = vbLeftJustify
        .ColAlignment(4) = vbLeftJustify
        .ColAlignment(5) = vbLeftJustify
    End With
    
    With Grid2
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
        .ColWidth(4) = txtUnitPrice.Width + 100
        .ColWidth(5) = 4200
        
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "PartID"
        .TextMatrix(0, 2) = "PartDescription"
        .TextMatrix(0, 3) = "Qty"
        .TextMatrix(0, 4) = "UnitPrice"
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
    With NSCustomer
        .ClearColumn
        .AddColumn "CustomerID", 1500
        .AddColumn "Description", 1800
        .AddColumn "Address", 3500
        .AddColumn "LandlineNo", 3500
        
        .Connection = CN.ConnectionString
        .SQLFields = "CustomerID,Description,Address,LandlineNo"
        .sqlTables = "Customers"
        .sqlSortOrder = "CustomerID ASC"
        
        .BoundField = "CustomerID"
        .PageBy = 25
        .DisplayCol = 1
        
        .setDropWindowSize 6800, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Select Customer"
    End With
    
    With NSStatus
        .ClearColumn
        .AddColumn "StatusID", 1200
        .AddColumn "Description", 4500
        
        .Connection = CN.ConnectionString
        .SQLFields = "StatusID,Description"
        .sqlTables = "Sales_Status"
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
        .AddColumn "Price", 1500
        .AddColumn "Inventory", 1500
        
        .Connection = CN.ConnectionString
        .SQLFields = "PartID,PartDescription,Price,Inventory"
        .sqlTables = "Spare_Parts"
        .sqlSortOrder = "PartID ASC"
        
        .BoundField = "PartID"
        .PageBy = 25
        .DisplayCol = 1
        
        .setDropWindowSize 6800, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Select PartID/Spare Part"
    End With
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim RSProduct As New Recordset
    
    If State = AddStateMode Then Exit Sub
    
    RSProduct.CursorLocation = adUseClient
    RSProduct.Open "SELECT * FROM Sales_Order_Details WHERE SalesOrderID=" & PK, CN, adOpenDynamic, adLockOptimistic
    
    If RSProduct.RecordCount > 0 Then
        RSProduct.MoveFirst
        While Not RSProduct.EOF
            CurrRow = getFlexPos(Grid, 1, RSProduct!PartID)
        
            'Add to grid
            With Grid
                If CurrRow < 0 Then
                   DelRecwSQL "Sales_Order_Details", "SalesOrderDetailID", "", True, RSProduct!SalesOrderDetailID
        
                End If
            End With
            RSProduct.MoveNext
        Wend
    End If
    
    Set RSProduct = Nothing
End Sub

Private Sub GeneratePK()
    PK = getIndex("Sales_Order")
End Sub

Private Sub DisplayForEditing()
On Error GoTo err
    
    With RS
        txtEntry(0).Text = PK
        NSCustomer.Text = .Fields("CustomerID")
        txtEntry(4).Text = .Fields("CustomerName")
        txtEntry(1).Text = .Fields("Address")
        txtEntry(2).Text = .Fields("TelNo")
        DTDP.Value = .Fields("Date")
        txtEntry(3).Text = .Fields("Salesman")
        txtEntry(5).Text = .Fields("SInstruction")
        txtRemarks.Text = .Fields("Remarks")
        cmbFP.Text = .Fields("FOP")
        NSStatus.Text = .Fields("StatusID")
        txtEntry(6).Text = .Fields("StatusDesc")
        txtGrossAmount.Text = toMoney(.Fields("Gross"))
        txtDesc.Text = toMoney(.Fields("Discount"))
        txtTaxBase.Text = toMoney(.Fields("TaxBase"))
        txtVat.Text = toMoney(.Fields("VAT"))
        txtNetAmount.Text = toMoney(.Fields("NetAmount"))
        txtCT.Text = toMoney(.Fields("CT"))
        txtCA.Text = toMoney(.Fields("CA"))
        
        cIGross = toMoney(.Fields("Gross"))
        cIAmount = toMoney(.Fields("NetAmount"))
    End With

    cIRowCount = 0


    If RSSales.RecordCount > 0 Then
        RSSales.MoveFirst
        While Not RSSales.EOF
        
          cIRowCount = cIRowCount + 1
          
            With Grid
                If .Rows = 2 And .TextMatrix(1, 5) = "" Then
                    .TextMatrix(1, 1) = RSSales![PartID]
                    .TextMatrix(1, 2) = RSSales![PartDescription]
                    .TextMatrix(1, 3) = toNumber(RSSales![Qty])
                    .TextMatrix(1, 4) = toMoney(RSSales![UnitPrice])
                    .TextMatrix(1, 5) = toMoney(RSSales![GrossAmount])

                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSSales![PartID]
                    .TextMatrix(.Rows - 1, 2) = RSSales![PartDescription]
                    .TextMatrix(.Rows - 1, 3) = toNumber(RSSales![Qty])
                    .TextMatrix(.Rows - 1, 4) = toMoney(RSSales![UnitPrice])
                    .TextMatrix(.Rows - 1, 5) = toMoney(RSSales![GrossAmount])
                    
                    
                End If
            End With
            RSSales.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 5

        If State = EditStateMode Then
            If NSStatus.Text = "CM" Then
                Grid.FixedRows = Grid.Row: Grid.SelectionMode = flexSelectionFree
                Grid.FixedCols = 2
                
            Else
                Grid.FixedRows = Grid.Row:
                Grid.FixedCols = 1

            End If
        End If

    End If

    RSSales.Close
    Set RSSales = Nothing
    
    Exit Sub
    
err:
    If err.Number = 94 Then Resume Next
End Sub

