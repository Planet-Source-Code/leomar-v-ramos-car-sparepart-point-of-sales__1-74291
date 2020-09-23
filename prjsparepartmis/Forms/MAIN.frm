VERSION 5.00
Object = "{C9680CB9-8919-4ED0-A47D-8DC07382CB7B}#1.0#0"; "NSStyleButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248850FC-2BAF-48AF-99D6-220E54FE68CA}#1.0#0"; "HookMenu.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm MAIN 
   BackColor       =   &H8000000C&
   Caption         =   "CMOSXP v1.0.1"
   ClientHeight    =   8280
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9555
   Icon            =   "MAIN.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4800
      Top             =   3720
   End
   Begin VB.Timer tmrMemStatus 
      Interval        =   1000
      Left            =   4800
      Top             =   4200
   End
   Begin VB.PictureBox picLeft 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   6120
      Left            =   7245
      ScaleHeight     =   6120
      ScaleWidth      =   2310
      TabIndex        =   8
      Top             =   1740
      Width           =   2310
      Begin VB.Frame fraMenu 
         Height          =   465
         Left            =   0
         TabIndex        =   9
         Top             =   840
         Width           =   2250
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Active Forms"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   600
            TabIndex        =   10
            Top             =   150
            Width           =   1170
         End
      End
      Begin MSComctlLib.ListView lvWin 
         Height          =   4050
         Left            =   0
         TabIndex        =   11
         Top             =   1320
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   7144
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "MAIN.frx":57E2
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Form Name"
            Object.Width           =   3969
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   1950
         Picture         =   "MAIN.frx":64BC
         Top             =   4950
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Image5 
         Height          =   960
         Left            =   1950
         Picture         =   "MAIN.frx":7206
         Top             =   6030
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblVMem 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblPMem 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AVAILABLE FREE MEMORY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   165
         Left            =   100
         TabIndex        =   14
         Top             =   120
         Width           =   2070
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00008000&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   2250
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00008000&
         X1              =   825
         X2              =   825
         Y1              =   345
         Y2              =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Virtual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   165
         Left            =   100
         TabIndex        =   13
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Physical"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   165
         Left            =   100
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   855
         Left            =   0
         Top             =   45
         Width           =   2250
      End
   End
   Begin VB.PictureBox picSeparator 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   6120
      Left            =   7125
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6120
      ScaleWidth      =   120
      TabIndex        =   5
      Top             =   1740
      Width           =   125
      Begin StyleButtonX.StyleButton StyleButton2 
         Height          =   1095
         Left            =   0
         TabIndex        =   6
         Top             =   1920
         Width           =   125
         _ExtentX        =   212
         _ExtentY        =   1931
         UpColorTop1     =   -2147483633
         UpColorTop2     =   -2147483633
         UpColorTop3     =   -2147483633
         UpColorTop4     =   -2147483633
         UpColorButtom1  =   -2147483633
         UpColorButtom2  =   -2147483633
         UpColorButtom3  =   -2147483633
         UpColorButtom4  =   -2147483633
         UpColorLeft1    =   -2147483633
         UpColorLeft2    =   -2147483633
         UpColorLeft3    =   -2147483633
         UpColorLeft4    =   -2147483633
         UpColorRight1   =   -2147483633
         UpColorRight2   =   -2147483633
         UpColorRight3   =   -2147483633
         UpColorRight4   =   -2147483633
         DownColorTop1   =   7021576
         DownColorTop2   =   -2147483633
         DownColorTop3   =   -2147483633
         DownColorTop4   =   -2147483633
         DownColorButtom1=   7021576
         DownColorButtom2=   -2147483633
         DownColorButtom3=   -2147483633
         DownColorButtom4=   -2147483633
         DownColorLeft1  =   7021576
         DownColorLeft2  =   -2147483633
         DownColorLeft3  =   -2147483633
         DownColorLeft4  =   -2147483633
         DownColorRight1 =   7021576
         DownColorRight2 =   -2147483633
         DownColorRight3 =   -2147483633
         DownColorRight4 =   -2147483633
         HoverColorTop1  =   7021576
         HoverColorTop2  =   -2147483633
         HoverColorTop3  =   -2147483633
         HoverColorTop4  =   -2147483633
         HoverColorButtom1=   7021576
         HoverColorButtom2=   -2147483633
         HoverColorButtom3=   -2147483633
         HoverColorButtom4=   -2147483633
         HoverColorLeft1 =   7021576
         HoverColorLeft2 =   -2147483633
         HoverColorLeft3 =   -2147483633
         HoverColorLeft4 =   -2147483633
         HoverColorRight1=   7021576
         HoverColorRight2=   -2147483633
         HoverColorRight3=   -2147483633
         HoverColorRight4=   -2147483633
         FocusColorTop1  =   7021576
         FocusColorTop2  =   -2147483633
         FocusColorTop3  =   -2147483633
         FocusColorTop4  =   -2147483633
         FocusColorButtom1=   7021576
         FocusColorButtom2=   -2147483633
         FocusColorButtom3=   -2147483633
         FocusColorButtom4=   -2147483633
         FocusColorLeft1 =   7021576
         FocusColorLeft2 =   -2147483633
         FocusColorLeft3 =   -2147483633
         FocusColorLeft4 =   -2147483633
         FocusColorRight1=   7021576
         FocusColorRight2=   -2147483633
         FocusColorRight3=   -2147483633
         FocusColorRight4=   -2147483633
         DisabledColorTop1=   -2147483633
         DisabledColorTop2=   -2147483633
         DisabledColorTop3=   -2147483633
         DisabledColorTop4=   -2147483633
         DisabledColorButtom1=   -2147483633
         DisabledColorButtom2=   -2147483633
         DisabledColorButtom3=   -2147483633
         DisabledColorButtom4=   -2147483633
         DisabledColorLeft1=   -2147483633
         DisabledColorLeft2=   -2147483633
         DisabledColorLeft3=   -2147483633
         DisabledColorLeft4=   -2147483633
         DisabledColorRight1=   -2147483633
         DisabledColorRight2=   -2147483633
         DisabledColorRight3=   -2147483633
         DisabledColorRight4=   -2147483633
         Caption         =   ""
         MousePointer    =   1
         BackColorUp     =   -2147483638
         BackColorDown   =   11899524
         BackColorHover  =   14073525
         BackColorFocus  =   14604246
         BackColorDisabled=   -2147483636
         DotsInCornerColor=   16777215
         MoveWhenClick   =   0   'False
         ForeColorUp     =   -2147483630
         ForeColorDown   =   -2147483634
         ForeColorHover  =   -2147483630
         ForeColorFocus  =   -2147483630
         ForeColorDisabled=   12632256
         BeginProperty FontUp {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFocus {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowBorderLevel2=   0   'False
         DistanceBetweenPictureAndCaption=   -50
      End
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   2
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   9555
      TabIndex        =   3
      Top             =   7860
      Width           =   9555
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   10
      Index           =   4
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   9555
      TabIndex        =   2
      Top             =   7875
      Width           =   9555
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   10
      Index           =   5
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   9555
      TabIndex        =   1
      Top             =   7890
      Width           =   9555
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   3
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   8250
      Width           =   9555
   End
   Begin HookMenu.ctxHookMenu HookMenu 
      Left            =   4200
      Top             =   4200
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   8
      Bmp:1           =   "MAIN.frx":7F50
      Key:1           =   "#mnuRACN"
      Bmp:2           =   "MAIN.frx":8378
      Mask:2          =   16777215
      Key:2           =   "#mnuRAES"
      Bmp:3           =   "MAIN.frx":86CA
      Key:3           =   "#mnuRAP"
      Bmp:4           =   "MAIN.frx":8AF2
      Mask:4          =   16777215
      Key:4           =   "#mnuRADS"
      Bmp:5           =   "MAIN.frx":8E44
      Key:5           =   "#mnuRARR"
      Bmp:6           =   "MAIN.frx":926C
      Mask:6          =   16777215
      Key:6           =   "#mnuRAC"
      Bmp:7           =   "MAIN.frx":95BE
      Key:7           =   "#mUCM"
      Bmp:8           =   "MAIN.frx":99E6
      Key:8           =   "#mCM"
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
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   7905
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "MAIN.frx":9E0E
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "ACTIVE USER:"
            TextSave        =   "ACTIVE USER:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "MAIN.frx":A1AA
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Today:"
            TextSave        =   "Today:"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "1/16/2011"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "9:29 PM"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
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
   Begin prjcmosxp.ACPRibbon ACPMenu 
      Align           =   1  'Align Top
      Height          =   1740
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   3069
      BackColor       =   4210752
      ForeColor       =   -2147483630
   End
   Begin MSComDlg.CommonDialog CDExporter 
      Left            =   4200
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   4800
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":A544
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":A6D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":D558
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i32x32 
      Left            =   1560
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":DF6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":E5D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":FF64
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":10576
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":10AC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":11089
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":116CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":11CD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":231EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":23752
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":23E05
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":246DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":24CE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":252E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":25926
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":25F85
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":26515
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":26B0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":271CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2770C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":27D3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":283D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":28A86
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":290A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":296C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":29C3A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Visible         =   0   'False
      Begin VB.Menu mnuRACN 
         Caption         =   "Create New Entry"
      End
      Begin VB.Menu mnuRAES 
         Caption         =   "Edit Selected"
      End
      Begin VB.Menu mnuRADS 
         Caption         =   "Delete Selected"
      End
      Begin VB.Menu mnuRARR 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuRAP 
         Caption         =   "Export"
      End
      Begin VB.Menu mnuRAC 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mReport 
      Caption         =   "Reports"
      Visible         =   0   'False
      Begin VB.Menu mAU 
         Caption         =   "All User"
      End
      Begin VB.Menu mAUT 
         Caption         =   "All User Type"
      End
      Begin VB.Menu mAC 
         Caption         =   "All Customer"
      End
      Begin VB.Menu mAS 
         Caption         =   "All Supplier"
      End
      Begin VB.Menu mACM 
         Caption         =   "All Car Make"
      End
      Begin VB.Menu mACModel 
         Caption         =   "All Car Model"
      End
      Begin VB.Menu mAPC 
         Caption         =   "All Part Category"
      End
      Begin VB.Menu mASP 
         Caption         =   "All Spare Part"
      End
   End
   Begin VB.Menu mUpload 
      Caption         =   "Upload"
      Visible         =   0   'False
      Begin VB.Menu mUCM 
         Caption         =   "Upload Car Type..."
      End
      Begin VB.Menu mCM 
         Caption         =   "Delete Car Type..."
      End
   End
End
Attribute VB_Name = "MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim cursor_pos As POINTAPI

Dim resize_down     As Boolean
Dim show_mnu        As Boolean
Dim pos_num         As Integer
Dim Theme           As Integer

Public CloseMe      As Boolean


Private Sub ACPMenu_ButtonClick(ByVal ID As String, ByVal Caption As String)
If ID = "User Accounts" Then
    If ACTIVE_USER.USER_ISADMIN = False Then
        MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
    Else
        LoadForm frmUser
    End If
ElseIf ID = "User Types" Then
    If ACTIVE_USER.USER_ISADMIN = False Then
        MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
    Else
        LoadForm frmUserType
    End If
ElseIf ID = "Suppliers" Then
    LoadForm frmSupplier
ElseIf ID = "Customers" Then
    LoadForm frmCustomer
ElseIf ID = "Zipcodes" Then
    LoadForm frmZipcode
ElseIf ID = "Car Makes" Then
    If ACTIVE_USER.USER_ISADMIN = False Then
        MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
    Else
        LoadForm frmCarMake
    End If
ElseIf ID = "Car Types" Then
    If ACTIVE_USER.USER_ISADMIN = False Then
        MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
    Else
        LoadForm frmCarType
    End If
ElseIf ID = "Part Categories" Then
    If ACTIVE_USER.USER_ISADMIN = False Then
        MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
    Else
        LoadForm frmPCategory
    End If
ElseIf ID = "Spare Parts" Then
    If ACTIVE_USER.USER_ISADMIN = False Then
        MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
    Else
        LoadForm frmSparepart
    End If
    
'TRANSACTIONS
ElseIf ID = "Sales Entries" Then
    LoadForm frmSales
ElseIf ID = "Purchase Entries" Then
    LoadForm frmPurchase
    
'REPORTS
ElseIf ID = "All User" Then
    mAU_Click
ElseIf ID = "All User Type" Then
    mAUT_Click
ElseIf ID = "All Customer" Then
    mAC_Click
ElseIf ID = "All Supplier" Then
    mAS_Click
ElseIf ID = "All Car Brand" Then
    mACM_Click
ElseIf ID = "All Car Types" Then
    mACModel_Click
ElseIf ID = "All Part Category" Then
    mAPC_Click
ElseIf ID = "All Spare Part" Then
    mASP_Click
ElseIf ID = "Sales Per Part By Date" Then
    frmSalesPerPartByDate.show vbModal
ElseIf ID = "Sales Total Per Customer" Then
    frmSalesTotalPerCustomer.show vbModal

'SETTINGS
ElseIf ID = "Theme Option" Then
    Theme = Theme + 1
    
    If Theme = 3 Then Theme = 0
    
    ACPMenu.Theme = Theme
    ACPMenu.Refresh
    
    MAIN.Picture = ACPMenu.LoadBackground
    MAIN.BackColor = ACPMenu.BackColor
    
    picLeft.BackColor = MAIN.ACPMenu.BackColor
    picSeparator.BackColor = MAIN.ACPMenu.BackColor
    picLine(4).BackColor = MAIN.ACPMenu.BackColor
    fraMenu.BackColor = MAIN.ACPMenu.BackColor
    StyleButton2.BackColorFocus = MAIN.ACPMenu.BackColor
    StyleButton2.BackColorUp = MAIN.ACPMenu.BackColor
    
ElseIf ID = "Change Pass" Then
    frmChangePassword.show vbModal
ElseIf ID = "Change Pass" Then
    frmChangePassword.show vbModal
ElseIf ID = "Business Information" Then
    If ACTIVE_USER.USER_ISADMIN = False Then
        MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
    Else
        frmBusiness.show vbModal
    End If
    
'UTILITIES
ElseIf ID = "Lock" Then
    frmLock.show vbModal
ElseIf ID = "Calculator" Then
    Shell "calc.exe", vbNormalFocus
ElseIf ID = "Notepad" Then
    Shell "notepad.exe", vbNormalFocus
ElseIf ID = "Windows Explorer" Then
    Shell "Explorer.exe", vbNormalFocus
ElseIf ID = "Date/Time Setting" Then
    Shell "control.exe date/time", vbNormalFocus
ElseIf ID = "Sales Order Status Update" Then
    If ACTIVE_USER.USER_ISADMIN = False Then
        MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
    Else
        frmSOSU.show vbModal
    End If
ElseIf ID = "Purchase Order Status Update" Then
    If ACTIVE_USER.USER_ISADMIN = False Then
        MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
    Else
        frmPOSU.show vbModal
    End If
    
'EXIT APPLICATION
ElseIf ID = "Shutdown" Then
    If MsgBox("Are you sure you want to shutdown the system?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    UnloadChilds
    
    StatusBar.Panels(3).Text = vbNullString
    StatusBar.Panels(4).Text = vbNullString
    
    ACTIVE_USER.USERNAME = vbNullString
    ACTIVE_USER.USERTYPE = vbNullString
    
    Unload Me
ElseIf ID = "Logoff" Then
    If MsgBox("Are you sure you want to log off?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    UnloadChilds
    
    StatusBar.Panels(3).Text = vbNullString
    StatusBar.Panels(4).Text = vbNullString
    
    ACTIVE_USER.USERNAME = vbNullString
    ACTIVE_USER.USERTYPE = vbNullString
    
    Unload Me
    frmLogin.show vbModal
ElseIf ID = "Switch" Then
    If MsgBox("Please log out first before switching to another user.Proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    UnloadChilds
    
    StatusBar.Panels(3).Text = vbNullString
    StatusBar.Panels(4).Text = vbNullString
    
    ACTIVE_USER.USERNAME = vbNullString
    ACTIVE_USER.USERTYPE = vbNullString
    
    'Unload Me
    frmLogin.show vbModal


End If
End Sub

Private Sub ACPMenu_TabClick(ByVal ID As String, ByVal Caption As String)
Call UnloadChilds
End Sub

Private Sub lvWin_Click()
If lvWin.ListItems.Count < 1 Then Exit Sub
    
    Select Case lvWin.SelectedItem.Key

        Case "frmSupplier": LoadForm frmSupplier
        Case "frmCustomer": LoadForm frmCustomer
        Case "frmZipcode": LoadForm frmZipcode
        Case "frmUser": LoadForm frmUser
        Case "frmUserType": LoadForm frmUserType
        Case "frmCarType": LoadForm frmCarType
        Case "frmCarMake": LoadForm frmCarMake
        Case "frmPCategory": LoadForm frmPCategory
        
    End Select

End Sub

Private Sub mAC_Click()
    Dim qSQL As String
    qSQL = "SELECT Customers.* FROM Customers ORDER BY CustomerID ASC"
    
    With rptAllCustomer
        .rptCN.Connect = ";DATABASE=" & App.Path & "\Database\CMOSXP_DB.mdb;PWD=qwerty123;"
        .rptCN.DatabaseName = App.Path & "\Database\CMOSXP_DB.mdb"
        .rptCN.RecordSource = qSQL
        
        .lblDate.Caption = Now
        .lblCompany.Caption = ACTIVE_COMPANY.COMPANYNAME
        .lblAddress.Caption = ACTIVE_COMPANY.ADDRESS
        
        .txtCustomerID.DataField = "CustomerID"
        .txtCompany.DataField = "Description"
        .txtAddress.DataField = "Address"
        .txtOwnerName.DataField = "OwnerName"
        .txtContactNo.DataField = "LandlineNo"
        .show
    End With
End Sub

Private Sub mACM_Click()
    Dim iSQL As String
    iSQL = "SELECT Car_Makes.* FROM Car_Makes ORDER BY MakeID ASC"
    
    With rptAllCarMake
        .rptCN.Connect = ";DATABASE=" & App.Path & "\Database\CMOSXP_DB.mdb;PWD=qwerty123;"
        .rptCN.DatabaseName = App.Path & "\Database\CMOSXP_DB.mdb"
        .rptCN.RecordSource = iSQL
        
        .lblDate.Caption = Now
        .lblCompany.Caption = ACTIVE_COMPANY.COMPANYNAME
        .lblAddress.Caption = ACTIVE_COMPANY.ADDRESS
        
        .txtMakeID.DataField = "MakeID"
        .txtMakeName.DataField = "MakeName"
        .txtRemarks.DataField = "Remarks"
        .Photo.DataField = "picFile"
        
        .show
    End With
End Sub

Private Sub mACModel_Click()
    Dim zSQL As String
    zSQL = "SELECT Car_Types.* FROM Car_Types ORDER BY CarTypeID ASC"
    
    With rptAllCarModel
        .rptCN.Connect = ";DATABASE=" & App.Path & "\Database\CMOSXP_DB.mdb;PWD=qwerty123;"
        .rptCN.DatabaseName = App.Path & "\Database\CMOSXP_DB.mdb"
        .rptCN.RecordSource = zSQL
        
        .lblDate.Caption = Now
        .lblCompany.Caption = ACTIVE_COMPANY.COMPANYNAME
        .lblAddress.Caption = ACTIVE_COMPANY.ADDRESS
        
        .txtModelID.DataField = "CarTypeID"
        .txtModelName.DataField = "CarTypeName"
        .txtMakeID.DataField = "MakeID"
        .txtRemarks.DataField = "Remarks"
        .Photo.DataField = "picFile"
        
        .show
    End With

End Sub

Private Sub mAPC_Click()
Dim kSQL As String
    kSQL = "SELECT Part_Categories.* FROM Part_Categories ORDER BY PCategoryID ASC"
    
    With rptAllPartCategory
        .rptCN.Connect = ";DATABASE=" & App.Path & "\Database\CMOSXP_DB.mdb;PWD=qwerty123;"
        .rptCN.DatabaseName = App.Path & "\Database\CMOSXP_DB.mdb"
        .rptCN.RecordSource = kSQL
        
        .lblDate.Caption = Now
        .lblCompany.Caption = ACTIVE_COMPANY.COMPANYNAME
        .lblAddress.Caption = ACTIVE_COMPANY.ADDRESS
        
        .txtPartID.DataField = "PCategoryID"
        .txtDescription.DataField = "PCategoryName"
        .txtRemarks.DataField = "Remarks"
        .txtDE.DataField = "DateEncoded"
        .txtEB.DataField = "EncodedBy"
        
        .show
    End With
End Sub

Private Sub mAS_Click()
    Dim tSQL As String
    tSQL = "SELECT Suppliers.* FROM Suppliers ORDER BY SupplierID ASC"
    
    With rptAllSupplier
        .rptCN.Connect = ";DATABASE=" & App.Path & "\Database\CMOSXP_DB.mdb;PWD=qwerty123;"
        .rptCN.DatabaseName = App.Path & "\Database\CMOSXP_DB.mdb"
        .rptCN.RecordSource = tSQL
        
        .lblDate.Caption = Now
        .lblCompany.Caption = ACTIVE_COMPANY.COMPANYNAME
        .lblAddress.Caption = ACTIVE_COMPANY.ADDRESS
        
        .txtSupplierID.DataField = "SupplierID"
        .txtDescription.DataField = "Description"
        .txtAddress.DataField = "Address"
        .txtContactPerson.DataField = "ContactPerson"
        .txtBusinessNo.DataField = "BusinessNo"

        .show
    End With
End Sub

Private Sub mASP_Click()
    Dim mSQL As String
    mSQL = "SELECT Spare_Parts.* FROM Spare_Parts ORDER BY PartID ASC"
    
    With rptAllSparepart
        .rptCN.Connect = ";DATABASE=" & App.Path & "\Database\CMOSXP_DB.mdb;PWD=qwerty123;"
        .rptCN.DatabaseName = App.Path & "\Database\CMOSXP_DB.mdb"
        .rptCN.RecordSource = mSQL
        
        .lblDate.Caption = Now
        .lblCompany.Caption = ACTIVE_COMPANY.COMPANYNAME
        .lblAddress.Caption = ACTIVE_COMPANY.ADDRESS
        
        .txtPartID.DataField = "PartID"
        .txtPartNo.DataField = "PartNumber"
        .txtDescription.DataField = "PartDescription"
        .txtPrice.DataField = "Price"
        .txtCarMake.DataField = "MakeName"
        .txtCarType.DataField = "CarTypeName"
        .txtCategory.DataField = "PCategoryName"
        
        .txtSP.DataField = "SupplierPrice"
        
        .txtYear.DataField = "Year"
        .txtInventory.DataField = "Inventory"
        
        .txtCapacity.DataField = "Capacity"
        .txtGearbox.DataField = "Gearbox"
        .Photo.DataField = "Photo"
        
        .show
    End With
End Sub

Private Sub mAU_Click()
    Dim sSQL As String
    sSQL = "SELECT Users.* FROM Users ORDER BY UserID ASC"
    
    With rptAllUser
        .rptCN.Connect = ";DATABASE=" & App.Path & "\Database\CMOSXP_DB.mdb;PWD=qwerty123;"
        .rptCN.DatabaseName = App.Path & "\Database\CMOSXP_DB.mdb"
        .rptCN.RecordSource = sSQL
        
        .lblDate.Caption = Now
        .lblCompany.Caption = ACTIVE_COMPANY.COMPANYNAME
        .lblAddress.Caption = ACTIVE_COMPANY.ADDRESS
        
        .txtUserID.DataField = "UserID"
        .txtFullname.DataField = "Fullname"
        .txtUserType.DataField = "UserType"
        .txtLoginName.DataField = "Username"
        .txtStatus.DataField = "StatusCD"
        .txtRemarks.DataField = "Remarks"
        
        .show
    End With
End Sub

Private Sub mAUT_Click()
    Dim jSQL As String
    jSQL = "SELECT User_Types.* FROM User_Types ORDER BY UserTypeID ASC"
    
    With rptAllUserType
        .rptCN.Connect = ";DATABASE=" & App.Path & "\Database\CMOSXP_DB.mdb;PWD=qwerty123;"
        .rptCN.DatabaseName = App.Path & "\Database\CMOSXP_DB.mdb"
        .rptCN.RecordSource = jSQL
        
        .lblDate.Caption = Now
        .lblCompany.Caption = ACTIVE_COMPANY.COMPANYNAME
        .lblAddress.Caption = ACTIVE_COMPANY.ADDRESS
        
        .txtUserTypeID.DataField = "UserTypeID"
        .txtDescription.DataField = "UserType"
        .txtRemarks.DataField = "Remarks"
        .txtDE.DataField = "DateEncoded"
        .txtEB.DataField = "EncodedBy"
        
        .show
    End With
End Sub

Private Sub mCM_Click()
On Error Resume Next
frmCarTypeAE.CommandPass "Delete Car Type"
End Sub

Private Sub MDIForm_Activate()
On Error Resume Next
    If END_APP = True Then End: Exit Sub
    
    picLeft.BackColor = MAIN.ACPMenu.BackColor
    picSeparator.BackColor = MAIN.ACPMenu.BackColor
    picLine(4).BackColor = MAIN.ACPMenu.BackColor
    fraMenu.BackColor = MAIN.ACPMenu.BackColor
    StyleButton2.BackColorFocus = MAIN.ACPMenu.BackColor
    StyleButton2.BackColorUp = MAIN.ACPMenu.BackColor
    lvWin.FlatScrollBar = True
End Sub

Private Sub MDIForm_Initialize()
    On Error Resume Next
        ' this will fail if Comctl not available
        '  - unlikely now though!
        Dim iccex As tagInitCommonControlsEx
        With iccex
            .lngSize = LenB(iccex)
            .lngICC = ICC_USEREX_CLASSES
        End With
        InitCommonControlsEx iccex
End Sub

Private Sub MDIForm_Load()
On Error GoTo ErrHandler

    Set lvWin.SmallIcons = i16x16
    Set lvWin.Icons = i16x16
    
    show_mnu = True
    show_menu (show_mnu)
    
    Theme = 2
    
    ACPMenu.Theme = Theme
    
    MAIN.Picture = ACPMenu.LoadBackground
    MAIN.BackColor = ACPMenu.BackColor
    
    
    ACPMenu.ImageList = i32x32
    ACPMenu.ButtonCenter = False
    
    ''''TABS
    ACPMenu.AddTab "1", "Masters"
    ACPMenu.AddTab "2", "Transactions"
    ACPMenu.AddTab "3", "Reports"
    ACPMenu.AddTab "4", "Utilities"
    ACPMenu.AddTab "5", "Settings"
    ACPMenu.AddTab "6", "Exit Application"
    
    ''''CATEGORIES
    ACPMenu.AddCat "Master Files", "1", "Master Files", True
    ACPMenu.AddCat "Miscellaneous", "1", "Miscellaneous", True
    ACPMenu.AddCat "System User", "1", "System User", True
    
    ACPMenu.AddCat "Transactions", "2", "Transactions", True
    ACPMenu.AddCat "Status Update Utility", "2", "Status Update Utility", True
    
    ACPMenu.AddCat "Masterfile Reports", "3", "Masterfile Reports", True
    ACPMenu.AddCat "Car Brands and Types", "3", "Car Brands and Types", True
    ACPMenu.AddCat "Sales Reports", "3", "Sales Reports", True
    ACPMenu.AddCat "Locker", "4", "Locker", True
    ACPMenu.AddCat "Utilities", "4", "Utilities", True
    ACPMenu.AddCat "Settings", "5", "Settings", False
    ACPMenu.AddCat "Preferences", "5", "Preferences", False
    ACPMenu.AddCat "Exit Application", "6", "Exit Application", True
    
    'For Masterfiles
    ACPMenu.AddButton "Car Makes", "Master Files", "Car Makes", 13
    ACPMenu.AddButton "Car Types", "Master Files", "Car Types", 12
    ACPMenu.AddButton "Part Categories", "Master Files", "Part" & vbNewLine & " Categories", 14
    ACPMenu.AddButton "Spare Parts", "Master Files", "Spare Parts", 15
    
    ACPMenu.AddButton "Customers", "Miscellaneous", "Customers", 17
    ACPMenu.AddButton "Suppliers", "Miscellaneous", "Suppliers", 16
    ACPMenu.AddButton "Zipcodes", "Miscellaneous", "Zipcodes", 18
    
    ACPMenu.AddButton "User Accounts", "System User", "User Accounts", 7
    ACPMenu.AddButton "User Types", "System User", "User Types", 22
    
    'For Transactions
    ACPMenu.AddButton "Sales Entries", "Transactions", "Sales" & vbNewLine & " Entries", 23
    ACPMenu.AddButton "Purchase Entries", "Transactions", "Purchase " & vbNewLine & "Entries", 24
    
    ACPMenu.AddButton "Sales Order Status Update", "Status Update Utility", "Sales  " & vbNewLine & "Order Update", 25
    ACPMenu.AddButton "Purchase Order Status Update", "Status Update Utility", "Purchase  " & vbNewLine & "Order Update", 26

    'For Reports
    ACPMenu.AddButton "All User", "Masterfile Reports", "User" & vbNewLine & "  Masterlist", 11
    ACPMenu.AddButton "All User Type", "Masterfile Reports", "User Type" & vbNewLine & " Masterlist ", 11
    ACPMenu.AddButton "All Customer", "Masterfile Reports", "Customer" & vbNewLine & "Masterlist", 11
    ACPMenu.AddButton "All Supplier", "Masterfile Reports", "Supplier" & vbNewLine & "Masterlist", 11
    
    ACPMenu.AddButton "All Car Brand", "Car Brands and Types", "Car" & vbNewLine & "Makes", 11
    ACPMenu.AddButton "All Car Types", "Car Brands and Types", "Car" & vbNewLine & "Types", 11
    ACPMenu.AddButton "All Part Category", "Car Brands and Types", "Part" & vbNewLine & "Category", 11
    ACPMenu.AddButton "All Spare Part", "Car Brands and Types", "Spare" & vbNewLine & "Parts", 11
    
    ACPMenu.AddButton "Sales Per Part By Date", "Sales Reports", "Sales Per" & vbNewLine & "Part By Date", 11
    ACPMenu.AddButton "Sales Total Per Customer", "Sales Reports", "Sales Total" & vbNewLine & "Per Customer", 11
    
    
    'For Utilities
    ACPMenu.AddButton "Lock", "Locker", "Lock", 6
    ACPMenu.AddButton "Calculator", "Utilities", "Calculator", 3
    ACPMenu.AddButton "Notepad", "Utilities", "Notepad", 4
    ACPMenu.AddButton "Windows Explorer", "Utilities", "Windows " & vbNewLine & " Explorer", 5
        
    'For Settings
    ACPMenu.AddButton "Theme Option", "Settings", "Setup" & vbNewLine & "Theme", 19
    ACPMenu.AddButton "Date/Time Setting", "Settings", "Date/Time" & vbNewLine & " Setting", 1
    ACPMenu.AddButton "Change Pass", "Preferences", "Change" & vbNewLine & " Password", 20
    ACPMenu.AddButton "Business Information", "Preferences", "Business" & vbNewLine & "Information", 21
    
    ACPMenu.AddButton "Switch", "Exit Application", "Switch" & vbNewLine & "User", 8
    ACPMenu.AddButton "Logoff", "Exit Application", "Logoff", 9
    ACPMenu.AddButton "Shutdown", "Exit Application", "Shutdown", 10

    ACPMenu.Refresh
    
    If Connected2DB = False Then Unload Me: Exit Sub
    DisplayBusinessInfo
    
    
    Me.show
    frmLogin.show vbModal

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
    fraMenu.Width = lvWin.Width
    
    fraMenu.Left = Shape1.Left
    lvWin.Left = Shape1.Left
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Call UnloadChilds
    Set MAIN = Nothing
    Set CN = Nothing
    End
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Reply As Integer

Reply = MsgBox("This will terminate the application.Do you want to proceed?", vbExclamation + vbYesNo)

If Reply = vbNo Then
    Cancel = 1
End If
End Sub

Private Sub mnuRACN_Click()
    On Error Resume Next
    ActiveForm.CommandPass "New"
End Sub

Private Sub mnuRADS_Click()
    On Error Resume Next
    ActiveForm.CommandPass "Delete"
End Sub

Private Sub mnuRAES_Click()
    On Error Resume Next
    ActiveForm.CommandPass "Update"
End Sub

Private Sub mnuRAP_Click()
    On Error Resume Next
    ActiveForm.CommandPass "Export"
End Sub

Private Sub mnuRARR_Click()
    On Error Resume Next
    ActiveForm.CommandPass "Refresh"
End Sub

Private Sub mnuRAC_Click()
    On Error Resume Next
    ActiveForm.CommandPass "Close"
End Sub

Private Sub mUCM_Click()
On Error Resume Next
frmCarTypeAE.CommandPass "Upload"
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    lvWin.Width = picLeft.ScaleWidth
    lvWin.Height = picLeft.ScaleHeight - lvWin.Top - 20
End Sub

Private Sub picSeparator_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If show_mnu = False Then Exit Sub
    If Button = vbLeftButton Then
        tmrResize.Enabled = True
        resize_down = True
    End If
End Sub

Private Sub picSeparator_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If show_mnu = False Then Exit Sub
    If Button = vbLeftButton Then
        tmrResize.Enabled = False
        resize_down = False
    End If
End Sub

Private Sub picSeparator_Resize()
    Call center_obj_vertical(picSeparator, StyleButton2)
End Sub

Private Sub StyleButton2_Click()
    show_mnu = Not show_mnu
    show_menu (show_mnu)
End Sub


Private Sub StyleButton2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    picLeft.BackColor = MAIN.ACPMenu.BackColor
    picSeparator.BackColor = MAIN.ACPMenu.BackColor
    picLine(4).BackColor = MAIN.ACPMenu.BackColor
    fraMenu.BackColor = MAIN.ACPMenu.BackColor
    StyleButton2.BackColorFocus = MAIN.ACPMenu.BackColor
    StyleButton2.BackColorUp = MAIN.ACPMenu.BackColor
End Sub

Private Sub tmrMemStatus_Timer()
    Call GlobalMemoryStatus(MEM_STAT)
    lblPMem.Caption = Format((MEM_STAT.dwAvailPhys / 1024) / 1024, "#,##0.0") & " MB"
    lblVMem.Caption = Format((MEM_STAT.dwAvailVirtual / 1024) / 1024, "#,##0.0") & " MB"

End Sub

Private Sub tmrResize_Timer()
    On Error Resume Next
    GetCursorPos cursor_pos
    picLeft.Width = (Me.Width - ((cursor_pos.x * Screen.TwipsPerPixelX) - Me.Left)) - 90
End Sub

Public Sub UnloadChilds()
''Unload all active forms
Dim Form As Form
   For Each Form In Forms
      ''Unload all active childs
      If Form.Name <> Me.Name And Form.Name <> "frmShortcuts" Then Unload Form
   Next Form
   
Set Form = Nothing
End Sub

Private Sub show_menu(ByVal show As Boolean)
    Dim img As Image
    If show = True Then
        Set img = Image2
    Else
        Set img = Image5
    End If
    'Set the style button graphics
    With StyleButton2
        Set .PictureDown = img.Picture
        Set .PictureFocus = img.Picture
        Set .PictureHover = img.Picture
        Set .PictureUp = img.Picture
    End With
    'Set picture visibility
    picLeft.Visible = show
    
    If show = True Then StyleButton2.ToolTipText = "Hide": picSeparator.MousePointer = vbSizeWE Else picSeparator.MousePointer = vbArrow: StyleButton2.ToolTipText = "Expand"
    
    Set img = Nothing
End Sub
Private Sub DisplayBusinessInfo()
On Error Resume Next
Set RS_COMPANY = New ADODB.Recordset

RS_COMPANY.CursorLocation = adUseClient
RS_COMPANY.Open "SELECT * FROM Company_Info", CN, adOpenStatic, adLockReadOnly

With ACTIVE_COMPANY
    .COMPANYID = RS_COMPANY.Fields("CompanyID")
    .COMPANYNAME = RS_COMPANY.Fields("CompanyName")
    .ADDRESS = RS_COMPANY.Fields("Address")
    .BUSINESSNO = RS_COMPANY.Fields("BusinessNo")
    .EMAIL = RS_COMPANY.Fields("Email")
    .FAXNO = RS_COMPANY.Fields("FaxNo")
End With

End Sub

Public Sub AddToWin(ByVal srcDName As String, ByVal srcFormName As String)
    On Error Resume Next
    Dim xItem As ListItem
    
    Set xItem = lvWin.ListItems.Add(, srcFormName, srcDName, 1, 1)
    xItem.ToolTipText = srcDName
    xItem.SubItems(1) = "***" & srcDName & "***"
    xItem.Selected = True
    
    Set xItem = Nothing
End Sub

Public Sub RemToWin(ByVal srcDName As String)
    On Error Resume Next
    search_in_listview lvWin, "***" & srcDName & "***"
    lvWin.ListItems.Remove (lvWin.SelectedItem.Index)
End Sub
