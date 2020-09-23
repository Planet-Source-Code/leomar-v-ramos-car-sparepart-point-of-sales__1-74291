VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "NSDataCombo.ocx"
Begin VB.Form frmSOSU 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Order Status Update"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10470
   Icon            =   "frmSOSU.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   10260
      Begin VB.TextBox txtStatus 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   1800
      End
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   345
         Left            =   3750
         TabIndex        =   10
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "Update"
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
      Begin ctrlNSDataCombo.NSDataCombo NSStatus 
         Height          =   315
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   900
         _ExtentX        =   1588
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "StatusCD"
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
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.Frame fraShow 
      Caption         =   "Show Record Where?"
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   50
      Width           =   10260
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4080
         TabIndex        =   3
         Top             =   480
         Width           =   1920
      End
      Begin VB.ComboBox cboFields 
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
         ItemData        =   "frmSOSU.frx":08CA
         Left            =   240
         List            =   "frmSOSU.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cboOperator 
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
         ItemData        =   "frmSOSU.frx":08CE
         Left            =   1845
         List            =   "frmSOSU.frx":08E4
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin lvButton.lvButtons_H cmdSearch 
         Height          =   345
         Left            =   6120
         TabIndex        =   7
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "&Search"
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
         Image           =   "frmSOSU.frx":08FD
         cBack           =   -2147483633
      End
      Begin VB.Image picArrow 
         Height          =   255
         Left            =   3720
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fields"
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
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operator"
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
         Left            =   1845
         TabIndex        =   5
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value 1"
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
         Left            =   4080
         TabIndex        =   4
         Top             =   240
         Width           =   650
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4515
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   7964
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmSOSU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RSSalesStatus                   As New Recordset
Dim srcItem                         As ListItem

Dim strSearch                       As Long
Dim jSQL                            As String



Private Sub cmdSearch_Click()
On Error GoTo ErrSearch
    If is_empty(txtFilter(0), False) = True Then txtFilter(0).SetFocus: Exit Sub
    
    If cboOperator.ListIndex <> 6 Then If txtFilter(0).Text = "" Then txtFilter(0).SetFocus: Exit Sub
    
    strSearch = txtFilter(0).Text
       
    jSQL = "SELECT qry_Sales_Order.* " & _
        "FROM qry_Sales_Order " & _
        "WHERE " & cboFields.Text & _
        "" & cboOperator.Text & _
        "" & strSearch
        
    Set RSSalesStatus = New ADODB.Recordset
    If RSSalesStatus.State = adStateOpen Then RSSalesStatus.Close
    RSSalesStatus.Open jSQL, CN, adOpenDynamic, adLockOptimistic

    If RSSalesStatus.RecordCount = 0 Then
        MsgBox "Record(s) not found.Please try again!", vbExclamation
        Exit Sub
    Else
        Call FillListview
        lvSizeColumns lvList
    End If
    
    Exit Sub
ErrSearch:
    MsgBox "Unexpected error occured.Please try another field and search with long data type parameter!", vbExclamation
    Exit Sub
End Sub

Private Sub Form_Activate()
On Error Resume Next
Dim i As Integer

Me.BackColor = MAIN.ACPMenu.BackColor
fraShow.BackColor = MAIN.ACPMenu.BackColor
Frame1.BackColor = MAIN.ACPMenu.BackColor

lvList.FlatScrollBar = True

cboFields.Clear
For i = 1 To lvList.ColumnHeaders.Count
    cboFields.AddItem lvList.ColumnHeaders(i).Text
Next i

cboFields.ListIndex = 0
cboOperator.ListIndex = 0

End Sub

Private Sub Form_Initialize()
FillHeader
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
ElseIf KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
CenterForm frmSOSU
InitializeNSD

On Error GoTo ErrHandler
    
picArrow.Picture = MAIN.i16x16.ListImages(3).Picture
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation
End Sub

Private Sub NSStatus_Change()
txtStatus.Text = NSStatus.getSelValueAt(2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmSOSU = Nothing
Set RSSalesStatus = Nothing
End Sub

Private Sub InitializeNSD()
    With NSStatus
        .ClearColumn
        .AddColumn "StatusID", 1200
        .AddColumn "Description", 5500
        
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
End Sub

Public Sub FillListview()
On Error Resume Next
With lvList
    .View = lvwReport
    .Gridlines = False
    
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "SalesOrderID"
    .ColumnHeaders.Add , , "CustomerID"
    .ColumnHeaders.Add , , "CustomerName"
    .ColumnHeaders.Add , , "Address"
    .ColumnHeaders.Add , , "Date"
    .ColumnHeaders.Add , , "Gross"
    .ColumnHeaders.Add , , "Discount"
    .ColumnHeaders.Add , , "NetAmount"
    .ColumnHeaders.Add , , "StatusID"
    .ColumnHeaders.Add , , "StatusDesc"
    .ColumnHeaders.Add , , "Remarks"
    
    .ListItems.Clear
    Do While Not RSSalesStatus.EOF
    Set srcItem = .ListItems.Add(, , RSSalesStatus.Fields("SalesOrderID"))
        srcItem.SubItems(1) = RSSalesStatus.Fields("CustomerID")
        srcItem.SubItems(2) = RSSalesStatus.Fields("CustomerName")
        srcItem.SubItems(3) = RSSalesStatus.Fields("Address")
        srcItem.SubItems(4) = RSSalesStatus.Fields("Date")
        srcItem.SubItems(5) = toMoney(RSSalesStatus.Fields("Gross"))
        srcItem.SubItems(6) = toNumber(RSSalesStatus.Fields("Discount"))
        srcItem.SubItems(7) = toMoney(RSSalesStatus.Fields("NetAmount"))
        srcItem.SubItems(8) = RSSalesStatus.Fields("StatusID")
        srcItem.SubItems(9) = RSSalesStatus.Fields("StatusDesc")
        srcItem.SubItems(10) = RSSalesStatus.Fields("Remarks")
    RSSalesStatus.MoveNext
    Loop
End With
End Sub

Public Sub FillHeader()
On Error Resume Next
With lvList
    .View = lvwReport
    
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "SalesOrderID"
    .ColumnHeaders.Add , , "CustomerID"
    .ColumnHeaders.Add , , "CustomerName"
    .ColumnHeaders.Add , , "Address"
    .ColumnHeaders.Add , , "Date"
    .ColumnHeaders.Add , , "Gross"
    .ColumnHeaders.Add , , "Discount"
    .ColumnHeaders.Add , , "NetAmount"
    .ColumnHeaders.Add , , "StatusID"
    .ColumnHeaders.Add , , "StatusDesc"
    .ColumnHeaders.Add , , "Remarks"
    
End With
End Sub


