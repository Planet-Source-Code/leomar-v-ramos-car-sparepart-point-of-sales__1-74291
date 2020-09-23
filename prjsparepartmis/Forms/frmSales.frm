VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls.ocx"
Begin VB.Form frmSales 
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8370
   ScaleWidth      =   10095
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   10095
      TabIndex        =   6
      Top             =   7965
      Width           =   10095
      Begin prjcmosxp.Liner Liner1 
         Height          =   30
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   53
      End
      Begin VB.Label lblRecSum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Records"
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
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   690
      End
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3750
      TabIndex        =   1
      Top             =   600
      Width           =   2775
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
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   5175
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Right click here to see more option..."
      Top             =   1080
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9128
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
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
   Begin b8Controls4.b8TitleBar b8TB 
      Height          =   375
      Left            =   45
      TabIndex        =   3
      Top             =   50
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   661
      Caption         =   "Manage Sales Entries"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ShadowColor     =   0
      BorderColor     =   4210752
      BackColor       =   8421504
   End
   Begin lvButton.lvButtons_H cmdSearch 
      Height          =   345
      Left            =   6630
      TabIndex        =   4
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
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
      Image           =   "frmSales.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   405
      Left            =   50
      TabIndex        =   9
      Top             =   1515
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "&Update"
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
      cFHover         =   4210752
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmSales.frx":077A
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdNew 
      Height          =   405
      Left            =   50
      TabIndex        =   10
      Top             =   1080
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "&New"
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
      cFHover         =   4210752
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmSales.frx":0909
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   405
      Left            =   50
      TabIndex        =   11
      Top             =   1950
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "&Delete"
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
      cFHover         =   4210752
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmSales.frx":0A63
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   405
      Left            =   50
      TabIndex        =   12
      Top             =   2400
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "&Refresh"
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
      cFHover         =   4210752
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmSales.frx":3D00
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   405
      Left            =   50
      TabIndex        =   13
      Top             =   4440
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "&Close"
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
      cFHover         =   4210752
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmSales.frx":447A
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdExport 
      Height          =   405
      Left            =   50
      TabIndex        =   14
      Top             =   4005
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "&Export"
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
      cFHover         =   4210752
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmSales.frx":76DC
      cBack           =   16119285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Search:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   45
      TabIndex        =   5
      Top             =   600
      Width           =   1140
   End
   Begin VB.Image picArrow 
      Height          =   255
      Left            =   3435
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim srcItem                        As ListItem
Dim srcRecord                      As String
Dim srcSales                       As Variant
Dim srcSQL                         As String

Private Sub cmdNew_Click()
CommandPass "New"
End Sub


Private Sub cmdUpdate_Click()
On Error Resume Next
    CommandPass "Update"
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
    CommandPass "Delete"
End Sub

Private Sub cmdRefresh_Click()
On Error Resume Next
    CommandPass "Refresh"
End Sub

Private Sub cmdExport_Click()
On Error Resume Next
    CommandPass "Export"
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
    CommandPass "Close"
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next

srcSQL = "SELECT qry_Sales_Order.* " & _
        "FROM qry_Sales_Order " & _
        "WHERE (((" & cboFields.Text & ") Like '%" & txtSearch.Text & "%'))"

Set RS_SALES = New ADODB.Recordset
If RS_SALES.State = adStateOpen Then RS_SALES.Close
RS_SALES.Open srcSQL, CN, adOpenDynamic, adLockOptimistic

If txtSearch.Text = vbNullString Then Exit Sub:

If RS_SALES.RecordCount < 1 Then
    MsgBox "No record(s) found in the list!", vbExclamation
    txtSearch.SetFocus
    Exit Sub
Else
    Call FillListview
    Call lvSizeColumns(lvList)
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
Dim i As Integer
HighlightInWin Name

With MAIN
    Me.BackColor = .ACPMenu.BackColor
    Me.Picture = .ACPMenu.LoadBackground
    picFooter.BackColor = .ACPMenu.BackColor
End With

lvList.FlatScrollBar = True

cboFields.Clear
For i = 1 To lvList.ColumnHeaders.Count
    cboFields.AddItem lvList.ColumnHeaders(i).Text
Next i

cboFields.ListIndex = 0

End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
MAIN.AddToWin "Sales Entries", Name


picArrow.Picture = MAIN.i16x16.ListImages(3).Picture

Set lvList.SmallIcons = MAIN.i16x16
Set lvList.Icons = MAIN.i16x16

srcSQL = "SELECT qry_Sales_Order.* " & _
            " FROM qry_Sales_Order " & _
            " ORDER BY qry_Sales_Order.SalesOrderID ASC "

Set RS_SALES = New ADODB.Recordset
If RS_SALES.State = adStateOpen Then RS_SALES.Close
RS_SALES.Open srcSQL, CN, adOpenDynamic, adLockPessimistic

srcSales = "NONE"
srcRecord = vbNullString

Call FillListview
Call lvSizeColumns(lvList)
Call RefreshRecSum

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        
        b8TB.Width = ScaleWidth
        Liner1.Width = ScaleWidth
        lvList.Width = Me.ScaleWidth - 1130
        lvList.Height = Me.ScaleHeight - 1600
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
MAIN.RemToWin Me.Caption

Set frmSales = Nothing
Set RS_SALES = Nothing
End Sub

Private Sub lvList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lvList.Sorted And _
        ColumnHeader.Index - 1 = lvList.SortKey Then
        lvList.SortOrder = 1 - lvList.SortOrder
    Else
        lvList.SortOrder = lvwAscending
        lvList.SortKey = ColumnHeader.Index - 1
    End If
    lvList.Sorted = True
End Sub

Private Sub lvList_DblClick()
On Error Resume Next
CommandPass "Update"
End Sub

Private Sub lvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
srcSales = lvList.SelectedItem.Index
srcRecord = lvList.ListItems.Item(srcSales).Text
Call RefreshRecSum
End Sub

Private Sub lvList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu MAIN.mnuAction
End Sub

Private Sub txtSearch_GotFocus()
HLText txtSearch
End Sub


Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSearch_Click
End If
End Sub

Public Sub CommandPass(ByVal srcPerformWhat As String)
On Error GoTo errPerformWhat
Select Case srcPerformWhat
    Case "New" 'New
            With frmSalesAE
                .State = AddStateMode
                .show vbModal
            End With
    Case "Update" 'Update
            If srcRecord = vbNullString Then
                MsgBox "Invalid selection.Can't proceed to this operation!", vbExclamation
                Exit Sub
            Else
                With frmSalesAE
                    .State = EditStateMode
                    .PK = srcRecord
                    .show vbModal
                End With
            End If
            
    Case "Delete" 'Delete
            If lvList.ListItems.Count < 1 Then
            MsgBox "There's no record to delete!", vbExclamation
            Exit Sub
            End If
            
            If srcRecord = vbNullString Then
                MsgBox "Invalid selection.Can't proceed to this operation!", vbExclamation
                Exit Sub
            End If
            
            If MsgBox("Are you sure you want to delete this record?", vbCritical + vbYesNo) = vbYes Then
                sSQL_Delete "DELETE FROM Sales_Order_Details WHERE SalesOrderID=" & srcRecord
                sSQL_Delete "DELETE FROM Sales_Order WHERE SalesOrderID=" & srcRecord
                
                
                MsgBox "Selected record successfully deleted!", vbInformation
                Call ReloadListview
            Else
                Exit Sub
            End If
            
    Case "Refresh" 'Refresh
           Call ReloadListview
           
    Case "Export" 'Preview
            With lvList
                If .ListItems.Count = 0 Then
                    MsgBox "There's no records to export!Please check it.", vbExclamation
                    Exit Sub
                End If
            End With
                         
            XLSFILENAME = ""
            
            With MAIN.CDExporter
                .Filter = "Excel Workbook (.xls)"
                .ShowSave
            XLSFILENAME = .FileName
            End With
            
            If XLSFILENAME = "" Then
            Exit Sub
            End If
            
            
            Call ExportListview(lvList, XLSFILENAME)
            MsgBox "Records successfully exported!", vbInformation
            XLSFILENAME = ""
            Call ReloadListview
            
    Case "Close" 'Close
            Unload Me
End Select
Exit Sub
errPerformWhat:
     MsgBox "Error Number:" & err.Number & vbNewLine & _
            "Description:" & err.Description, vbExclamation
End Sub

Private Sub FillListview()
On Error Resume Next
With lvList
    .View = lvwReport
    .Gridlines = False
    
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "SalesOrderID"
    .ColumnHeaders.Add , , "CustomerID"
    .ColumnHeaders.Add , , "CompanyName"
    .ColumnHeaders.Add , , "Address"
    .ColumnHeaders.Add , , "Date"
    .ColumnHeaders.Add , , "Salesman"
    .ColumnHeaders.Add , , "Gross"
    .ColumnHeaders.Add , , "Discount"
    .ColumnHeaders.Add , , "NetAmount"
    .ColumnHeaders.Add , , "StatusID"
    .ColumnHeaders.Add , , "Remarks"
    
    
    .ListItems.Clear
    Do While Not RS_SALES.EOF
    Set srcItem = .ListItems.Add(, , RS_SALES.Fields("SalesOrderID"), 1, 1)
        srcItem.SubItems(1) = RS_SALES.Fields("CustomerID")
        srcItem.SubItems(2) = RS_SALES.Fields("CustomerName")
        srcItem.SubItems(3) = RS_SALES.Fields("Address")
        srcItem.SubItems(4) = RS_SALES.Fields("Date")
        srcItem.SubItems(5) = RS_SALES.Fields("Salesman")
        srcItem.SubItems(6) = toMoney(RS_SALES.Fields("Gross"))
        srcItem.SubItems(7) = toNumber(RS_SALES.Fields("Discount"))
        srcItem.SubItems(8) = toMoney(RS_SALES.Fields("NetAmount"))
        srcItem.SubItems(9) = RS_SALES.Fields("StatusID")
        srcItem.SubItems(10) = RS_SALES.Fields("Remarks")
    RS_SALES.MoveNext
    Loop
End With
End Sub


Private Sub ReloadListview()
On Error Resume Next

srcSQL = " SELECT qry_Sales_Order.* " & _
            " FROM qry_Sales_Order " & _
            " ORDER BY qry_Sales_Order.SalesOrderID ASC"

Set RS_SALES = New ADODB.Recordset
If RS_SALES.State = adStateOpen Then RS_SALES.Close
RS_SALES.Open srcSQL, CN, adOpenDynamic, adLockOptimistic

srcSales = "NONE"
srcRecord = vbNullString

Call FillListview
Call lvSizeColumns(lvList)
Call RefreshRecSum
End Sub

Private Sub RefreshRecSum()
    lblRecSum.Caption = "Record: " & srcSales & " of " & lvList.ListItems.Count
End Sub






