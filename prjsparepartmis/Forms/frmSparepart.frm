VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSparepart 
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   15240
   Begin MSDataListLib.DataCombo dcMake 
      Height          =   315
      Left            =   0
      TabIndex        =   10
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
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
      Left            =   6750
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   15240
      TabIndex        =   0
      Top             =   7350
      Width           =   15240
      Begin VB.PictureBox picStatus 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   10920
         ScaleHeight     =   345
         ScaleWidth      =   4275
         TabIndex        =   19
         Top             =   120
         Width           =   4275
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00008000&
            Height          =   255
            Left            =   1560
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   22
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H000000FF&
            Height          =   255
            Left            =   3000
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   21
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   360
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   20
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Reorder Level"
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
            Index           =   4
            Left            =   1920
            TabIndex        =   25
            Top             =   30
            Width           =   1005
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Out of Stock"
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
            Left            =   3255
            TabIndex        =   24
            Top             =   30
            Width           =   900
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status OK"
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
            Left            =   630
            TabIndex        =   23
            Top             =   15
            Width           =   720
         End
      End
      Begin prjcmosxp.Liner Liner1 
         Height          =   30
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   10335
         _ExtentX        =   18230
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
         TabIndex        =   2
         Top             =   120
         Width           =   690
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   5175
      Left            =   1080
      TabIndex        =   4
      Top             =   1320
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9128
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
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
      TabIndex        =   5
      Top             =   45
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   661
      Caption         =   "Manage Spare Parts"
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
      Left            =   8760
      TabIndex        =   6
      Top             =   840
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
      Image           =   "frmSparepart.frx":0000
      cBack           =   -2147483633
   End
   Begin MSDataListLib.DataCombo dcType 
      Height          =   315
      Left            =   2160
      TabIndex        =   11
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
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
   Begin MSDataListLib.DataCombo dcCategory 
      Height          =   315
      Left            =   4320
      TabIndex        =   12
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
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
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   405
      Left            =   50
      TabIndex        =   13
      Top             =   1755
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
      Image           =   "frmSparepart.frx":077A
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdNew 
      Height          =   405
      Left            =   50
      TabIndex        =   14
      Top             =   1320
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
      Image           =   "frmSparepart.frx":0909
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   405
      Left            =   50
      TabIndex        =   15
      Top             =   2190
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
      Image           =   "frmSparepart.frx":0A63
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   405
      Left            =   50
      TabIndex        =   16
      Top             =   2640
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
      Image           =   "frmSparepart.frx":3D00
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   405
      Left            =   50
      TabIndex        =   17
      Top             =   4680
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
      Image           =   "frmSparepart.frx":447A
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdExport 
      Height          =   405
      Left            =   50
      TabIndex        =   18
      Top             =   4245
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
      Image           =   "frmSparepart.frx":76DC
      cBack           =   16119285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part Category"
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
      Left            =   4320
      TabIndex        =   9
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Car Type/Model"
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
      Left            =   2160
      TabIndex        =   8
      Top             =   600
      Width           =   1350
   End
   Begin VB.Image picArrow 
      Height          =   255
      Left            =   6435
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Car Make"
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
      Index           =   0
      Left            =   45
      TabIndex        =   7
      Top             =   600
      Width           =   1365
   End
End
Attribute VB_Name = "frmSparepart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim srcItem                        As ListItem
Dim srcRecord                      As String
Dim srcSparepart                   As Variant
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
Dim qSQL As String

qSQL = "SELECT qry_Spareparts.* " & _
        "FROM qry_Spareparts " & _
        "WHERE MakeName='" & dcMake.Text & _
        "' AND CarTypeName='" & dcType.Text & _
        "' AND PCategoryName='" & dcCategory.Text & _
        "' AND PartDescription Like '%" & txtSearch.Text & "%'"

Set RS_SPAREPART = New ADODB.Recordset
If RS_SPAREPART.State = adStateOpen Then RS_SPAREPART.Close
RS_SPAREPART.Open qSQL, CN, adOpenDynamic, adLockOptimistic

If txtSearch.Text = vbNullString Then Exit Sub:

If RS_SPAREPART.RecordCount < 1 Then
    MsgBox "No record(s) found in the list!", vbExclamation
    txtSearch.SetFocus
    Exit Sub
Else
    Call FillListview
    Call lvSizeColumns(lvList)
End If
End Sub

Private Sub dcMake_Click(Area As Integer)
    bind_dc "SELECT Car_Types.* FROM Car_Types WHERE MakeName='" & dcMake.Text & "'", "CarTypeName", dcType, "CarTypeName", True
End Sub


Private Sub Form_Activate()
On Error Resume Next
HighlightInWin Name

With MAIN
    Me.BackColor = .ACPMenu.BackColor
    Me.Picture = .ACPMenu.LoadBackground
    picFooter.BackColor = .ACPMenu.BackColor
    picStatus.BackColor = .ACPMenu.BackColor
End With

lvList.FlatScrollBar = True

End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
MAIN.AddToWin "Spare Parts", Name

picArrow.Picture = MAIN.i16x16.ListImages(3).Picture

Set lvList.SmallIcons = MAIN.i16x16
Set lvList.Icons = MAIN.i16x16

bind_dc "SELECT Car_Makes.* FROM Car_Makes", "MakeName", dcMake
bind_dc "SELECT Car_Types.* FROM Car_Types", "CarTypeName", dcType
bind_dc "SELECT Part_Categories.* FROM Part_Categories", "PCategoryName", dcCategory

Call ReloadListview

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
        lvList.Height = Me.ScaleHeight - 1750
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
MAIN.RemToWin Me.Caption

Set frmSparepart = Nothing
Set RS_SPAREPART = Nothing
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
srcSparepart = lvList.SelectedItem.Index
srcRecord = lvList.ListItems.Item(srcSparepart).Text
Call RefreshRecSum
End Sub

Private Sub lvList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu MAIN.mnuAction
End Sub

Private Sub picFooter_Resize()
picStatus.Left = picFooter.ScaleWidth - picStatus.ScaleWidth
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
            With frmSparepartAE
                .State = AddStateMode
                .show vbModal
            End With
    Case "Update" 'Update
            If srcRecord = vbNullString Then
                MsgBox "Invalid selection.Can't proceed to this operation!", vbExclamation
                Exit Sub
            Else
                With frmSparepartAE
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
                sSQL_Delete "DELETE FROM Spare_Parts WHERE PartID='" & srcRecord & "'"
                
                Kill (App.Path & "\Graphics\Spare Parts\" & srcRecord & ".img")
                
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
                .Filter = "Excel Files(*.xls)|*.xls|Excel 2007 Files(*.xlsx)|*.xlsx"
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
    .Gridlines = False
    .View = lvwReport
    
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "PartID"
    .ColumnHeaders.Add , , "PartNumber"
    .ColumnHeaders.Add , , "PartDescription"
    .ColumnHeaders.Add , , "Price"
    .ColumnHeaders.Add , , "Inventory"
    .ColumnHeaders.Add , , "ReOrder"
    .ColumnHeaders.Add , , "MakeID"
    .ColumnHeaders.Add , , "MakeName"
    .ColumnHeaders.Add , , "CarTypeID"
    .ColumnHeaders.Add , , "CarTypeName"
    .ColumnHeaders.Add , , "PCategoryID"
    .ColumnHeaders.Add , , "PCategoryName"
    .ColumnHeaders.Add , , "Year"
    .ColumnHeaders.Add , , "Capacity"
    .ColumnHeaders.Add , , "Gearbox"
    
    .ListItems.Clear
    Do While Not RS_SPAREPART.EOF
    Set srcItem = .ListItems.Add(, , RS_SPAREPART.Fields("PartID"), 1, 1)
        srcItem.SubItems(1) = RS_SPAREPART.Fields("PartNumber")
        srcItem.SubItems(2) = RS_SPAREPART.Fields("PartDescription")
        srcItem.SubItems(3) = toMoney(RS_SPAREPART.Fields("Price"))
        srcItem.SubItems(4) = toNumber(RS_SPAREPART.Fields("Inventory"))
        srcItem.SubItems(5) = toNumber(RS_SPAREPART.Fields("ReOrder"))
        srcItem.SubItems(6) = RS_SPAREPART.Fields("MakeID")
        srcItem.SubItems(7) = RS_SPAREPART.Fields("MakeName")
        srcItem.SubItems(8) = RS_SPAREPART.Fields("CarTypeID")
        srcItem.SubItems(9) = RS_SPAREPART.Fields("CarTypeName")
        srcItem.SubItems(10) = RS_SPAREPART.Fields("PCategoryID")
        srcItem.SubItems(11) = RS_SPAREPART.Fields("PCategoryName")
        srcItem.SubItems(12) = RS_SPAREPART.Fields("Year")
        srcItem.SubItems(13) = RS_SPAREPART.Fields("Capacity")
        srcItem.SubItems(14) = RS_SPAREPART.Fields("Gearbox")
        
        If RS_SPAREPART!Inventory <= 1 Then
            srcItem.ListSubItems(4).ForeColor = &HFF&
            srcItem.ListSubItems(4).Bold = True
        Else
            If RS_SPAREPART!Inventory <= RS_SPAREPART!ReOrder Then
                srcItem.ListSubItems(4).ForeColor = &H8000&
                srcItem.ListSubItems(4).Bold = True
            End If
        End If
                    
    RS_SPAREPART.MoveNext
    Loop
End With
End Sub


Private Sub ReloadListview()
On Error Resume Next
srcSQL = " SELECT Spare_Parts.* " & _
            " FROM Spare_Parts " & _
            " ORDER BY Spare_Parts.PartID ASC"

Set RS_SPAREPART = New ADODB.Recordset
If RS_SPAREPART.State = adStateOpen Then RS_SPAREPART.Close
RS_SPAREPART.Open srcSQL, CN, adOpenDynamic, adLockOptimistic

srcSparepart = "NONE"
srcRecord = vbNullString

Call FillListview
Call lvSizeColumns(lvList)
Call RefreshRecSum

End Sub

Private Sub RefreshRecSum()
    lblRecSum.Caption = "Record: " & srcSparepart & " of " & lvList.ListItems.Count
End Sub

Private Sub txtSearch_LostFocus()
unHLText txtSearch
End Sub




