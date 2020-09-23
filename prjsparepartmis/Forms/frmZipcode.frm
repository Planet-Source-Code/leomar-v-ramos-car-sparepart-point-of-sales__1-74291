VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls.ocx"
Begin VB.Form frmZipcode 
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   10260
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
      Left            =   1365
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   600
      Width           =   2055
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
      Left            =   3840
      TabIndex        =   11
      Top             =   600
      Width           =   2775
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   10260
      TabIndex        =   0
      Top             =   8145
      Width           =   10260
      Begin prjcmosxp.Liner Liner1 
         Height          =   30
         Left            =   0
         TabIndex        =   10
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
         TabIndex        =   1
         Top             =   120
         Width           =   690
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3615
      Left            =   900
      TabIndex        =   2
      Top             =   1080
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   555
      Left            =   75
      TabIndex        =   3
      Top             =   1680
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   979
      Caption         =   "&Update"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   4210752
      cBhover         =   16119285
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmZipcode.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdNew 
      Height          =   555
      Left            =   75
      TabIndex        =   4
      Top             =   1080
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   979
      Caption         =   "&New"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   4210752
      cBhover         =   16119285
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmZipcode.frx":0577
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   555
      Left            =   75
      TabIndex        =   5
      Top             =   2280
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   979
      Caption         =   "&Delete"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   4210752
      cBhover         =   16119285
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmZipcode.frx":06D1
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   555
      Left            =   75
      TabIndex        =   6
      Top             =   2880
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   979
      Caption         =   "&Refresh"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   4210752
      cBhover         =   16119285
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmZipcode.frx":396E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   555
      Left            =   75
      TabIndex        =   7
      Top             =   4080
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   979
      Caption         =   "&Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   4210752
      cBhover         =   16119285
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmZipcode.frx":40E8
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdExport 
      Height          =   555
      Left            =   75
      TabIndex        =   8
      Top             =   3480
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   979
      Caption         =   "&Export"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   4210752
      cBhover         =   16119285
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmZipcode.frx":734A
      cBack           =   -2147483633
   End
   Begin b8Controls4.b8TitleBar b8TB 
      Height          =   375
      Left            =   50
      TabIndex        =   9
      Top             =   50
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   661
      Caption         =   "Manage Zipcodes"
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
      Left            =   6720
      TabIndex        =   14
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
      Image           =   "frmZipcode.frx":74A4
      cBack           =   -2147483633
   End
   Begin VB.Image picArrow 
      Height          =   255
      Left            =   3525
      Top             =   600
      Width           =   255
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
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   1140
   End
End
Attribute VB_Name = "frmZipcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim srcSQL                      As String
Dim srcItem                     As ListItem
Dim srcRecord                   As String
Dim srcZipcode                  As Variant

Private Sub cmdSearch_Click()
On Error Resume Next
Dim sSQL As String

sSQL = "SELECT Zipcodes.* " & _
        "FROM Zipcodes " & _
        "WHERE (((" & cboFields.Text & ") Like '%" & txtSearch.Text & "%'))"

Set RS_ZIPCODE = New ADODB.Recordset
If RS_ZIPCODE.State = adStateOpen Then RS_ZIPCODE.Close
RS_ZIPCODE.Open sSQL, CN, adOpenDynamic, adLockOptimistic

If txtSearch.Text = vbNullString Then Exit Sub:

If RS_ZIPCODE.RecordCount < 1 Then
    MsgBox "No record(s) found in the list!", vbExclamation
    txtSearch.SetFocus
    Exit Sub
Else
    Call FillListview
    Call lvSizeColumns(lvList)
End If
End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next
CommandPass "Update"
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
MAIN.AddToWin "Zipcodes", Name

picArrow.Picture = MAIN.i16x16.ListImages(3).Picture

Set lvList.SmallIcons = MAIN.i16x16
Set lvList.Icons = MAIN.i16x16

srcSQL = "SELECT Zipcodes.* " & _
            " FROM Zipcodes " & _
            " ORDER BY Zipcodes.ZipCode ASC "

Set RS_ZIPCODE = New ADODB.Recordset
If RS_ZIPCODE.State = adStateOpen Then RS_ZIPCODE.Close
RS_ZIPCODE.Open srcSQL, CN, adOpenDynamic, adLockPessimistic

srcZipcode = "NONE"
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
        lvList.Width = Me.ScaleWidth - 950
        lvList.Height = Me.ScaleHeight - 1600
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
MAIN.RemToWin Me.Caption

Set frmZipcode = Nothing
Set RS_ZIPCODE = Nothing
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

Private Sub lvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next

srcZipcode = lvList.SelectedItem.Index
srcRecord = lvList.ListItems.Item(srcZipcode).Text
Call RefreshRecSum
End Sub

Private Sub lvList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu MAIN.mnuAction
End Sub

Private Sub lvList_DblClick()
On Error Resume Next
CommandPass "Update"
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
CommandPass "Close"
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
CommandPass "Delete"
End Sub

Private Sub cmdExport_Click()
On Error Resume Next
CommandPass "Export"
End Sub

Private Sub cmdNew_Click()
On Error Resume Next
CommandPass "New"
End Sub

Private Sub cmdRefresh_Click()
On Error Resume Next
CommandPass "Refresh"
End Sub


Public Sub CommandPass(ByVal srcPerformWhat As String)
On Error GoTo errPerformWhat
Select Case srcPerformWhat
    Case "New" 'New
            With frmZipcodeAE
                .State = AddStateMode
                .show vbModal
            End With
    Case "Update" 'Update
            If srcRecord = vbNullString Then
                MsgBox "Invalid selection.Can't proceed to this operation!", vbExclamation
                Exit Sub
            Else
                With frmZipcodeAE
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
                sSQL_Delete "DELETE FROM Zipcodes WHERE ZipCode='" & srcRecord & "'"
                MsgBox "Selected record successfully deleted!", vbInformation, Me.Caption
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
    .View = lvwReport
    .FullRowSelect = True
    .Gridlines = False
    
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "ZipCode"
    .ColumnHeaders.Add , , "CityTown"
    .ColumnHeaders.Add , , "Province"
    
    .ListItems.Clear
    Do While Not RS_ZIPCODE.EOF
    Set srcItem = .ListItems.Add(, , RS_ZIPCODE.Fields("ZipCode"), 1, 1)
        srcItem.SubItems(1) = RS_ZIPCODE.Fields("CityTown")
        srcItem.SubItems(2) = RS_ZIPCODE.Fields("Province")
    RS_ZIPCODE.MoveNext
    Loop
End With
End Sub


Private Sub ReloadListview()
On Error Resume Next
srcSQL = " SELECT Zipcodes.* " & _
            " FROM Zipcodes " & _
            " ORDER BY Zipcodes.ZipCode ASC"

Set RS_ZIPCODE = New ADODB.Recordset
If RS_ZIPCODE.State = adStateOpen Then RS_ZIPCODE.Close
RS_ZIPCODE.Open srcSQL, CN, adOpenDynamic, adLockOptimistic

srcRecord = vbNullString
srcZipcode = "NONE"

Call FillListview
Call lvSizeColumns(lvList)
Call RefreshRecSum

End Sub

Private Sub RefreshRecSum()
    lblRecSum.Caption = "Record: " & srcZipcode & " of " & lvList.ListItems.Count
End Sub



Private Sub txtSearch_LostFocus()
unHLText txtSearch
End Sub
