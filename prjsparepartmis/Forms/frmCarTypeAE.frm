VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "NSDataCombo.ocx"
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "XPTab.ocx"
Begin VB.Form frmCarTypeAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CDCM 
      Left            =   120
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prjXTab.XTab XPTab 
      Height          =   4455
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7858
      TabCaption(0)   =   "Car Type"
      TabContCtrlCnt(0)=   9
      Tab(0)ContCtrlCap(1)=   "txtEntry2"
      Tab(0)ContCtrlCap(2)=   "txtCarMakeID"
      Tab(0)ContCtrlCap(3)=   "txtEntry0"
      Tab(0)ContCtrlCap(4)=   "txtEntry1"
      Tab(0)ContCtrlCap(5)=   "txtEntry3"
      Tab(0)ContCtrlCap(6)=   "Labels12"
      Tab(0)ContCtrlCap(7)=   "Labels0"
      Tab(0)ContCtrlCap(8)=   "Labels1"
      Tab(0)ContCtrlCap(9)=   "Label8"
      TabCaption(1)   =   "Photo"
      TabContCtrlCnt(1)=   2
      Tab(1)ContCtrlCap(1)=   "picFile"
      Tab(1)ContCtrlCap(2)=   "shpPicture"
      TabCaption(2)   =   "General Info"
      TabContCtrlCnt(2)=   8
      Tab(2)ContCtrlCap(1)=   "txtEntry7"
      Tab(2)ContCtrlCap(2)=   "txtEntry6"
      Tab(2)ContCtrlCap(3)=   "txtEntry5"
      Tab(2)ContCtrlCap(4)=   "txtEntry4"
      Tab(2)ContCtrlCap(5)=   "Labels3"
      Tab(2)ContCtrlCap(6)=   "Labels2"
      Tab(2)ContCtrlCap(7)=   "Label4"
      Tab(2)ContCtrlCap(8)=   "Label1"
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
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   7
         Left            =   -73320
         Locked          =   -1  'True
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   23
         Tag             =   "Name"
         Top             =   1680
         Width           =   2500
      End
      Begin VB.TextBox txtEntry 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   6
         Left            =   -73320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1320
         Width           =   2500
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
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   5
         Left            =   -73320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   17
         Tag             =   "Name"
         Top             =   960
         Width           =   2500
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
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   4
         Left            =   -73320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   16
         Tag             =   "Name"
         Top             =   600
         Width           =   2500
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
         Index           =   2
         Left            =   1440
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Tag             =   "Name"
         Top             =   1680
         Width           =   3735
      End
      Begin ctrlNSDataCombo.NSDataCombo txtCarMakeID 
         Height          =   330
         Left            =   1440
         TabIndex        =   2
         Top             =   1320
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   1725
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
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   1
         Tag             =   "Name"
         Top             =   960
         Width           =   3730
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
         Left            =   2820
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   3
         Tag             =   "Name"
         Top             =   1320
         Width           =   2355
      End
      Begin VB.Image picFile 
         Height          =   3855
         Left            =   -74880
         Stretch         =   -1  'True
         ToolTipText     =   "Right click here to see other option..."
         Top             =   480
         Width           =   8295
      End
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Encoded"
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
         Index           =   3
         Left            =   -74760
         TabIndex        =   22
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Encoded By"
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
         Index           =   2
         Left            =   -74760
         TabIndex        =   21
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modified By"
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
         TabIndex        =   20
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Date Modified"
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
         TabIndex        =   19
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CarTypeID"
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
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CarType Name"
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
         TabIndex        =   13
         Top             =   960
         Width           =   1065
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
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   675
      End
      Begin VB.Shape shpPicture 
         BorderColor     =   &H00808080&
         Height          =   3855
         Left            =   -74880
         Top             =   480
         Width           =   8295
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
      ScaleWidth      =   8745
      TabIndex        =   7
      Top             =   0
      Width           =   8745
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmCarTypeAE.frx":0000
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
         TabIndex        =   9
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "CAR TYPE DETAILS"
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
         TabIndex        =   8
         Top             =   240
         Width           =   3495
      End
   End
   Begin prjcmosxp.Liner Liner1 
      Height          =   30
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   53
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   345
      Left            =   6240
      TabIndex        =   5
      Top             =   5640
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
      Left            =   7515
      TabIndex        =   6
      Top             =   5640
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
End
Attribute VB_Name = "frmCarTypeAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                        As FORM_STATE
Public PK                           As String

Dim FN                              As String
Dim sSQL                            As String
Dim imgBytes()                      As Byte

Private Sub cmdSave_Click()
            On Error Resume Next
            If is_empty(txtEntry(0)) = True Then txtEntry(0).SetFocus: Exit Sub
            If is_empty(txtEntry(1)) = True Then txtEntry(1).SetFocus: Exit Sub
            If is_empty(txtEntry(2)) = True Then txtEntry(2).SetFocus: Exit Sub
            If is_empty(txtEntry(3)) = True Then txtEntry(3).SetFocus: Exit Sub
            If is_empty(txtCarMakeID) = True Then txtCarMakeID.SetFocus: Exit Sub



            If picFile.Picture = LoadPicture("") Then
                MsgBox "Please upload picture file for this car type/model!", vbExclamation
                Exit Sub
            End If
            
            If State = AddStateMode Then
                RS_CARTYPE.AddNew
                
                If isRecordExist("Car_Types", "CarTypeID", txtEntry(0).Text, True) = True Then
                    MsgBox "Car TypeID already exist in the database.Please check it!", vbExclamation
                    Exit Sub
                End If
                
                RS_CARTYPE.Fields("CarTypeID") = txtEntry(0).Text
                RS_CARTYPE.Fields("CarTypeName") = txtEntry(1).Text
                RS_CARTYPE.Fields("MakeID") = txtCarMakeID.Text
                RS_CARTYPE.Fields("MakeName") = txtEntry(3).Text
                RS_CARTYPE.Fields("Remarks") = txtEntry(2).Text
                RS_CARTYPE.Fields("picFile").AppendChunk imgBytes
                RS_CARTYPE.Fields("DateEncoded") = Format(Now, "M/d/yyyy")
                RS_CARTYPE.Fields("EncodedBy") = ACTIVE_USER.USERNAME
                RS_CARTYPE.Update
                
                MsgBox "New car model has been successfully saved!", vbInformation
                SavePicture picFile.Picture, App.Path & "\Graphics\Car Types\" & txtEntry(0).Text & ".img"
                Unload Me
            
            ElseIf State = EditStateMode Then
                RS_CARTYPE.Fields("CarTypeID") = txtEntry(0).Text
                RS_CARTYPE.Fields("CarTypeName") = txtEntry(1).Text
                RS_CARTYPE.Fields("MakeID") = txtCarMakeID.Text
                RS_CARTYPE.Fields("MakeName") = txtEntry(3).Text
                RS_CARTYPE.Fields("Remarks") = txtEntry(2).Text
                RS_CARTYPE.Fields("picFile").AppendChunk imgBytes
                RS_CARTYPE.Fields("LastDateModified") = Now
                RS_CARTYPE.Fields("ModifiedBy") = ACTIVE_USER.USERNAME
                RS_CARTYPE.Update
                
                MsgBox "Information saved successfully!", vbInformation
                SavePicture picFile.Picture, App.Path & "\Graphics\Car Types\" & txtEntry(0).Text & ".img"
                Unload Me
                
            End If
End Sub

Private Sub Form_Activate()
On Error Resume Next

    Me.BackColor = MAIN.ACPMenu.BackColor
    XPTab.TabStripBackColor = MAIN.ACPMenu.BackColor
    
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
CenterForm frmCarTypeAE
InitializeNSD
    
If State = AddStateMode Then
    Me.Caption = "Create New Entry"
    txtEntry(0).Locked = False
    
    Call LoadNullPicFile
    
ElseIf State = EditStateMode Then
    On Error Resume Next
    txtEntry(0).Locked = True
    Me.Caption = "Modify Existing Entry"
    
    sSQL = "SELECT Car_Types.* " & _
                "FROM Car_Types " & _
                "WHERE (((Car_Types.CarTypeID)='" & PK & "'))"

    Set RS_CARTYPE = New ADODB.Recordset
    If RS_CARTYPE.State = adStateOpen Then RS_CARTYPE.Close
    RS_CARTYPE.Open sSQL, CN, adOpenDynamic, adLockOptimistic
    
    With RS_CARTYPE
        txtEntry(0).Text = .Fields("CarTypeID")
        txtEntry(1).Text = .Fields("CarTypeName")
        txtCarMakeID.Text = .Fields("MakeID")
        txtEntry(3).Text = .Fields("MakeName")
        txtEntry(2).Text = .Fields("Remarks")
        txtEntry(4).Text = .Fields("DateEncoded")
        txtEntry(5).Text = .Fields("EncodedBy")
        txtEntry(6).Text = .Fields("LastDateModified")
        txtEntry(7).Text = .Fields("ModifiedBy")
    End With

    Call LoadPicFile
End If

ErrHandler:
    If err.Number = 53 Then
        Call LoadNullPicFile
    Else
        Call LoadPicFile
    End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub



Private Sub Form_Unload(Cancel As Integer)
Set frmCarTypeAE = Nothing
Set RS_CARTYPE = Nothing

frmCarType.CommandPass "Refresh"
End Sub

Private Sub picFile_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu MAIN.mUpload
End Sub

Private Sub txtCarMakeID_Change()
txtEntry(3).Text = txtCarMakeID.getSelValueAt(2)
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
HLText txtEntry(Index)
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
unHLText txtEntry(Index)
End Sub

Private Sub InitializeNSD()
On Error Resume Next
With txtCarMakeID
        .ClearColumn
        .AddColumn "MakeID", 1500.89
        .AddColumn "MakeName", 2500.88
        .AddColumn "Remarks", 2100.23
        .Connection = CN.ConnectionString
        
        .SQLFields = "MakeID,MakeName,Remarks"
        .sqlTables = "Car_Makes"
        .sqlSortOrder = "MakeID ASC"
        
        .BoundField = "MakeID"
        .PageBy = 10
        .DisplayCol = 1
        
        .setDropWindowSize 6500, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Select Car Make"
End With
End Sub

Public Sub CommandPass(ByVal srcPerformWhat As String)
On Error GoTo errPerformWhat
Select Case srcPerformWhat
    Case "Upload" 'New
            Call BrowsePicFile
    Case "Delete Car Type"
            Call LoadNullPicFile
   
End Select
Exit Sub
errPerformWhat:
     MsgBox "Error Number:" & err.Number & vbNewLine & _
            "Description:" & err.Description, vbExclamation
End Sub

Private Sub LoadPicFile()
On Error Resume Next
    
    FN = App.Path & "\Graphics\Car Types\" & txtEntry(0).Text & ".img"
    
    With CDCM
        .CancelError = True
        .FileName = FN
        
        If .FileName <> "" Then
            Me.MousePointer = vbHourglass
            
            picFile.Picture = LoadPicture(.FileName)
            
            ReDim imgBytes(FileLen(.FileName))
            Open .FileName For Binary As #1
                Get #1, , imgBytes
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
        
        ReDim imgBytes(FileLen(.FileName))
        Open .FileName For Binary As #1
            Get #1, , imgBytes
        Close #1
        
        Me.MousePointer = vbDefault
    Else
        Call LoadNullPicFile
    End If
End With
End Sub

Private Sub LoadNullPicFile()
On Error Resume Next
    
    FN = App.Path & "\Graphics\Car Types\" & "Null.img"
    
    With CDCM
        .CancelError = True
        .FileName = FN
        
        If .FileName <> "" Then
            Me.MousePointer = vbHourglass
            
            picFile.Picture = LoadPicture(.FileName)
            
            ReDim imgBytes(FileLen(.FileName))
            Open .FileName For Binary As #1
                Get #1, , imgBytes
            Close #1
        
            Me.MousePointer = vbDefault
        End If
    End With
End Sub

