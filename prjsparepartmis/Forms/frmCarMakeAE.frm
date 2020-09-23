VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCarMakeAE 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   9450
      TabIndex        =   15
      Top             =   0
      Width           =   9450
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "CAR MAKES/BRANDS DETAILS"
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
         TabIndex        =   17
         Top             =   240
         Width           =   3495
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
         TabIndex        =   16
         Top             =   480
         Width           =   2895
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmCarMakeAE.frx":0000
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame fraCM 
      Height          =   2775
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   2655
      Begin VB.Image picFile 
         Height          =   2415
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2415
      End
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
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1965
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
      Left            =   5040
      MaxLength       =   100
      TabIndex        =   1
      Tag             =   "Name"
      Top             =   2040
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
      Height          =   1725
      Index           =   2
      Left            =   5040
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Tag             =   "Name"
      Top             =   2400
      Width           =   4335
   End
   Begin prjcmosxp.Liner Liner2 
      Height          =   30
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   53
   End
   Begin prjcmosxp.Liner Liner1 
      Height          =   30
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   53
   End
   Begin lvButton.lvButtons_H cmdHistory 
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   4440
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
      Left            =   6960
      TabIndex        =   3
      Top             =   4440
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
      Left            =   8235
      TabIndex        =   4
      Top             =   4440
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
   Begin lvButton.lvButtons_H cmdBrowse 
      Height          =   345
      Left            =   2880
      TabIndex        =   9
      Top             =   3435
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   609
      Caption         =   "Upload..."
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
   Begin lvButton.lvButtons_H cmdClear 
      Height          =   345
      Left            =   2880
      TabIndex        =   10
      Top             =   3840
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   609
      Caption         =   "&Clear"
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
   Begin MSComDlg.CommonDialog CD 
      Left            =   1560
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   3960
      TabIndex        =   13
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Make ID"
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
      Left            =   3960
      TabIndex        =   12
      Top             =   1680
      Width           =   585
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Make Name"
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
      Left            =   3960
      TabIndex        =   11
      Top             =   2055
      Width           =   825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Car Makes/Brands Info."
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
      Left            =   195
      TabIndex        =   8
      Top             =   1080
      Width           =   2865
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000010&
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   120
      Top             =   1080
      Width           =   9255
   End
End
Attribute VB_Name = "frmCarMakeAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                        As FORM_STATE
Public PK                           As String

Dim sSQL                            As String
Dim dataBytes()                     As Byte
Dim FN                              As String



Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHistory_Click()
On Error Resume Next
    Dim DE As String
    Dim DM As String
    Dim EB As String
    Dim MB As String
    
    DE = Format$(RS_CARMAKE.Fields("DateEncoded"), "MMM-dd-yyyy HH:MM AMPM")
    DM = Format$(RS_CARMAKE.Fields("LastDateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    EB = getValueAt("SELECT * FROM Car_Makes WHERE MakeID = '" & txtEntry(0).Text & "'", "EncodedBy")
    MB = getValueAt("SELECT * FROM Car_Makes WHERE MakeID = '" & txtEntry(0).Text & "'", "ModifiedBy")
    
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

Private Sub cmdBrowse_Click()
On Error Resume Next
With CD
    .Filter = "JPG Files(*.jpeg)|*.jpg|JPEG Files(*.jpeg)|*.jpg" & _
    "|GIF Files(*.gif)|*.gif|Bitmap Files(*.bmp)|*.bmp|All Supported Files|*.jpeg;*.jpg;*.gif;*.bmp"
    
    .CancelError = True
    .ShowOpen
    
    If .FileName <> "" Then
        Me.MousePointer = vbHourglass
        
        picFile.Picture = LoadPicture(.FileName)
        
        ReDim dataBytes(FileLen(.FileName))
        Open .FileName For Binary As #1
            Get #1, , dataBytes
        Close #1
        
        Me.MousePointer = vbDefault
    Else
        Call LoadNullPicFile
    End If
End With
End Sub

Private Sub cmdClear_Click()
On Error Resume Next
    Call LoadNullPicFile
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
                MsgBox "Please upload picture file for this car type/model!", vbExclamation
                Exit Sub
            End If
            
            If State = AddStateMode Then
                
                If isRecordExist("Car_Makes", "MakeID", txtEntry(0).Text, True) = True Then
                    MsgBox "Make ID already exist in the database.Please check it!", vbExclamation
                    Exit Sub
                End If
                
                RS_CARMAKE.AddNew
                RS_CARMAKE.Fields("MakeID") = txtEntry(0).Text
                RS_CARMAKE.Fields("MakeName") = txtEntry(1).Text
                RS_CARMAKE.Fields("Remarks") = txtEntry(2).Text
                RS_CARMAKE.Fields("picFile").AppendChunk dataBytes
                RS_CARMAKE.Fields("DateEncoded") = Format(Now, "M/d/yyyy")
                RS_CARMAKE.Fields("EncodedBy") = ACTIVE_USER.USERNAME
                RS_CARMAKE.Update

                MsgBox "New car make/brand has been successfully saved!", vbInformation
                SavePicture picFile.Picture, App.Path & "\Graphics\Car Makes\" & txtEntry(0).Text & ".img"
                
                Unload Me
            
            ElseIf State = EditStateMode Then
                RS_CARMAKE.Fields("MakeID") = txtEntry(0).Text
                RS_CARMAKE.Fields("MakeName") = txtEntry(1).Text
                RS_CARMAKE.Fields("Remarks") = txtEntry(2).Text
                RS_CARMAKE.Fields("picFile").AppendChunk dataBytes
                RS_CARMAKE.Fields("LastDateModified") = Now
                RS_CARMAKE.Fields("ModifiedBy") = ACTIVE_USER.USERNAME
                RS_CARMAKE.Update
                
                MsgBox "Information saved successfully!", vbInformation
                SavePicture picFile.Picture, App.Path & "\Graphics\Car Makes\" & txtEntry(0).Text & ".img"
                
                Unload Me
                
            End If

End Sub

Private Sub Form_Activate()
On Error Resume Next
Me.BackColor = MAIN.ACPMenu.BackColor
fraCM.BackColor = MAIN.ACPMenu.BackColor

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
CenterForm frmCarMakeAE

If State = AddStateMode Then
    Me.Caption = "Create New Entry"
    txtEntry(0).Locked = False
    cmdHistory.Enabled = False
    
    Call LoadNullPicFile
    
ElseIf State = EditStateMode Then
    txtEntry(0).Locked = True
    Me.Caption = "Modify Entry"
    cmdHistory.Enabled = True
    
    sSQL = "SELECT Car_Makes.* " & _
                "FROM Car_Makes " & _
                "WHERE (((Car_Makes.MakeID)='" & PK & "'))"

    Set RS_CARMAKE = New ADODB.Recordset
    If RS_CARMAKE.State = adStateOpen Then RS_CARMAKE.Close
    RS_CARMAKE.Open sSQL, CN, adOpenDynamic, adLockOptimistic
    
    With RS_CARMAKE
        txtEntry(0).Text = .Fields("MakeID")
        txtEntry(1).Text = .Fields("MakeName")
        txtEntry(2).Text = .Fields("Remarks")
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

Private Sub Form_Unload(Cancel As Integer)
Set frmCarMakeAE = Nothing
Set RS_CARMAKE = Nothing

FN = vbNullString
frmCarMake.CommandPass "Refresh"
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
HLText txtEntry(Index)
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
unHLText txtEntry(Index)
End Sub

Private Sub LoadPicFile()
On Error Resume Next
    
    FN = App.Path & "\Graphics\Car Makes\" & txtEntry(0).Text & ".img"
    
    With CD
        .CancelError = True
        .FileName = FN
        
        If .FileName <> "" Then
            Me.MousePointer = vbHourglass
            
            picFile.Picture = LoadPicture(.FileName)
            
            ReDim dataBytes(FileLen(.FileName))
            Open .FileName For Binary As #1
                Get #1, , dataBytes
            Close #1
        
            Me.MousePointer = vbDefault
        End If
    End With
    
End Sub

Private Sub LoadNullPicFile()
On Error Resume Next
    
    FN = App.Path & "\Graphics\Car Makes\" & "Null.img"
    
    With CD
        .CancelError = True
        .FileName = FN
        
        If .FileName <> "" Then
            Me.MousePointer = vbHourglass
            
            picFile.Picture = LoadPicture(.FileName)
            
            ReDim dataBytes(FileLen(.FileName))
            Open .FileName For Binary As #1
                Get #1, , dataBytes
            Close #1
        
            Me.MousePointer = vbDefault
        End If
    End With
End Sub
