VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "NSDataCombo.ocx"
Begin VB.Form frmUserAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New Entry"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin prjcmosxp.Liner Liner2 
      Height          =   30
      Left            =   120
      TabIndex        =   27
      Top             =   4080
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   53
   End
   Begin prjcmosxp.Liner Liner1 
      Height          =   30
      Left            =   0
      TabIndex        =   26
      Top             =   960
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   53
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Administrator?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1395
      TabIndex        =   24
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox txtUserType 
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
      Height          =   330
      Left            =   6300
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1560
      Width           =   2350
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
      ScaleWidth      =   8805
      TabIndex        =   21
      Top             =   0
      Width           =   8805
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmUserAE.frx":0000
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
         Index           =   2
         Left            =   960
         TabIndex        =   25
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "USER DETAILS"
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
         TabIndex        =   22
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.TextBox txtRemarks 
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
      Height          =   930
      Left            =   5280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2640
      Width           =   3300
   End
   Begin VB.TextBox txtPassword 
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
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1395
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   3000
      Width           =   2400
   End
   Begin VB.TextBox txtUsername 
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
      Height          =   330
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2640
      Width           =   2400
   End
   Begin VB.TextBox txtFullname 
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
      Height          =   330
      Left            =   5280
      TabIndex        =   3
      Top             =   1200
      Width           =   3360
   End
   Begin VB.TextBox txtLName 
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
      Height          =   330
      Left            =   1395
      TabIndex        =   2
      Top             =   1920
      Width           =   2400
   End
   Begin VB.TextBox txtFName 
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
      Height          =   330
      Left            =   1395
      TabIndex        =   1
      Top             =   1560
      Width           =   2400
   End
   Begin VB.TextBox txtUserID 
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
      Height          =   330
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   2400
   End
   Begin VB.ComboBox cboStatusCD 
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
      ItemData        =   "frmUserAE.frx":38CF
      Left            =   5280
      List            =   "frmUserAE.frx":38D9
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin lvButton.lvButtons_H cmdHistory 
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   2160
      _ExtentX        =   3810
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
      Left            =   6195
      TabIndex        =   9
      Top             =   4320
      Width           =   1245
      _ExtentX        =   2196
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
      TabIndex        =   10
      Top             =   4320
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
   Begin ctrlNSDataCombo.NSDataCombo txtUserTypeID 
      Height          =   330
      Left            =   5280
      TabIndex        =   4
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Index           =   0
      Left            =   4080
      TabIndex        =   20
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UserTypeID"
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
      Left            =   4080
      TabIndex        =   19
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   4065
      TabIndex        =   18
      Top             =   1920
      Width           =   675
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   690
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login Name"
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
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   825
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fullname"
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
      Left            =   4095
      TabIndex        =   15
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lastname"
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
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   690
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First name"
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
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UserID"
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
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "frmUserAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public State                        As FORM_STATE
Public PK                           As String

Dim srcSQL                          As String


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHistory_Click()
On Error Resume Next
    Dim DE As String
    Dim DM As String
    Dim EB As String
    Dim MB As String
    
    DE = Format$(RS_USER.Fields("DateEncoded"), "MMM-dd-yyyy HH:MM AMPM")
    DM = Format$(RS_USER.Fields("LastDateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    EB = getValueAt("SELECT * FROM Users WHERE UserID = '" & txtUserID.Text & "'", "EncodedBy")
    MB = getValueAt("SELECT * FROM Users WHERE UserID = '" & txtUserID.Text & "'", "ModifiedBy")
    
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
            
            If State = AddStateMode Then
                RS_USER.AddNew
                RS_USER.Fields("UserID") = txtUserID.Text
                RS_USER.Fields("Firstname") = txtFName.Text
                RS_USER.Fields("Lastname") = txtLName.Text
                RS_USER.Fields("Fullname") = txtFullname.Text
                RS_USER.Fields("Username") = txtUsername.Text
                RS_USER.Fields("Password") = txtPassword.Text
                RS_USER.Fields("UserTypeID") = txtUserTypeID.Text
                RS_USER.Fields("UserType") = txtUserType.Text
                RS_USER.Fields("IsAdmin") = changeYNValue(Check1.Value)
                RS_USER.Fields("StatusCD") = cboStatusCD.Text
                RS_USER.Fields("Remarks") = txtRemarks.Text
                RS_USER.Fields("DateEncoded") = Format(Now, "M/d/yyyy")
                RS_USER.Fields("EncodedBy") = ACTIVE_USER.USERNAME
                RS_USER.Update
                
                MsgBox "New user has been successfully saved!", vbInformation
                Unload Me
            
            ElseIf State = EditStateMode Then
                RS_USER.Fields("UserID") = txtUserID.Text
                RS_USER.Fields("Firstname") = txtFName.Text
                RS_USER.Fields("Lastname") = txtLName.Text
                RS_USER.Fields("Fullname") = txtFullname.Text
                RS_USER.Fields("Username") = txtUsername.Text
                RS_USER.Fields("Password") = txtPassword.Text
                RS_USER.Fields("UserTypeID") = txtUserTypeID.Text
                RS_USER.Fields("UserType") = txtUserType.Text
                RS_USER.Fields("IsAdmin") = changeYNValue(Check1.Value)
                RS_USER.Fields("StatusCD") = cboStatusCD.Text
                RS_USER.Fields("Remarks") = txtRemarks.Text
                RS_USER.Fields("LastDateModified") = Now
                RS_USER.Fields("ModifiedBy") = ACTIVE_USER.USERNAME
                RS_USER.Update
                
            MsgBox "Information saved successfully!", vbInformation, Me.Caption
            Unload Me
            
            End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
    cboStatusCD.ListIndex = 0
    txtUserID.SetFocus
    
    Me.BackColor = MAIN.ACPMenu.BackColor
    Check1.BackColor = MAIN.ACPMenu.BackColor
End Sub

Private Sub Form_Load()
On Error GoTo ErrTrapper
Call InitializeNSD
CenterForm frmUserAE

If State = AddStateMode Then
    Me.Caption = "Create New Entry"
    txtUserID.Text = "USR-" & Format(GenerateCD("Users"), "00000")
    cboStatusCD.Text = "ACTIVE"
    cmdHistory.Enabled = False
    
    If txtUserType.Text = "ADMINISTRATOR" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
Else
    Me.Caption = "Modify Entry"
    cmdHistory.Enabled = True
    txtUserID.Locked = True
    txtUsername.Locked = True
    txtPassword.Locked = True
    
    srcSQL = "SELECT Users.* " & _
                "FROM Users " & _
                "WHERE (((Users.UserID)='" & PK & "'))"

    Set RS_USER = New ADODB.Recordset
    If RS_USER.State = adStateOpen Then RS_USER.Close
    RS_USER.Open srcSQL, CN, adOpenDynamic, adLockOptimistic
    
    DisplayForEditing
End If

Exit Sub
ErrTrapper:
    MsgBox "Error Number: " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
ElseIf KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmUserAE = Nothing
Set RS_USER = Nothing
frmUser.CommandPass "Refresh"
End Sub


Private Sub InitializeNSD()
On Error Resume Next
With txtUserTypeID
        .ClearColumn
        .AddColumn "UserTypeID", 1500.89
        .AddColumn "Description", 2500.88
        .AddColumn "Remarks", 2100.23
        .Connection = CN.ConnectionString
        
        .SQLFields = "UserTypeID,UserType,Remarks"
        .sqlTables = "User_Types"
        .sqlSortOrder = "UserTypeID ASC"
        
        .BoundField = "Description"
        .PageBy = 10
        .DisplayCol = 1
        
        .setDropWindowSize 6500, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Select User Type"
End With
End Sub

Private Sub DisplayForEditing()
On Error GoTo ErrHandler
    
    With RS_USER
        txtUserID.Text = .Fields("UserID")
        txtFName.Text = .Fields("Firstname")
        txtLName.Text = .Fields("Lastname")
        txtFullname.Text = .Fields("Fullname")
        txtUserTypeID.Text = .Fields("UserTypeID")
        Check1.Value = changeYNValue(.Fields("IsAdmin"))
        txtUserType.Text = .Fields("UserType")
        cboStatusCD.Text = .Fields("StatusCD")
        txtUsername.Text = .Fields("Username")
        txtPassword.Text = .Fields("Password")
        txtRemarks.Text = .Fields("Remarks")
    End With
    
    Exit Sub
ErrHandler:
    If err.Number = 94 Then Resume Next
End Sub


Private Sub txtLName_LostFocus()
txtFullname.Text = txtFName.Text & Space(1) & txtLName.Text
End Sub

Private Sub txtUserTypeID_Change()
txtUserType.Text = txtUserTypeID.getSelValueAt(2)

If txtUserType.Text = "ADMINISTRATOR" Then
    Check1.Value = 1
Else
    Check1.Value = 0
End If
End Sub
