VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmUserTypeAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New Entry"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin prjcmosxp.Liner Liner2 
      Height          =   30
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   53
   End
   Begin prjcmosxp.Liner Liner1 
      Height          =   30
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   53
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
      ScaleWidth      =   6555
      TabIndex        =   9
      Top             =   0
      Width           =   6555
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmUserTypeAE.frx":0000
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
         TabIndex        =   11
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "USER TYPE DETAILS"
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
         TabIndex        =   10
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.TextBox txtUserTypeID 
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
      Width           =   2160
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1395
      TabIndex        =   1
      Top             =   1560
      Width           =   5040
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
      Height          =   1170
      Left            =   1395
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1920
      Width           =   5040
   End
   Begin lvButton.lvButtons_H cmdHistory 
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   3480
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
      Left            =   4035
      TabIndex        =   3
      Top             =   3480
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
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   345
      Left            =   5235
      TabIndex        =   4
      Top             =   3480
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
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      TabIndex        =   7
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label Label6 
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
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   615
   End
End
Attribute VB_Name = "frmUserTypeAE"
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
    
    DE = Format$(RS_USERTYPE.Fields("DateEncoded"), "MMM-dd-yyyy HH:MM AMPM")
    DM = Format$(RS_USERTYPE.Fields("LastDateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    EB = getValueAt("SELECT * FROM User_Types WHERE UserTypeID = '" & txtUserTypeID.Text & "'", "EncodedBy")
    MB = getValueAt("SELECT * FROM User_Types WHERE UserTypeID = '" & txtUserTypeID.Text & "'", "ModifiedBy")
    
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
                Set RS_USERTYPE = New ADODB.Recordset
                sSQL_Insert "INSERT INTO User_Types (UserTypeID, UserType, Remarks,DateEncoded, EncodedBy) VALUES ('" & txtUserTypeID.Text & "', '" & txtDescription.Text & "', '" & txtRemarks.Text & _
                "','" & Format(Now, "M/d/yyyy") & "', '" & ACTIVE_USER.USERNAME & "')"

                MsgBox "New user type entry has been successfully saved!", vbInformation
                Unload Me
            
            ElseIf State = EditStateMode Then

                Set RS_USERTYPE = New ADODB.Recordset
                sSQL_Update "UPDATE User_Types SET UserType= '" & txtDescription.Text & "', Remarks= '" & txtRemarks.Text & _
                "', LastDateModified= '" & Now & "', ModifiedBy= '" & ACTIVE_USER.USERNAME & "' WHERE UserTypeID='" & txtUserTypeID.Text & "'"
                
                MsgBox "Information saved successfully!", vbInformation
                Unload Me
            
            End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
    Me.BackColor = MAIN.ACPMenu.BackColor
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
CenterForm frmUserTypeAE

If State = AddStateMode Then
    Me.Caption = "Create New Entry"
    txtUserTypeID.Text = Format(GenerateCD("User_Types"), "00000")
    cmdHistory.Enabled = False
Else
    Me.Caption = "Modify Entry"
    cmdHistory.Enabled = True

    srcSQL = "SELECT User_Types.* " & _
                "FROM User_Types " & _
                "WHERE (((User_Types.UserTypeID)='" & PK & "'))"

    Set RS_USERTYPE = New ADODB.Recordset
    If RS_USERTYPE.State = adStateOpen Then RS_USERTYPE.Close
    RS_USERTYPE.Open srcSQL, CN, adOpenDynamic, adLockOptimistic
    
    With RS_USERTYPE
        txtUserTypeID.Text = .Fields("UserTypeID")
        txtDescription.Text = .Fields("UserType")
        txtRemarks.Text = .Fields("Remarks")
    End With

End If

Exit Sub
ErrHandler:
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
Set frmUserTypeAE = Nothing
Set RS_USERTYPE = Nothing
frmUserType.CommandPass "Refresh"
End Sub

