VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPCategoryAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
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
      Height          =   1155
      Index           =   2
      Left            =   1320
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Tag             =   "Name"
      Top             =   1920
      Width           =   5115
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
      Left            =   1320
      MaxLength       =   100
      TabIndex        =   1
      Tag             =   "Name"
      Top             =   1560
      Width           =   3195
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
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1965
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
      ScaleWidth      =   6600
      TabIndex        =   8
      Top             =   0
      Width           =   6600
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "PART CATEGORY DETAILS"
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
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmPCategoryAE.frx":0000
         Top             =   120
         Width           =   720
      End
   End
   Begin prjcmosxp.Liner Liner2 
      Height          =   30
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   53
   End
   Begin prjcmosxp.Liner Liner1 
      Height          =   30
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   53
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
      TabIndex        =   13
      Top             =   1920
      Width           =   615
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
      TabIndex        =   12
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CategoryID"
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
      Left            =   135
      TabIndex        =   11
      Top             =   1200
      Width           =   840
   End
End
Attribute VB_Name = "frmPCategoryAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                        As FORM_STATE
Public PK                           As String

Dim sSQL                            As String

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
                RS_PCATEGORY.AddNew
                
                If isRecordExist("Part_Categories", "PCategoryID", txtEntry(0).Text, True) = True Then
                    MsgBox "Part Category ID already exist in the database.Please check it!", vbExclamation
                    Exit Sub
                End If
                
                RS_PCATEGORY.Fields("PCategoryID") = txtEntry(0).Text
                RS_PCATEGORY.Fields("PCategoryName") = txtEntry(1).Text
                RS_PCATEGORY.Fields("Remarks") = txtEntry(2).Text
                RS_PCATEGORY.Fields("DateEncoded") = Format(Now, "M/d/yyyy")
                RS_PCATEGORY.Fields("EncodedBy") = ACTIVE_USER.USERNAME
                RS_PCATEGORY.Update
                
                MsgBox "New part category has been successfully saved!", vbInformation
                Unload Me
            
            ElseIf State = EditStateMode Then
                RS_PCATEGORY.Fields("PCategoryID") = txtEntry(0).Text
                RS_PCATEGORY.Fields("PCategoryName") = txtEntry(1).Text
                RS_PCATEGORY.Fields("Remarks") = txtEntry(2).Text
                RS_PCATEGORY.Fields("LastDateModified") = Now
                RS_PCATEGORY.Fields("ModifiedBy") = ACTIVE_USER.USERNAME
                RS_PCATEGORY.Update
                
                MsgBox "Information saved successfully!", vbInformation
                Unload Me
                
            End If
End Sub

Private Sub Form_Activate()
On Error Resume Next

    Me.BackColor = MAIN.ACPMenu.BackColor
    
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
CenterForm frmPCategoryAE

    
If State = AddStateMode Then
    Me.Caption = "Create New Entry"
    txtEntry(0).Locked = False
    cmdHistory.Enabled = False
    
ElseIf State = EditStateMode Then
    txtEntry(0).Locked = True
    Me.Caption = "Modify Existing Entry"
    cmdHistory.Enabled = True
    
    sSQL = "SELECT Part_Categories.* " & _
                "FROM Part_Categories " & _
                "WHERE (((Part_Categories.PCategoryID)='" & PK & "'))"

    Set RS_PCATEGORY = New ADODB.Recordset
    If RS_PCATEGORY.State = adStateOpen Then RS_PCATEGORY.Close
    RS_PCATEGORY.Open sSQL, CN, adOpenDynamic, adLockOptimistic
    
    With RS_PCATEGORY
        txtEntry(0).Text = .Fields("PCategoryID")
        txtEntry(1).Text = .Fields("PCategoryName")
        txtEntry(2).Text = .Fields("Remarks")
    End With

End If

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHistory_Click()
On Error Resume Next
    Dim DE As String
    Dim DM As String
    Dim EB As String
    Dim MB As String
    
    DE = Format$(RS_PCATEGORY.Fields("DateEncoded"), "MMM-dd-yyyy HH:MM AMPM")
    DM = Format$(RS_PCATEGORY.Fields("LastDateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    EB = getValueAt("SELECT * FROM Part_Categories WHERE PCategoryID = '" & txtEntry(0).Text & "'", "EncodedBy")
    MB = getValueAt("SELECT * FROM Part_Categories WHERE PCategoryID = '" & txtEntry(0).Text & "'", "ModifiedBy")
    
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

Private Sub Form_Unload(Cancel As Integer)
Set frmPCategoryAE = Nothing
Set RS_PCATEGORY = Nothing

frmPCategory.CommandPass "Refresh"
End Sub


Private Sub txtEntry_GotFocus(Index As Integer)
HLText txtEntry(Index)
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
unHLText txtEntry(Index)
End Sub



