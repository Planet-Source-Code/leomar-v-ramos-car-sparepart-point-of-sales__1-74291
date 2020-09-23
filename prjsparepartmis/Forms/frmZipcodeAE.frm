VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmZipcodeAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin prjcmosxp.Liner Liner2 
      Height          =   30
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   53
   End
   Begin prjcmosxp.Liner Liner1 
      Height          =   30
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   53
   End
   Begin VB.TextBox txtZipcode 
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
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1200
      Width           =   2160
   End
   Begin VB.TextBox txtCityTown 
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
      Width           =   4320
   End
   Begin VB.TextBox txtProvince 
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
      Width           =   4320
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
      ScaleWidth      =   5925
      TabIndex        =   6
      Top             =   0
      Width           =   5925
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmZipcodeAE.frx":0000
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "ZIP CODE DETAILS"
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
         TabIndex        =   7
         Top             =   480
         Width           =   2895
      End
   End
   Begin lvButton.lvButtons_H cmdHistory 
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   2640
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
      Left            =   3435
      TabIndex        =   3
      Top             =   2640
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
      Left            =   4635
      TabIndex        =   4
      Top             =   2640
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
      Caption         =   "Zipcode"
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
      Left            =   100
      TabIndex        =   11
      Top             =   1200
      Width           =   555
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City/Town"
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
      Left            =   100
      TabIndex        =   10
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "State/Province"
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
      Left            =   100
      TabIndex        =   9
      Top             =   1920
      Width           =   1065
   End
End
Attribute VB_Name = "frmZipcodeAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                        As FORM_STATE
Public PK                           As String

Dim sSQL                            As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHistory_Click()
On Error Resume Next
    Dim DE As String
    Dim DM As String
    Dim EB As String
    Dim MB As String
    
    DE = Format$(RS_ZIPCODE.Fields("DateEncoded"), "MMM-dd-yyyy HH:MM AMPM")
    DM = Format$(RS_ZIPCODE.Fields("LastDateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    EB = getValueAt("SELECT * FROM Zipcodes WHERE ZipCode = '" & txtZipcode.Text & "'", "EncodedBy")
    MB = getValueAt("SELECT * FROM Zipcodes WHERE ZipCode = '" & txtZipcode.Text & "'", "ModifiedBy")
    
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
            
                If isRecordExist("Zipcodes", "ZipCode", txtZipcode.Text, True) = True Then
                    MsgBox "Zipcode already exist in the database.Please check it!", vbExclamation
                    Exit Sub
                End If
                
                Set RS_ZIPCODE = New ADODB.Recordset
                sSQL_Insert "INSERT INTO Zipcodes (ZipCode, CityTown, Province, DateEncoded, EncodedBy) VALUES ('" & txtZipcode.Text & "', '" & txtCityTown.Text & "', '" & txtProvince.Text & "', '" & Format(Now, "M/d/yyyy") & _
                "', '" & ACTIVE_USER.USERNAME & "')"

                MsgBox "New zipcode entry has been successfully saved!", vbInformation
                Unload Me
            
            ElseIf State = EditStateMode Then

                Set RS_ZIPCODE = New ADODB.Recordset
                sSQL_Update "UPDATE Zipcodes SET CityTown= '" & txtCityTown.Text & "', Province= '" & txtProvince.Text & "', LastDateModified= '" & Now & _
                "', ModifiedBy= '" & ACTIVE_USER.USERNAME & "' WHERE ZipCode='" & txtZipcode.Text & "'"
                
                MsgBox "Information saved successfully!", vbInformation
                Unload Me
                
            End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
Me.BackColor = MAIN.ACPMenu.BackColor
txtZipcode.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
CenterForm frmZipcodeAE

If State = AddStateMode Then
    Me.Caption = "Create New Entry"
    txtZipcode.Locked = False
    cmdHistory.Enabled = False
ElseIf State = EditStateMode Then
    txtZipcode.Locked = True
    Me.Caption = "Modify Entry"
    cmdHistory.Enabled = True
    
    sSQL = "SELECT Zipcodes.* " & _
                "FROM Zipcodes " & _
                "WHERE (((Zipcodes.ZipCode)='" & PK & "'))"

    Set RS_ZIPCODE = New ADODB.Recordset
    If RS_ZIPCODE.State = adStateOpen Then RS_ZIPCODE.Close
    RS_ZIPCODE.Open sSQL, CN, adOpenDynamic, adLockOptimistic
    
    With RS_ZIPCODE
        txtZipcode.Text = .Fields("ZipCode")
        txtCityTown.Text = .Fields("CityTown")
        txtProvince.Text = .Fields("Province")
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
Set frmZipcodeAE = Nothing
Set RS_ZIPCODE = Nothing
frmZipcode.CommandPass "Refresh"
End Sub

Private Sub txtCityTown_GotFocus()
HLText txtCityTown
End Sub

Private Sub txtCityTown_LostFocus()
unHLText txtCityTown
End Sub

Private Sub txtProvince_GotFocus()
HLText txtProvince
End Sub

Private Sub txtProvince_LostFocus()
unHLText txtProvince
End Sub

Private Sub txtZipcode_GotFocus()
HLText txtZipcode
End Sub

Private Sub txtZipcode_LostFocus()
unHLText txtZipcode
End Sub
