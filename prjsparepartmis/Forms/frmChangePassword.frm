VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmChangePassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin prjcmosxp.Liner Liner2 
      Height          =   30
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   5775
      _ExtentX        =   10186
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
      ScaleWidth      =   5985
      TabIndex        =   11
      Top             =   0
      Width           =   5985
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "CHANGE PASSWORD"
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   480
         Width           =   2895
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmChangePassword.frx":0000
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.TextBox txtUsername 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2235
      TabIndex        =   0
      Top             =   1200
      Width           =   3500
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
      Left            =   2235
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   3500
   End
   Begin VB.TextBox txtNewPassword 
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
      Left            =   2235
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1920
      Width           =   3500
   End
   Begin VB.TextBox txtRetype 
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
      Left            =   2235
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2280
      Width           =   3500
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   345
      Left            =   3240
      TabIndex        =   4
      Top             =   2880
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   345
      Left            =   4560
      TabIndex        =   5
      Top             =   2880
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
   Begin prjcmosxp.Liner Liner1 
      Height          =   30
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   53
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   9
      Top             =   1200
      Width           =   825
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password"
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
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
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
      Left            =   375
      TabIndex        =   7
      Top             =   1920
      Width           =   1050
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Re-type Password"
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
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   1320
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                        As FORM_STATE
Dim srcSQL                          As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim obj As Control
Dim sRS As Recordset
            For Each obj In Me
            If TypeOf obj Is TextBox Or TypeOf obj Is ComboBox Then
                If obj.Text = "" Then
                    MsgBox obj.Name & " could not be left blank. Please complete the field.", vbExclamation, Me.Caption
                    obj.SetFocus
                    Exit Sub
                End If
            End If
            Next obj
            
            If txtNewPassword.Text <> txtRetype.Text Then
                MsgBox "Passwords did not match.Please check it!", vbExclamation
                Exit Sub
            End If
            
            Set sRS = New ADODB.Recordset
            If sRS.State = adStateOpen Then sRS.Close
            sRS.Open "SELECT * FROM Users WHERE Username='" & ACTIVE_USER.USERNAME & "'", CN, adOpenStatic, adLockReadOnly
            
            If sRS.Fields("Password") <> txtPassword.Text Then
                MsgBox "Password did not match.Please check it!", vbExclamation
                Exit Sub
            End If
            
            If State = EditStateMode Then
                Set RS_USER = New ADODB.Recordset
                sSQL_Update "UPDATE Users SET Users.Password='" & txtNewPassword.Text & "' WHERE Users.Username='" & ACTIVE_USER.USERNAME & "'"
                
                MsgBox "Password has been successfully updated!", vbInformation
                Unload Me
            End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
Me.BackColor = MAIN.ACPMenu.BackColor

txtPassword.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
CenterForm frmChangePassword

State = EditStateMode

    srcSQL = "SELECT Users.* " & _
            "FROM Users " & _
            "WHERE (((Users.Username)='" & ACTIVE_USER.USERNAME & "'))"

    Set RS_USER = New ADODB.Recordset
    If RS_USER.State = adStateOpen Then RS_USER.Close
    RS_USER.Open srcSQL, CN, adOpenDynamic, adLockOptimistic
    
    With RS_USER
        txtUsername.Text = .Fields("Username")
    End With

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation, Me.Caption
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
ElseIf KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmChangePassword = Nothing
Set RS_USER = Nothing
End Sub


Private Sub txtUsername_GotFocus()
HLText txtUsername
End Sub


