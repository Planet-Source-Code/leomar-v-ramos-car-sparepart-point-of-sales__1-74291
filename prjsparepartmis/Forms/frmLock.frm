VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmLock 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8565
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10590
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFullname 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   4590
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   8760
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.TextBox txtUsername 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   4590
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   9240
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.Frame frePass 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2535
      Left            =   2760
      TabIndex        =   1
      Top             =   2880
      Width           =   4695
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   600
         Width           =   3420
      End
      Begin lvButton.lvButtons_H cmdUnlock 
         Height          =   345
         Left            =   3600
         TabIndex        =   2
         Top             =   600
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   609
         Caption         =   "&Unlock"
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
         Image           =   "frmLock.frx":0000
         cBack           =   -2147483633
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "valid password!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Fullname"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3600
      TabIndex        =   10
      Top             =   8880
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Login Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3600
      TabIndex        =   9
      Top             =   9240
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This system is locked by the Administrator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   1080
      TabIndex        =   8
      Top             =   480
      Width           =   3450
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S Y S T E M  L O C K E D!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   240
      Width           =   2805
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frmLock.frx":6862
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL                         As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdUnlock_Click()
On Error GoTo ErrHandler
If txtPassword.Text = vbNullString Then
    Exit Sub
End If

If txtPassword.Text <> ACTIVE_USER.PASSWORD Then

    lblPrompt.Caption = "Access failed:"
    lblMsg.Caption = "Invalid Password!"
    lblPrompt.Visible = True
    lblMsg.Visible = True
    
    txtPassword.Text = vbNullString
    txtPassword.SetFocus
    
    Exit Sub
Else
    Unload Me
End If
Exit Sub
ErrHandler:
    MsgBox "Error Number:" & err.Number & vbNewLine & _
            "Description:" & err.Description, vbExclamation, Me.Caption
End Sub

Private Sub Form_Activate()
On Error Resume Next
txtPassword.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
CenterForm frmLock

sSQL = "SELECT Users.* " & _
            "FROM Users " & _
            "WHERE Username='" & ACTIVE_USER.USERNAME & "'"


Set RS_USER = New ADODB.Recordset
If RS_USER.State = adStateOpen Then RS_USER.Close
RS_USER.Open sSQL, CN, adOpenStatic, adLockReadOnly

txtFullname.Text = ACTIVE_USER.FULLNAME
txtUsername.Text = ACTIVE_USER.USERNAME

Exit Sub
ErrHandler:
    MsgBox "Error Number:" & err.Number & vbNewLine & _
            "Description:" & err.Description, vbExclamation
End Sub


Private Sub Form_Resize()
On Error Resume Next
frePass.Left = (ScaleWidth / 2) - (frePass.Width / 2)
frePass.Top = (ScaleHeight / 2) - (frePass.Height / 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RS_USER = Nothing
Set frmLock = Nothing
End Sub

Private Sub txtFullname_GotFocus()
HLText txtFullname
End Sub

Private Sub txtPassword_GotFocus()
HLText txtPassword
End Sub

Private Sub txtUsername_GotFocus()
HLText txtUsername
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdUnlock_Click
End If
End Sub





