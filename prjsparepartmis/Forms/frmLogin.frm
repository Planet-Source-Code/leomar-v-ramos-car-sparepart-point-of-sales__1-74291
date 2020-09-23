VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5280
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10215
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
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
      Height          =   315
      Left            =   7440
      TabIndex        =   0
      Text            =   "admin"
      Top             =   2040
      Width           =   2715
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   7440
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "admin"
      Top             =   2400
      Width           =   2715
   End
   Begin VB.TextBox txtServer 
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
      IMEMode         =   3  'DISABLE
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   2760
      Width           =   2715
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
      ScaleWidth      =   10215
      TabIndex        =   11
      Top             =   0
      Width           =   10215
      Begin prjcmosxp.Liner Liner1 
         Height          =   30
         Left            =   0
         TabIndex        =   15
         Top             =   960
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   53
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   120
         Picture         =   "frmLogin.frx":0E42
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "R A M O S O F T ®™"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "p r o j e c t"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "c m o s x p v1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1995
         TabIndex        =   12
         Top             =   480
         Width           =   1695
      End
   End
   Begin lvButton.lvButtons_H cmdLogin 
      Height          =   345
      Left            =   7680
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      Caption         =   "&Login"
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
      Image           =   "frmLogin.frx":1C2F
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   345
      Left            =   8880
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
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
      Image           =   "frmLogin.frx":8491
      cBack           =   -2147483633
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">>> USER LOGIN "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6360
      TabIndex        =   10
      Top             =   1080
      Width           =   2025
   End
   Begin VB.Image Image3 
      Height          =   600
      Left            =   6240
      Picture         =   "frmLogin.frx":B6F3
      Top             =   960
      Width           =   3960
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.ramosoft.co.nr"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   5010
      Width           =   1980
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © By Ramosoft . All Rights Reserved 2010"
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
      Left            =   6360
      TabIndex        =   8
      Top             =   5010
      Width           =   3810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Server"
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
      Left            =   6450
      TabIndex        =   7
      Top             =   2760
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Password"
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
      Left            =   6450
      TabIndex        =   6
      Top             =   2400
      Width           =   690
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
      Height          =   195
      Left            =   6450
      TabIndex        =   5
      Top             =   2040
      Width           =   825
   End
   Begin VB.Image Image2 
      Height          =   3975
      Left            =   -120
      Picture         =   "frmLogin.frx":BB42
      Top             =   960
      Width           =   6405
   End
   Begin b8Controls4.b83DRect b83DRect1 
      Height          =   3975
      Left            =   0
      Top             =   960
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7011
      Color1          =   16777215
      Color2          =   16777215
      Color3          =   14737632
      Color4          =   14737632
      BackColor       =   16119285
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Dim dwLen                            As Long
Dim strString                        As String

Private Sub cmdCancel_Click()
    END_APP = True
    Unload Me
End Sub

Private Sub cmdLogin_Click()
If txtUsername.Text = "" Then
    MsgBox "Username and/or Password is incorrect.Try Again!", vbExclamation
    txtUsername.SetFocus
    Exit Sub
End If

If txtPassword.Text = "" Then
    MsgBox "Username and/or Password is incorrect.Try Again!", vbExclamation
    txtPassword.SetFocus
    Exit Sub
End If

If txtServer.Text = "" Then
    MsgBox "Unable to connect to the server.Please check your connection!", vbExclamation
    txtServer.SetFocus
    Exit Sub
End If

Set RS_USER = New ADODB.Recordset
If RS_USER.State = adStateOpen Then RS_USER.Close
RS_USER.Open "SELECT users.* FROM users WHERE Username='" & txtUsername.Text & "' AND Password ='" & txtPassword.Text & "' AND StatusCD ='" & "ACTIVE" & "'", CN, adOpenStatic, adLockReadOnly

If RS_USER.BOF Or RS_USER.EOF = True Then
    MsgBox "Username and/or Password is incorrect.Try Again!", vbExclamation
    Exit Sub

ElseIf RS_USER.Fields("StatusCD") = "INACTIVE" Then
    MsgBox "User account is no longer active.Contact your administrator to re-activate your account!", vbExclamation
    Exit Sub

ElseIf Not RS_USER.Fields("Username") = txtUsername.Text Then
    MsgBox "Username and/or Password is incorrect.Try Again!", vbExclamation
    Exit Sub

ElseIf Not RS_USER.Fields("Password") = txtPassword.Text Then
    MsgBox "Username and/or Password is incorrect.Try Again!", vbExclamation
    Exit Sub

Else

    ACTIVE_USER.USERID = RS_USER.Fields("UserID")
    ACTIVE_USER.FULLNAME = RS_USER.Fields("Fullname")
    ACTIVE_USER.USERNAME = RS_USER.Fields("Username")
    ACTIVE_USER.PASSWORD = RS_USER.Fields("Password")
    ACTIVE_USER.USERTYPE = RS_USER.Fields("UserType")
    ACTIVE_USER.USER_ISADMIN = CBool(changeYNValue(getValueAt("SELECT Username,IsAdmin FROM Users WHERE Username='" & txtUsername.Text & "'", "IsAdmin")))
    
    With MAIN.StatusBar.Panels
        .Item(3).Text = ACTIVE_USER.FULLNAME
        .Item(4).Text = ACTIVE_USER.USERNAME
    End With
    Unload Me
    
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
If END_APP = True Then Unload Me: Exit Sub
txtUsername.SetFocus
End Sub


Private Sub Form_Load()
On Error GoTo ErrHandler
CenterForm frmLogin

'Create a buffer
dwLen = MAX_COMPUTERNAME_LENGTH + 1
strString = String(dwLen, "X")
'Get the computer name
GetComputerName strString, dwLen
strString = Left(strString, dwLen)

If Connected2DB = False Then END_APP = True: Unload Me: Exit Sub

txtServer.Text = strString

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation, Me.Caption
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    End
ElseIf KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then END_APP = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmLogin = Nothing
    Set RS_USER = Nothing
End Sub

Private Sub txtPassword_GotFocus()
HLText txtPassword
End Sub

Private Sub txtPassword_LostFocus()
unHLText txtPassword
End Sub

Private Sub txtServer_GotFocus()
HLText txtServer
End Sub

Private Sub txtServer_LostFocus()
unHLText txtServer
End Sub

Private Sub txtUsername_GotFocus()
HLText txtUsername
End Sub

Private Sub txtUsername_LostFocus()
unHLText txtUsername
End Sub
