VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmBusiness 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business Information"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6660
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
      ScaleWidth      =   6660
      TabIndex        =   14
      Top             =   0
      Width           =   6660
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmBusiness.frx":0000
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
         Index           =   1
         Left            =   960
         TabIndex        =   16
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "BUSINESS INFORMATION"
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
         TabIndex        =   15
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.TextBox txtAddress 
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
      Left            =   1470
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2160
      Width           =   5025
   End
   Begin VB.TextBox txtCompanyName 
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
      Left            =   1470
      TabIndex        =   1
      Top             =   1800
      Width           =   5025
   End
   Begin VB.TextBox txtCompanyID 
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
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1440
      Width           =   2400
   End
   Begin VB.TextBox txtBusinessNo 
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
      Left            =   1470
      TabIndex        =   3
      Top             =   3120
      Width           =   2505
   End
   Begin VB.TextBox txtFaxNo 
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
      Left            =   4800
      TabIndex        =   4
      Top             =   3120
      Width           =   1680
   End
   Begin VB.TextBox txtEmail 
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
      Left            =   1470
      TabIndex        =   5
      Top             =   3480
      Width           =   2505
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   345
      Left            =   3840
      TabIndex        =   6
      Top             =   4095
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   609
      Caption         =   "&Save Changes"
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
      Left            =   5400
      TabIndex        =   7
      Top             =   4095
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
   Begin prjcmosxp.Liner Liner2 
      Height          =   30
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   6495
      _ExtentX        =   16325
      _ExtentY        =   53
   End
   Begin prjcmosxp.Liner Liner1 
      Height          =   30
      Left            =   0
      TabIndex        =   18
      Top             =   960
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   53
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Details"
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
      TabIndex        =   19
      Top             =   1080
      Width           =   2865
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Top             =   2160
      Width           =   585
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CompanyName"
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
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CompanyID"
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
      TabIndex        =   11
      Top             =   1440
      Width           =   840
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Business Tel. No."
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
      TabIndex        =   10
      Top             =   3120
      Width           =   1230
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax No."
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
      TabIndex        =   9
      Top             =   3120
      Width           =   570
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address"
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
      Top             =   3480
      Width           =   990
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
Attribute VB_Name = "frmBusiness"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public State                        As FORM_STATE

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim obj As Control
            For Each obj In Me
            If TypeOf obj Is TextBox Or TypeOf obj Is ComboBox Then
                If obj.Text = "" Then
                    MsgBox obj.Name & " could not be left blank. Please complete the field.", vbExclamation, Me.Caption
                    obj.SetFocus
                    Exit Sub
                End If
            End If
            Next obj
            
            If State = EditStateMode Then
            
                Set RS_COMPANY = New ADODB.Recordset
                sSQL_Update "UPDATE Company_Info SET CompanyName= '" & txtCompanyName.Text & "', Address= '" & txtAddress.Text & _
                "',BusinessNo='" & txtBusinessNo.Text & "',FaxNo='" & txtFaxNo.Text & "', Email='" & txtEmail.Text & "',LastDateModified= '" & Now & "', ModifiedBy= '" & ACTIVE_USER.USERNAME & "' WHERE CompanyID='" & txtCompanyID.Text & "'"
                
                
                With ACTIVE_COMPANY
                    .COMPANYID = txtCompanyID.Text
                    .COMPANYNAME = txtCompanyName.Text
                    .ADDRESS = txtAddress.Text
                    .BUSINESSNO = txtBusinessNo.Text
                    .FAXNO = txtFaxNo.Text
                    .EMAIL = txtEmail.Text
                End With
                
                MsgBox "Information saved successfully!", vbInformation
                Unload Me
            End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
    Me.BackColor = MAIN.ACPMenu.BackColor
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
Dim sSQL As String
CenterForm frmBusiness


State = EditStateMode

sSQL = "SELECT Company_Info.* " & _
            "FROM Company_Info "

Set RS_COMPANY = New ADODB.Recordset
If RS_COMPANY.State = adStateOpen Then RS_COMPANY.Close
RS_COMPANY.Open sSQL, CN, adOpenDynamic, adLockOptimistic

With RS_COMPANY
    txtCompanyID.Text = .Fields("CompanyID")
    txtCompanyName.Text = .Fields("CompanyName")
    txtAddress.Text = .Fields("Address")
    txtBusinessNo.Text = .Fields("BusinessNo")
    txtFaxNo.Text = .Fields("Faxno")
    txtEmail.Text = .Fields("Email")
End With

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation, Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmBusiness = Nothing
Set RS_COMPANY = Nothing
End Sub

Private Sub txtAddress_GotFocus()
HLText txtAddress
End Sub

Private Sub txtAddress_LostFocus()
unHLText txtAddress

End Sub

Private Sub txtBusinessNo_GotFocus()
HLText txtBusinessNo
End Sub

Private Sub txtBusinessNo_LostFocus()
unHLText txtBusinessNo

End Sub

Private Sub txtCompanyID_GotFocus()
HLText txtCompanyID
End Sub

Private Sub txtCompanyID_LostFocus()
unHLText txtCompanyID

End Sub

Private Sub txtCompanyName_GotFocus()
HLText txtCompanyName
End Sub

Private Sub txtCompanyName_LostFocus()
unHLText txtCompanyName

End Sub

Private Sub txtEmail_GotFocus()
HLText txtEmail
End Sub

Private Sub txtEmail_LostFocus()
unHLText txtEmail

End Sub

Private Sub txtFaxNo_GotFocus()
HLText txtFaxNo
End Sub


Private Sub txtFaxNo_LostFocus()
unHLText txtFaxNo

End Sub
