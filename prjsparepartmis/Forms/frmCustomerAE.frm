VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmCustomerAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New Entry"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin prjcmosxp.Liner Liner2 
      Height          =   30
      Left            =   120
      TabIndex        =   35
      Top             =   5280
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   53
   End
   Begin prjcmosxp.Liner Liner1 
      Height          =   30
      Left            =   0
      TabIndex        =   34
      Top             =   960
      Width           =   10335
      _ExtentX        =   18230
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
      ScaleWidth      =   10290
      TabIndex        =   31
      Top             =   0
      Width           =   10290
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmCustomerAE.frx":0000
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
         TabIndex        =   33
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER DETAILS"
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
         TabIndex        =   32
         Top             =   240
         Width           =   2175
      End
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
      Height          =   285
      Index           =   11
      Left            =   5520
      MaxLength       =   100
      TabIndex        =   12
      Tag             =   "Name"
      Top             =   4440
      Width           =   2600
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
      Height          =   285
      Index           =   10
      Left            =   5520
      MaxLength       =   100
      TabIndex        =   11
      Tag             =   "Name"
      Top             =   4080
      Width           =   2600
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
      Height          =   285
      Index           =   9
      Left            =   5520
      MaxLength       =   100
      TabIndex        =   10
      Tag             =   "Name"
      Top             =   3720
      Width           =   4155
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
      Height          =   285
      Index           =   8
      Left            =   1320
      MaxLength       =   100
      TabIndex        =   8
      Tag             =   "Name"
      Top             =   4440
      Width           =   2600
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
      Height          =   285
      Index           =   7
      Left            =   1320
      MaxLength       =   100
      TabIndex        =   7
      Tag             =   "Name"
      Top             =   4080
      Width           =   2600
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
      Height          =   285
      Index           =   6
      Left            =   1320
      MaxLength       =   100
      TabIndex        =   6
      Tag             =   "Name"
      Top             =   3720
      Width           =   2600
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
      Height          =   285
      Index           =   5
      Left            =   5520
      MaxLength       =   100
      TabIndex        =   5
      Tag             =   "Name"
      Top             =   2160
      Width           =   2600
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
      Height          =   285
      Index           =   4
      Left            =   5520
      MaxLength       =   100
      TabIndex        =   4
      Tag             =   "Name"
      Top             =   1800
      Width           =   2600
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
      Height          =   285
      Index           =   3
      Left            =   5520
      MaxLength       =   100
      TabIndex        =   3
      Tag             =   "Name"
      Top             =   1440
      Width           =   2600
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
      Left            =   1320
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Tag             =   "Name"
      Top             =   2160
      Width           =   3195
   End
   Begin VB.ComboBox cmbGender 
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
      ItemData        =   "frmCustomerAE.frx":38CF
      Left            =   1320
      List            =   "frmCustomerAE.frx":38D9
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4800
      Width           =   2010
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
      Height          =   285
      Index           =   1
      Left            =   1320
      MaxLength       =   100
      TabIndex        =   1
      Tag             =   "Name"
      Top             =   1815
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
      Height          =   285
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1965
   End
   Begin lvButton.lvButtons_H cmdHistory 
      Height          =   345
      Left            =   120
      TabIndex        =   15
      Top             =   5520
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
      Left            =   7800
      TabIndex        =   13
      Top             =   5520
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
      Left            =   9075
      TabIndex        =   14
      Top             =   5520
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
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
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
      Index           =   13
      Left            =   4800
      TabIndex        =   30
      Top             =   4080
      Width           =   360
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Website"
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
      Index           =   14
      Left            =   4800
      TabIndex        =   29
      Top             =   4440
      Width           =   585
   End
   Begin VB.Label Labels 
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
      Index           =   15
      Left            =   4800
      TabIndex        =   28
      Top             =   3720
      Width           =   630
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      Index           =   10
      Left            =   120
      TabIndex        =   27
      Top             =   4890
      Width           =   525
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Landline"
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
      Index           =   7
      Left            =   4800
      TabIndex        =   26
      Top             =   1785
      Width           =   585
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Middlename"
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
      Index           =   6
      Left            =   120
      TabIndex        =   25
      Top             =   4530
      Width           =   840
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
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
      Index           =   5
      Left            =   4800
      TabIndex        =   24
      Top             =   1395
      Width           =   450
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Firstname"
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
      Index           =   4
      Left            =   120
      TabIndex        =   23
      Top             =   4140
      Width           =   705
   End
   Begin VB.Label Labels 
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
      Index           =   3
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   690
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
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
      TabIndex        =   21
      Top             =   1815
      Width           =   675
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerID"
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
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Labels 
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
      Index           =   12
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   585
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
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
      Index           =   9
      Left            =   4800
      TabIndex        =   18
      Top             =   2160
      Width           =   270
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Basic Information"
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
      TabIndex        =   17
      Top             =   1080
      Width           =   2865
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person Details"
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
      TabIndex        =   16
      Top             =   3360
      Width           =   3465
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000010&
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   120
      Top             =   3360
      Width           =   10095
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000010&
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   120
      Top             =   1080
      Width           =   10095
   End
End
Attribute VB_Name = "frmCustomerAE"
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
    
    DE = Format$(RS_CUSTOMER.Fields("DateEncoded"), "MMM-dd-yyyy HH:MM AMPM")
    DM = Format$(RS_CUSTOMER.Fields("LastDateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    EB = getValueAt("SELECT * FROM Customers WHERE CustomerID = '" & txtEntry(0).Text & "'", "EncodedBy")
    MB = getValueAt("SELECT * FROM Customers WHERE CustomerID = '" & txtEntry(0).Text & "'", "ModifiedBy")
    
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
            If TypeOf obj Is TextBox Then
                If obj.Text = "" Then
                    MsgBox obj.Name & " could not be left blank. Please complete the field.", vbExclamation
                    obj.SetFocus
                    Exit Sub
                End If
            End If
            Next obj
            
            If State = AddStateMode Then
                RS_CUSTOMER.AddNew
                RS_CUSTOMER.Fields("CustomerID") = txtEntry(0).Text
                RS_CUSTOMER.Fields("Description") = txtEntry(1).Text
                RS_CUSTOMER.Fields("Address") = txtEntry(2).Text
                RS_CUSTOMER.Fields("MobileNo") = txtEntry(3).Text
                RS_CUSTOMER.Fields("LandlineNo") = txtEntry(4).Text
                RS_CUSTOMER.Fields("FaxNo") = txtEntry(5).Text
                RS_CUSTOMER.Fields("Lastname") = txtEntry(6).Text
                RS_CUSTOMER.Fields("Firstname") = txtEntry(7).Text
                RS_CUSTOMER.Fields("Middlename") = txtEntry(8).Text
                RS_CUSTOMER.Fields("Gender") = cmbGender.Text
                
                RS_CUSTOMER.Fields("OwnerName") = txtEntry(9).Text
                RS_CUSTOMER.Fields("Email") = txtEntry(10).Text
                RS_CUSTOMER.Fields("Website") = txtEntry(11).Text
                RS_CUSTOMER.Fields("DateEncoded") = Format(Now, "M/d/yyyy")
                RS_CUSTOMER.Fields("EncodedBy") = ACTIVE_USER.USERNAME
                RS_CUSTOMER.Update
                
                MsgBox "New customer record has been successfully saved!", vbInformation
                
                frmCustomer.CommandPass "Refresh"
                Unload Me
            
            ElseIf State = EditStateMode Then
                RS_CUSTOMER.Fields("CustomerID") = txtEntry(0).Text
                RS_CUSTOMER.Fields("Description") = txtEntry(1).Text
                RS_CUSTOMER.Fields("Address") = txtEntry(2).Text
                RS_CUSTOMER.Fields("MobileNo") = txtEntry(3).Text
                RS_CUSTOMER.Fields("LandlineNo") = txtEntry(4).Text
                RS_CUSTOMER.Fields("FaxNo") = txtEntry(5).Text
                RS_CUSTOMER.Fields("Lastname") = txtEntry(6).Text
                RS_CUSTOMER.Fields("Firstname") = txtEntry(7).Text
                RS_CUSTOMER.Fields("Middlename") = txtEntry(8).Text
                RS_CUSTOMER.Fields("Gender") = cmbGender.Text
                RS_CUSTOMER.Fields("OwnerName") = txtEntry(9).Text
                RS_CUSTOMER.Fields("Email") = txtEntry(10).Text
                RS_CUSTOMER.Fields("Website") = txtEntry(11).Text
                RS_CUSTOMER.Fields("LastDateModified") = Now
                RS_CUSTOMER.Fields("ModifiedBy") = ACTIVE_USER.USERNAME
                RS_CUSTOMER.Update
                
                MsgBox "Information saved successfully!", vbInformation, Me.Caption
                
                frmCustomer.CommandPass "Refresh"
                Unload Me
            
            End If

End Sub

Private Sub Form_Activate()
On Error Resume Next
txtEntry(0).SetFocus

cmbGender.ListIndex = 0
Me.BackColor = MAIN.ACPMenu.BackColor
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
CenterForm frmCustomerAE

If State = AddStateMode Then
    Me.Caption = "Create New Entry"
    txtEntry(0).Text = "CUS-" & Format(GenerateCD("Customers"), "00000")
    cmdHistory.Enabled = False
Else
    Me.Caption = "Modify Entry"
    cmdHistory.Enabled = True
    
    srcSQL = "SELECT Customers.* " & _
                "FROM Customers " & _
                "WHERE (((Customers.CustomerID)='" & PK & "'))"

    Set RS_CUSTOMER = New ADODB.Recordset
    If RS_CUSTOMER.State = adStateOpen Then RS_CUSTOMER.Close
    RS_CUSTOMER.Open srcSQL, CN, adOpenDynamic, adLockOptimistic

    With RS_CUSTOMER
        txtEntry(0).Text = .Fields("CustomerID")
        txtEntry(1).Text = .Fields("Description")
        txtEntry(2).Text = .Fields("Address")
        txtEntry(3).Text = .Fields("MobileNo")
        txtEntry(4).Text = .Fields("LandlineNo")
        txtEntry(5).Text = .Fields("FaxNo")
        txtEntry(6).Text = .Fields("Lastname")
        txtEntry(7).Text = .Fields("Firstname")
        txtEntry(8).Text = .Fields("Middlename")
        cmbGender.Text = .Fields("Gender")
        txtEntry(9).Text = .Fields("OwnerName")
        txtEntry(10).Text = .Fields("Email")
        txtEntry(11).Text = .Fields("Website")
    End With

End If

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
Set frmCustomerAE = Nothing
Set RS_CUSTOMER = Nothing
End Sub


Private Sub txtEntry_GotFocus(Index As Integer)
HLText txtEntry(Index)
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
unHLText txtEntry(Index)

If Index = 6 Or Index = 7 Or Index = 8 Then
    txtEntry(9).Text = txtEntry(6).Text & "," & txtEntry(7).Text & Space(1) & Mid$(txtEntry(8).Text, 1, 1) & "."
End If
End Sub
