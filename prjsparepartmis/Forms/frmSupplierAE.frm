VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "NSDataCombo.ocx"
Begin VB.Form frmSupplierAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New Entry"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin prjcmosxp.Liner Liner2 
      Height          =   30
      Left            =   120
      TabIndex        =   35
      Top             =   5520
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   53
   End
   Begin prjcmosxp.Liner Liner1 
      Height          =   30
      Left            =   0
      TabIndex        =   34
      Top             =   960
      Width           =   9375
      _ExtentX        =   16536
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
      ScaleWidth      =   9330
      TabIndex        =   31
      Top             =   0
      Width           =   9330
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmSupplierAE.frx":0000
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
         Caption         =   "SUPPLIER DETAILS"
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
   Begin VB.TextBox txtEmailAddress 
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
      Left            =   1515
      TabIndex        =   7
      Top             =   4050
      Width           =   3100
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
      Left            =   1515
      TabIndex        =   6
      Top             =   3690
      Width           =   3100
   End
   Begin VB.TextBox txtRemarks 
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
      Height          =   855
      Left            =   1515
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   4440
      Width           =   3135
   End
   Begin VB.TextBox txtTeleFaxNo 
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
      Left            =   6195
      TabIndex        =   8
      Top             =   3360
      Width           =   3030
   End
   Begin VB.TextBox txtZipcode 
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
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   4050
      Width           =   1320
   End
   Begin VB.TextBox txtMobileNo 
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
      Left            =   6195
      TabIndex        =   9
      Top             =   3690
      Width           =   3030
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
      Height          =   315
      Left            =   6195
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4440
      Width           =   3050
   End
   Begin VB.TextBox txtContactPerson 
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
      Left            =   1515
      TabIndex        =   5
      Top             =   3360
      Width           =   3100
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
      Height          =   705
      Left            =   1515
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1920
      Width           =   3100
   End
   Begin VB.TextBox txtSearchTerm 
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
      Left            =   6195
      TabIndex        =   3
      Top             =   1200
      Width           =   3000
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
      Left            =   1515
      TabIndex        =   1
      Top             =   1560
      Width           =   3100
   End
   Begin VB.ComboBox cboStatus 
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
      ItemData        =   "frmSupplierAE.frx":38CF
      Left            =   6195
      List            =   "frmSupplierAE.frx":38D9
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1560
      Width           =   1560
   End
   Begin VB.TextBox txtSupplierID 
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
      Left            =   1515
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   1665
   End
   Begin lvButton.lvButtons_H cmdHistory 
      Height          =   345
      Left            =   120
      TabIndex        =   16
      Top             =   5655
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
      Left            =   6840
      TabIndex        =   14
      Top             =   5655
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
      Left            =   8115
      TabIndex        =   15
      Top             =   5655
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
   Begin ctrlNSDataCombo.NSDataCombo txtCityTown 
      Height          =   315
      Left            =   6195
      TabIndex        =   10
      Top             =   4050
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   556
      ForeColor       =   0
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Contact Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   75
      TabIndex        =   17
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   255
      Left            =   45
      Top             =   2880
      Width           =   9210
   End
   Begin VB.Label Label19 
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
      Left            =   150
      TabIndex        =   30
      Top             =   4050
      Width           =   990
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefax No."
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
      Left            =   4920
      TabIndex        =   29
      Top             =   3330
      Width           =   840
   End
   Begin VB.Label Label17 
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
      Left            =   150
      TabIndex        =   28
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Business Tel."
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
      Left            =   150
      TabIndex        =   27
      Top             =   3690
      Width           =   930
   End
   Begin VB.Label Label7 
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
      Left            =   4920
      TabIndex        =   26
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No."
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
      Left            =   4920
      TabIndex        =   25
      Top             =   3720
      Width           =   750
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Province"
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
      Left            =   4920
      TabIndex        =   24
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
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
      Left            =   150
      TabIndex        =   23
      Top             =   3360
      Width           =   1110
   End
   Begin VB.Label Label14 
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
      Left            =   4920
      TabIndex        =   22
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label Label13 
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
      Left            =   150
      TabIndex        =   21
      Top             =   1920
      Width           =   585
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search Term"
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
      Left            =   4920
      TabIndex        =   20
      Top             =   1200
      Width           =   900
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
      Left            =   150
      TabIndex        =   19
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SupplierCD"
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
      Left            =   150
      TabIndex        =   18
      Top             =   1200
      Width           =   780
   End
End
Attribute VB_Name = "frmSupplierAE"
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
    
    DE = Format$(RS_SUPPLIER.Fields("DateEncoded"), "MMM-dd-yyyy HH:MM AMPM")
    DM = Format$(RS_SUPPLIER.Fields("LastDateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    EB = getValueAt("SELECT * FROM Suppliers WHERE SupplierID = '" & txtSupplierID.Text & "'", "EncodedBy")
    MB = getValueAt("SELECT * FROM Suppliers WHERE SupplierID = '" & txtSupplierID.Text & "'", "ModifiedBy")
    
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
            Set RS_SUPPLIER = New ADODB.Recordset
            sSQL_Insert "INSERT INTO Suppliers (SupplierID, Description, ContactPerson, BusinessNo, TelefaxNo, MobileNo, Email, Address, CityTown, Province, ZipCode, DateEncoded, EncodedBy,Status, SearchTerm,Remarks) VALUES ('" & txtSupplierID.Text & _
            "', '" & txtDescription.Text & "', '" & txtContactPerson.Text & "', '" & txtBusinessNo.Text & "', '" & txtTeleFaxNo.Text & "', '" & txtMobileNo.Text & _
            "', '" & txtEmailAddress.Text & "', '" & txtAddress.Text & "', '" & txtCityTown.Text & "', '" & txtProvince.Text & "', '" & txtZipcode.Text & "', '" & Format(Now, "M/d/yyyy") & "', '" & ACTIVE_USER.USERNAME & "', '" & cboStatus.Text & "', '" & txtSearchTerm.Text & _
            "','" & txtRemarks.Text & "')"
            
            MsgBox "New supplier has been successfully saved!", vbInformation
            Unload Me
            
            ElseIf State = EditStateMode Then
            
            Set RS_SUPPLIER = New ADODB.Recordset
            sSQL_Update "UPDATE Suppliers SET Description= '" & txtDescription.Text & "', ContactPerson= '" & txtContactPerson.Text & "', BusinessNo= '" & txtBusinessNo.Text & "', TelefaxNo= '" & txtTeleFaxNo.Text & "', MobileNo= '" & txtMobileNo.Text & _
            "', Email= '" & txtEmailAddress.Text & "', Address= '" & txtAddress.Text & "', CityTown= '" & txtCityTown.Text & "', Province= '" & txtProvince.Text & "', ZipCode= '" & txtZipcode.Text & "', LastDateModified= '" & Now & "', ModifiedBy= '" & ACTIVE_USER.USERNAME & "', Status= '" & cboStatus.Text & _
            "', SearchTerm= '" & txtSearchTerm.Text & "' WHERE SupplierID='" & txtSupplierID.Text & "'"
            
            MsgBox "Information saved successfully!", vbInformation
            Unload Me
            End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
Me.BackColor = MAIN.ACPMenu.BackColor

cboStatus.ListIndex = 0
txtSupplierID.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
Call InitializeNSD
CenterForm frmSupplierAE

If State = AddStateMode Then
    Me.Caption = "Create New Entry"
    txtSupplierID.Text = "SUP-" & Format(GenerateCD("Suppliers"), "00000")
    cboStatus.Text = "ACTIVE"
    cmdHistory.Enabled = False
Else
    Me.Caption = "Modify Entry"
    cmdHistory.Enabled = True
    
    srcSQL = "SELECT Suppliers.* " & _
                "FROM Suppliers " & _
                "WHERE (((Suppliers.SupplierID)='" & PK & "'))"

    Set RS_SUPPLIER = New ADODB.Recordset
    If RS_SUPPLIER.State = adStateOpen Then RS_SUPPLIER.Close
    RS_SUPPLIER.Open srcSQL, CN, adOpenDynamic, adLockOptimistic

    With RS_SUPPLIER
        txtSupplierID.Text = .Fields("SupplierID")
        txtDescription.Text = .Fields("Description")
        txtAddress.Text = .Fields("Address")
        txtSearchTerm.Text = .Fields("SearchTerm")
        cboStatus.Text = .Fields("Status")
        txtContactPerson.Text = .Fields("ContactPerson")
        txtBusinessNo.Text = .Fields("BusinessNo")
        txtEmailAddress.Text = .Fields("Email")
        txtTeleFaxNo.Text = .Fields("TelefaxNo")
        txtMobileNo.Text = .Fields("MobileNo")
        txtCityTown.Text = .Fields("CityTown")
        txtProvince.Text = .Fields("Province")
        txtZipcode.Text = .Fields("ZipCode")
        txtRemarks.Text = .Fields("Remarks")
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
Set frmSupplierAE = Nothing
Set RS_SUPPLIER = Nothing
frmSupplier.CommandPass "Refresh"
End Sub


Private Sub InitializeNSD()
On Error Resume Next
    With txtCityTown
        .ClearColumn
        .AddColumn "ZipCode", 1700.882
        .AddColumn "CityTown", 1800
        .AddColumn "Province", 3100
        .Connection = CN.ConnectionString
        
        .SQLFields = "ZipCode,CityTown,Province"
        .sqlTables = "zipcodes"
        .sqlSortOrder = "ZipCode ASC"
        .BoundField = "ZipCode"
        .PageBy = 10
        .DisplayCol = 2
        
        .setDropWindowSize 6500, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Select City/Town "
    End With
End Sub

Private Sub txtCityTown_Change()
On Error Resume Next
    With txtCityTown
        txtProvince.Text = .getSelValueAt(3)
        txtZipcode.Text = .getSelValueAt(1)
    End With
End Sub
