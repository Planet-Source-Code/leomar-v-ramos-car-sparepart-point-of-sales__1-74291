VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSalesTotalPerCustomer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Total Per Customer"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   Icon            =   "frmSalesTotalPerCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjcmosxp.Liner Liner2 
      Height          =   30
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   53
   End
   Begin prjcmosxp.Liner Liner1 
      Height          =   30
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   4455
      _ExtentX        =   7858
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
      ScaleWidth      =   4830
      TabIndex        =   5
      Top             =   0
      Width           =   4830
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER SALES DATE PARAMETER"
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
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmSalesTotalPerCustomer.frx":08CA
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
         TabIndex        =   6
         Top             =   480
         Width           =   2895
      End
   End
   Begin MSComCtl2.DTPicker DTFrom 
      Height          =   315
      Left            =   675
      TabIndex        =   0
      Top             =   1200
      Width           =   1600
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "M/d/yyyy"
      Format          =   16515075
      CurrentDate     =   40544
   End
   Begin MSComCtl2.DTPicker DTTo 
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Top             =   1200
      Width           =   1600
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "M/d/yyyy"
      Format          =   16515075
      CurrentDate     =   40544
   End
   Begin lvButton.lvButtons_H cmdLoad 
      Height          =   345
      Left            =   2400
      TabIndex        =   2
      Top             =   2040
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   609
      Caption         =   "&Done"
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
      Left            =   3480
      TabIndex        =   3
      Top             =   2040
      Width           =   1005
      _ExtentX        =   1773
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
      Caption         =   "FROM:"
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
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image picArrow 
      Height          =   255
      Left            =   2400
      Top             =   1200
      Width           =   255
   End
End
Attribute VB_Name = "frmSalesTotalPerCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sSQL                        As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdLoad_Click()
On Error GoTo ErrLoad

sSQL = "SELECT qry_Sales_Order.* " & _
        "FROM qry_Sales_Order " & _
        " WHERE StatusID='CM' AND StatusDesc='COMPLETED' AND Date BETWEEN '" & CDate(DTFrom.Value) & _
        "' AND '" & CDate(DTTo.Value) & "'"
        
        
    With rptSalesTotalPerCustomer
        .rptCN.Connect = ";DATABASE=" & App.Path & "\Database\CMOSXP_DB.mdb;PWD=qwerty123;"
        .rptCN.DatabaseName = App.Path & "\Database\CMOSXP_DB.mdb"
        .rptCN.RecordSource = sSQL
        
        .lblCompany.Caption = ACTIVE_COMPANY.COMPANYNAME
        .lblAddress.Caption = ACTIVE_COMPANY.ADDRESS
        
        .lblFrom.Caption = DTFrom.Value
        .lblTo.Caption = DTTo.Value
        
        .txtSalesOrderID.DataField = "SalesOrderID"
        .txtCustomerID.DataField = "CustomerID"
        .txtTransDate.DataField = "Date"
        .txtSalesman.DataField = "Salesman"
        .txtGrossAmount.DataField = "Gross"
        .txtNetAmount.DataField = "NetAmount"
        .txtPayment.DataField = "FOP"
        .txtStatus.DataField = "StatusID"
        .txtStatusDesc.DataField = "StatusDesc"
        
        .show vbModal
    End With

    Unload Me

Exit Sub
ErrLoad:
   MsgBox "Error #:" & err.Number & vbNewLine & _
            "Description:" & err.Description, vbExclamation
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
ElseIf KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
CenterForm frmSalesTotalPerCustomer

On Error GoTo ErrHandler
picArrow.Picture = MAIN.i16x16.ListImages(3).Picture

DTFrom.Value = Format$(Now, "M/d/yyyy")
DTTo.Value = Format$(Now, "M/d/yyyy")

Exit Sub
ErrHandler:
    MsgBox "Error #:" & err.Number & vbNewLine & _
            "Description:" & err.Description, vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmSalesTotalPerCustomer = Nothing
End Sub


