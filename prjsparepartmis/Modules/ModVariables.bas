Attribute VB_Name = "ModVariables"
Option Explicit

Global END_APP                              As Boolean

Public RS_USER                              As New ADODB.Recordset
Public RS_USERTYPE                          As New ADODB.Recordset
Public RS_ZIPCODE                           As New ADODB.Recordset
Public RS_SUPPLIER                          As New ADODB.Recordset
Public RS_CUSTOMER                          As New ADODB.Recordset

Public RS_CARMAKE                           As New ADODB.Recordset
Public RS_CARTYPE                           As New ADODB.Recordset
Public RS_PCATEGORY                         As New ADODB.Recordset
Public RS_SPAREPART                         As New ADODB.Recordset
Public RS_COMPANY                           As New ADODB.Recordset

Public RS_SALES                             As New ADODB.Recordset
Public RS_PURCHASE                          As New ADODB.Recordset


Public ACTIVE_USER                          As USER_INFO
Public ACTIVE_COMPANY                       As COMPANY_INFO

Public XLSFILENAME                          As String

Public COMMAND_INSERT                       As New ADODB.Command
Public COMMAND_UPDATE                       As New ADODB.Command
Public COMMAND_DELETE                       As New ADODB.Command





