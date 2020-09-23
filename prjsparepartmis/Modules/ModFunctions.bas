Attribute VB_Name = "ModFunctions"
Option Explicit

Public Const CB_FINDSTRING = &H14C
Public Const CB_ERR = (-1)
Public Const CB_SHOWDROPDOWN = &H14F
Declare Function SendMessage Lib "user32" Alias _
                                 "SendMessageA" _
                                 (ByVal hWnd As Long, _
                                  ByVal wMsg As Long, _
                                  ByVal wParam As Long, _
                                  lParam As Any) As Long
                                  
'Function used to check if the record exist in Flex grid
Public Function getFlexPos(ByVal srcFlexGrd As MSHFlexGrid, ByVal srcWhatCol As Integer, ByVal srcFindWhat As String) As Integer
    Dim R As Long, ret As Integer
    
    ret = -1 'Means not found
    For R = 0 To srcFlexGrd.Rows - 1
        If srcFlexGrd.TextMatrix(R, srcWhatCol) = srcFindWhat Then ret = R: Exit For
    Next R
    
    getFlexPos = ret
    R = 0: ret = 0
End Function
                                  
                                  'Function that return the current index for a certain table
Public Function getIndex(ByVal srcTable As String) As Long
    On Error GoTo err
    Dim RS As New Recordset
    Dim RI As Long
    
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM PK_Generator WHERE TableName = '" & srcTable & "'", CN, adOpenStatic, adLockOptimistic
    
    RI = RS.Fields("NextNo")
    RS.Fields("NextNo") = RI + 1
    RS.Update
    
    getIndex = RI
    
    srcTable = ""
    RI = 0
    Set RS = Nothing
    Exit Function
err:
        ''Error when incounter a null value
        If err.Number = 94 Then
            getIndex = 1
            Resume Next
        Else
            MsgBox err.Description
        End If
End Function

Public Function GenerateCD(ByVal srcTable As String) As Integer
    Dim srcRS As New ADODB.Recordset
    If srcRS.State = 1 Then srcRS.Close
    srcRS.Open "SELECT MAX(Auto) FROM " & srcTable, CN, adOpenStatic, adLockReadOnly
    If IsNull(srcRS(0)) = True Then
        GenerateCD = 1
    Else
        GenerateCD = Val(srcRS(0)) + 1
    End If
End Function

'Function used to check if the record exit or not.
Public Function isRecordExist(ByVal sTable As String, ByVal sField As String, ByVal sStr As String, Optional isString As Boolean) As Boolean
    Dim RS As New Recordset

    RS.CursorLocation = adUseClient
    If isString = False Then
        RS.Open "Select * From " & sTable & " Where " & sField & " = " & sStr, CN, adOpenStatic, adLockOptimistic
    Else
        RS.Open "Select * From " & sTable & " Where " & sField & " = '" & sStr & "'", CN, adOpenStatic, adLockOptimistic
    End If
    If RS.RecordCount < 1 Then
        isRecordExist = False
    Else
        isRecordExist = True
    End If
    Set RS = Nothing
End Function

'Function used to check if the Ascii is a number or not (return 0 if number)
Public Function isNumber(ByVal sKeyAscii) As Integer
    If Not ((sKeyAscii >= 48 And sKeyAscii <= 57) Or sKeyAscii = 8 Or sKeyAscii = 46) Then
        isNumber = 0
    Else
        isNumber = sKeyAscii
    End If
End Function

'Function used to left split user fields
Public Function LeftSplitUF(ByVal srcUF As String) As String
    If srcUF = "*~~~~~*" Then LeftSplitUF = "": Exit Function
    Dim i As Integer
    Dim t As String
    For i = 1 To Len(srcUF)
        If Mid$(srcUF, i, 7) = "*~~~~~*" Then
            Exit For
        Else
            t = t & Mid$(srcUF, i, 1)
        End If
    Next i
    LeftSplitUF = t
    i = 0
    t = ""
End Function

'Function used to right split user fields
Public Function RightSplitUF(ByVal srcUF As String) As String
    If srcUF = "*~~~~~*" Then RightSplitUF = "": Exit Function
    Dim i As Integer
    Dim t As String
    For i = (InStr(1, srcUF, "*~~~~~*", vbTextCompare) + 7) To Len(srcUF)
        t = t & Mid$(srcUF, i, 1)
    Next i
    RightSplitUF = t
    i = 0
    t = ""
End Function

'Function that return true if the control is empty
Public Function is_empty(ByRef sText As Variant, Optional UseTagValue As Boolean) As Boolean
    On Error Resume Next
    If sText.Text = "" Then
        is_empty = True
        If UseTagValue = True Then
            MsgBox "The field '" & sText.Tag & "' is required.Please check it!", vbExclamation
        Else
            MsgBox "The field is required.Please check it!", vbExclamation
        End If
        sText.SetFocus
    Else
        is_empty = False
    End If
End Function

Public Function isCurrency(ByVal sKeyAscii, strCur As String) As Integer
    Dim i As Integer
    Dim intDot As Integer
    intDot = 0
    If Not ((sKeyAscii >= 48 And sKeyAscii <= 57) Or sKeyAscii = 8 Or sKeyAscii = 46) Then
        isCurrency = 0
    Else
        If sKeyAscii = 46 Then
            For i = 1 To Len(strCur)
                If Mid(strCur, i, 1) = "." Then
                    intDot = intDot + 1
                End If
            Next
            If intDot < 1 Then
                isCurrency = sKeyAscii
            Else
                isCurrency = 0
            End If
        Else
            isCurrency = sKeyAscii
        End If
    End If
End Function

'Function used to change the yes/no value
Public Function changeYNValue(ByVal srcStr As String) As String
    Select Case srcStr
        Case "Y": changeYNValue = "1"
        Case "N": changeYNValue = "0"
        Case "1": changeYNValue = "Y"
        Case "0": changeYNValue = "N"
    End Select
End Function

'Function that return true if the control is numeric
Public Function is_numeric(ByRef sText As String) As Boolean
    If IsNumeric(sText) = False Then
        is_numeric = False
        MsgBox "The field required a numeric input.Please check it!", vbExclamation
    Else
        is_numeric = True
    End If
End Function

Public Function Date_To_MMDDYY(ByVal strDate As String) As String
 Date_To_MMDDYY = Mid$(strDate, 1, 2) & Mid$(strDate, 4, 2) & Mid$(strDate, 7, 2)
End Function

'Function that return the value of a certain field
Public Function getValueAt(ByVal srcSQL As String, ByVal whichField As String) As String
    Dim RS As New Recordset
    
    RS.CursorLocation = adUseClient
    RS.Open srcSQL, CN, adOpenStatic, adLockReadOnly
    If RS.RecordCount > 0 Then getValueAt = RS.Fields(whichField)
    
    Set RS = Nothing
End Function

Public Function toNumber(ByVal srcCurrency As String, Optional RetZeroIfNegative As Boolean) As Double
    If srcCurrency = "" Then
        toNumber = 0
    Else
        Dim retValue As Double
        If InStr(1, srcCurrency, ",") > 0 Then
            retValue = Val(Replace(srcCurrency, ",", "", , , vbTextCompare))
        Else
            retValue = Val(srcCurrency)
        End If
        If RetZeroIfNegative = True Then
            If retValue < 1 Then retValue = 0
        End If
        toNumber = retValue
        retValue = 0
    End If
End Function

'Function that return the count of the rows in the table
Public Function getRecordCount(ByVal srcTable As String, Optional srcCondition As String, Optional isFormatted As Boolean) As String
    If srcCondition <> "" Then srcCondition = " " & srcCondition
    Dim RS As New Recordset
    
    RS.CursorLocation = adUseClient
    RS.Open "SELECT COUNT(PK) as TCount FROM " & srcTable & srcCondition, CN, adOpenStatic, adLockReadOnly
    If isFormatted = True Then
        getRecordCount = Format$(RS![TCount], "#,##0")
    Else
        getRecordCount = RS![TCount]
    End If
    Set RS = Nothing
End Function

'Function that will return a currenct format
Public Function toMoney(ByVal srcCurr As String) As String
   toMoney = Format$(srcCurr, "#,##0.00")
End Function

'Function used to determine if the object has been set
Public Function isObjectSet(srcObject As Object) As Boolean
    On Error GoTo err
    'I use tag because almost all controls have this
    srcObject.Tag = srcObject.Tag
    isObjectSet = True
    
    Exit Function
err:
    isObjectSet = False
End Function

'Function used to get the sum  of fields
Public Function getSumOfFields(ByVal sTable As String, ByVal sField As String, ByRef sCN As ADODB.Connection, Optional inclField As String, Optional sCondition As String) As Double
    On Error GoTo err
    Dim RS As New ADODB.Recordset

    RS.CursorLocation = adUseClient
    If sCondition <> "" Then sCondition = " GROUP BY " & inclField & " HAVING(" & sCondition & ")"
    If inclField <> "" Then inclField = "," & inclField
    RS.Open "SELECT Sum(" & sTable & "." & sField & ") AS fTotal" & inclField & " FROM " & sTable & sCondition, sCN, adOpenStatic, adLockOptimistic
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        Do While Not RS.EOF
            getSumOfFields = getSumOfFields + RS.Fields("fTotal")
            RS.MoveNext
        Loop
    Else
        getSumOfFields = 0
    End If
    
    Set RS = Nothing
    Exit Function
err:
        'Error when incounter a null value
        If err.Number = 94 Then getSumOfFields = 0: Resume Next
End Function

' ExportExcel.bas
'
'*****************************************************
' ExportListbox
' Purpose:   Exports listbox information to a MS Excel
'            spreadsheet.
' Inputs:
'   pListview:      Listview control reference
'   pFilename:      Name of Excel file to create.
'   Append:         Appends an existing spreadsheet
'
' Returns:          True if successful
'                   False otherwise.
'*****************************************************

Public Function ExportListview(ByRef pListview As MSComctlLib.ListView, _
    ByVal pFilename As String, _
    Optional ByVal WorksheetName As String = "Sheet1", _
    Optional Append As Boolean = False) As Boolean
    
    Dim CN As Object
    Dim CAT As Object
    Dim TBL As Object
    Dim COL As Object
    Dim strConnection As String
    Dim AListItem As MSComctlLib.ListItem
    Dim AColumnHeader As MSComctlLib.ColumnHeader
    Dim RS As Object
    Dim intLoop As Integer
    Dim intLoop2 As Integer
    
    On Error GoTo ErrHandler
    
    ' Make sure everything is ok with the inputs before
    ' continuing.
    ' pListView
    If pListview.View <> lvwReport Then
        MsgBox "Listview must be in Report mode.", _
            vbCritical + vbOKOnly, "ExportListview"
        GoTo NotSuccessful
    End If
    ' pFilename
    If Trim$(pFilename) = vbNullString Then
        MsgBox "No filename given.", vbCritical + vbOKOnly, _
            "ExportListview"
        GoTo NotSuccessful
    End If
    ' **********
    Set CN = CreateObject("ADODB.Connection")
    
    ' Create a connection to the Excel file using Jet's ISAM
    strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Extended Properties=Excel 8.0;" & _
        "Data Source=" & pFilename
    CN.Open strConnection
    
    ' No need to create a workbook if the spreadsheet already exists
    If Append Then GoTo AlreadyExists
    
    ' Create a Excel Workbook and set the connection to CN
    Set CAT = CreateObject("ADOX.Catalog")
    CAT.ActiveConnection = CN
    
    ' Create a worksheet for the cat
    Set TBL = CreateObject("ADOX.Table")
    TBL.Name = WorksheetName
    
    ' Do the column headers
    
    For Each AColumnHeader In pListview.ColumnHeaders
        Set COL = CreateObject("ADOX.Column")
        COL.Type = 130  ' adWChar
        COL.Name = AColumnHeader.Text
        TBL.Columns.Append COL
        Set COL = Nothing
    Next AColumnHeader
    
    ' Add this worksheet to the workbook
    CAT.Tables.Append TBL
    
AlreadyExists:
    Set RS = CreateObject("ADODB.Recordset")
    
    ' open the excel file that was just created as a recordset
    ' so we can add records.
    RS.Open WorksheetName, CN, 1, 3
    
    ' Grab every listitem out of the listview control
    
    For Each AListItem In pListview.ListItems
        ' Listitem and then all subitems
        RS.AddNew
        RS.Fields(0) = AListItem.Text
        ' subitems
        For intLoop = 1 To RS.Fields.Count - 1
            RS.Fields(intLoop) = AListItem.SubItems(intLoop)
        Next intLoop
        RS.Update
    Next AListItem
    
        

    ' Mark as success
    ExportListview = True
    GoTo CloseAndNothing
    
NotSuccessful:
    ExportListview = False
    
    ' clear all objects and exit
CloseAndNothing:
    On Error Resume Next
    RS.Close
    CN.Close
    Set CAT = Nothing
    Set CN = Nothing
    Set TBL = Nothing
    Set COL = Nothing
    Set AListItem = Nothing
    Set AColumnHeader = Nothing
    Set RS = Nothing
    Exit Function
    
ErrHandler:
    ' simply raise the error to the client
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    GoTo CloseAndNothing
End Function





