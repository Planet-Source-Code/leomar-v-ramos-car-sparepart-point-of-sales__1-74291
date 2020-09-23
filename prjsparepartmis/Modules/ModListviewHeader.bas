Attribute VB_Name = "ModListviewHeader"
'Listview Consts
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Public Sub lvSizeColumns(lv As ListView)
Dim Counter As Long
    'Resizes Listview Column Headers.
    For Counter = 0 To (lv.ColumnHeaders.Count - 1)
        Call SendMessage(lv.hWnd, LVM_SETCOLUMNWIDTH, Counter, _
        ByVal LVSCW_AUTOSIZE_USEHEADER)
    Next Counter
End Sub

