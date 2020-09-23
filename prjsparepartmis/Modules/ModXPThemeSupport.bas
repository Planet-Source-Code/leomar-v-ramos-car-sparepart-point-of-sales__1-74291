Attribute VB_Name = "ModXPThemeSupport"
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Public Const ICC_USEREX_CLASSES        As Long = &H200

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type



