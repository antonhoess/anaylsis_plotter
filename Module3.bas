Attribute VB_Name = "Module3"
Public Declare Function ClipCursor Lib "user32" (Rect As Rect) As Long
Public Declare Sub FreeCursor Lib "user32" Alias "ClipCursor" (ByVal Rect As Long)
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, Rect As Rect) As Long
Public Type Rect
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
'Public Winrect As Rect
'Public Const HWND_TOPMOST = -1
'Public Const HWND_NOTOPMOST = -2

