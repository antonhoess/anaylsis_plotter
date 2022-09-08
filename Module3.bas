Attribute VB_Name = "Module3"
Public Declare Function ClipCursor Lib "user32" (Rect As Rect) As Long
Public Declare Sub FreeCursor Lib "user32" Alias "ClipCursor" (ByVal Rect As Long)
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, Rect As Rect) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'Public Winrect As Rect

