Attribute VB_Name = "modMouse"
Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Function GetX()
Dim p As POINTAPI
GetCursorPos p
GetX = p.x
End Function

Public Function GetY()
Dim p As POINTAPI
GetCursorPos p
GetY = p.y
End Function
