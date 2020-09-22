Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer 'The key states
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer 'the key states

Public Function GetCapslock() As Boolean
    GetCapslock = CBool(GetKeyState(vbKeyCapital) And 1) 'Return or set the Capslock toggle.
End Function

Public Function GetShift() As Boolean
    GetShift = CBool(GetAsyncKeyState(vbKeyShift)) 'Return or set the Capslock toggle.
End Function

Function Shf(Shift, Char1, Char2)
'This function is exactly like the IIf function
'except without the Shift statement being present
'this relies on when you press the shift key and
'another key pressed at the same time
    If GetShift = True Then
        Shift = 1 'If shift is present
        Shf = Char1 'then the first character is displayed
    Else
        Shift = 0 'if shift isn't present
        Shf = Char2 'then the second character is displayed
    End If
End Function
