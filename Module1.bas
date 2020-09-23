Attribute VB_Name = "Module1"

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Public Declare Function InternetGetConnectedState _
Lib "wininet.dll" (ByRef lpdwFlags As Long, _
ByVal dwReserved As Long) As Long

Public Function Online() As Boolean 'Apparently doesnt work
On Error GoTo 2001
Online = InternetGetConnectedState(0&, 0&)
Exit Function
2001
Online = False
End Function


Sub over() 'tries its best to exit game
Do
DoEvents
'MsgBox "working"
End
Loop
End Sub
Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)


    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hWnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub


