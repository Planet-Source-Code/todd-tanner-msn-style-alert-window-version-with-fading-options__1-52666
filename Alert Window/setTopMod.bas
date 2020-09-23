Attribute VB_Name = "setTopMod"
Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
  
Const HWND_TOPMOST = -1
Const HWND_NOTTOPMOST = -2
Const SWP_NOSIZE = 1
Const SWP_NOMOVE = 2

Public Sub setTop(winHwnd As Long, Optional setT As Boolean = True)
    Dim flags As Integer
    flags = SWP_NOSIZE Or SWP_NOMOVE
    If setT Then
        lR = SetWindowPos(winHwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    Else
        lR = SetWindowPos(winHwnd, HWND_NOTTOPMOST, 0, 0, 0, 0, flags)
    End If
End Sub
