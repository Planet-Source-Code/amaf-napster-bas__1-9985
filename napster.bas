Attribute VB_Name = "napster"
' napster.bas (vb6) created by amaf
' first ever .bas for napster
' url: www.envy.nu/amaf
' email: amaf@email.com


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Declare Function sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function iswindowenabled Lib "user32" Alias "IsWindowEnabled" (ByVal hwnd As Long) As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26
Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const ENTER_KEY = 13
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE
Public Sub about()
' this module was created by amaf
' -------------------------------
' i created it in about 2 days with very short time
' so that's why it has so many bugs, but i wanted
' programmers to get a feel of how to get started
' with napster api. i will be releasing a newer
' version soon. so check my site out daily. i want
' atleast every programmer to add a new sub or fun.
' that i haven't added and release it on PSC. i
' would like it if i got some credit for this .bas

'-amaf
End Sub
Function napster_extracterror()
' this will extract the error from the error message
' very simple code
napster_extracterror = GetText(FindWindowEx(napster_error, 0, "RICHEDIT", ""))
End Function
Public Sub click_long(a As Long)
' this will click either a button or tab
Call PostMessage(a&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(a&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub napster_menu_getuserinfo()
' this is not working
' i released it to early
main& = FindWindow("NAPSTER", "Napster v2.0 BETA 7")
finger& = FindWindow("#32770", vbNullString)
RunMenuByString main&, "&Get User Information"
a$ = GetText(finger&)
End Sub
Public Sub napster_getuserinfo(user$)
' this is used to finger the user
Dim a%, b%, c%
Dim main As Long
main& = FindWindow("NAPSTER", vbNullString)
a% = Findchildbytitle(main&, "Finger a User")
If a% <> 0 Then
a% = FindWindow("#32770", "Finger a User")
b% = FindWindowEx(a%, 0, "Edit", "")
Call SendMessageByString(b%, WM_SETTEXT, 0&, user$)
c% = FindWindowEx(a%, 0, "Button", "OK")
Call SendMessage(c%, WM_KEYDOWN, VK_SPACE, 0)
Call SendMessage(c%, WM_KEYUP, VK_SPACE, 0)
Call SendMessage(c%, WM_KEYDOWN, VK_SPACE, 0)
Call SendMessage(c%, WM_KEYUP, VK_SPACE, 0)
End If
End Sub
Public Sub napster_connect()
' connect to napster
Dim a As Long
a& = FindWindow("NAPSTER", vbNullString)
RunMenuByString a&, "&Connect"
End Sub
Public Sub napster_disconnect()
' disconnect to napster
Dim a As Long
a& = FindWindow("NAPSTER", vbNullString)
RunMenuByString a&, "&Disconnect"
End Sub
Public Sub napster_close()
' close napster
Dim a As Long
a& = FindWindow("NAPSTER", vbNullString)
Call PostMessage(a&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub napster_toolbar_chat()
' not working yet :(
' problems with tab
Dim main As Long
Dim b As Long
Dim c As Long
main& = FindWindow("NAPSTER", vbNullString)
b& = FindWindowEx(main&, 0, "SysTabControl32", vbNullString)
c& = NextOfClassByCount(b&, "SysTabControl32", 5)
b& = FindWindowEx(main&, 0, "SysTabControl32", "")
click_long c&
End Sub
Public Function napster_connect_status()
' get connection status
Dim main As Long
Dim b As Long
Dim c As Long
main& = FindWindow("NAPSTER", vbNullString)
b& = FindWindowEx(main&, 0, "msctls_statusbar32", vbNullString)
napster_connect_status = GetText(b&)
End Function
Public Sub napster_movedown()
' not working yet :(
' this is that vscroll on the side
' it will move it to the side
Dim main As Long
Dim b As Long
Dim c As Long
main& = FindWindow("NAPSTER", vbNullString)
b& = FindWindowEx(main&, 0, "SysTabControl32", vbNullString)
c& = NextOfClassByCount(b&, "msctls_updown32", 0)
click_long c&
End Sub
Public Sub napster_punt(user$, message$)
' this is to punt the user
' but i closed the main chat win. to stop
' you from feeling the lag
Dim main As Long
Dim a As Long
Dim b As Long
main& = FindWindow("NAPSTER", vbNullString)
a& = FindWindow("#32770", "(" + user$ + ") Instant Message")
If a& <> 0 Then
b& = FindWindowEx(a&, 0, "RICHEDIT", "")
b& = NextOfClassByCount(a&, "RICHEDIT", 2)
Call PostMessage(b&, WM_CLOSE, 0&, 0&)
b& = NextOfClassByCount(b, "RICHEDIT", 1)
Call SendMessageByString(b&, WM_SETTEXT, 0&, message$)
End If
End Sub
Public Sub napster_sendim(user$)
' this is to send a simple im
Dim a%, b%, c%
Dim MessageOk As Long
main& = FindWindow("NAPSTER", vbNullString)
a% = FindWindow("#32770", "Instant message a User")
If a% <> 0 Then
b% = FindWindowEx(a%, 0, "Edit", "")
Call SendMessageByString(b%, WM_SETTEXT, 0&, user$)
c% = FindWindowEx(a%, 0, "Button", "OK")
Call SendMessage(c%, WM_KEYDOWN, VK_SPACE, 0)
Call SendMessage(c%, WM_KEYUP, VK_SPACE, 0)
Call SendMessage(c%, WM_KEYDOWN, VK_SPACE, 0)
Call SendMessage(c%, WM_KEYUP, VK_SPACE, 0)
End If
End Sub
Public Sub napster_okonerror()
' will click 'ok' on error msg
Dim a%, b%
a% = FindWindow("#32770", "Napster notification")
b% = FindWindowEx(a%, 0, "Button", "OK")
Call SendMessage(b%, WM_KEYDOWN, VK_SPACE, 0)
Call SendMessage(b%, WM_KEYUP, VK_SPACE, 0)
Call SendMessage(b%, WM_KEYDOWN, VK_SPACE, 0)
Call SendMessage(b%, WM_KEYUP, VK_SPACE, 0)
End Sub
Function napster_error()
' error window
napster_error = FindWindow("#32770", "Napster notification")
End Function
Function Replace(ByVal c As String, ByVal d As String, ByVal E As String) As String
' replacing chars
a$ = c
Do While InStr(a$, d$) <> 0
DoEvents
b% = InStr(a$, d$)
If Err Then Exit Do
a$ = Left(a$, b% - 1) & E$ & Right(a$, Len(a$) - b% - Len(d$) + 1)
Loop
Replace = a$
End Function
Function GetText(handle As Long) As String
' extract text
Dim Length As Long, Text As String
Length& = SendMessage(handle&, WM_GETTEXTLENGTH, 0, 0)
Text$ = String(Length, 0)
Call SendMessageByString(handle&, WM_GETTEXT, Length& + 1, Text$)
GetText$ = Text
End Function
Function napster_ver()
' get the napster ver
napster_ver = Replace(GetText(napster_win), "Napster", "")
End Function
Function napster_win()
' main napster win
napster_win = FindWindow("NAPSTER", vbNullString)
End Function
Sub ClickMenu(ProgramHWND As Long, lngMenuIndex1 As Long, lngMenuIndex2 As Long)
' goes through menu
Dim MenuHWND As Long, SubMenuHWND As Long, MenuItemId As Long
MenuHWND& = GetMenu(ProgramHWND&)
SubMenuHWND& = GetSubMenu(MenuHWND&, lngMenuIndex1&)
MenuItemId& = GetMenuItemID(SubMenuHWND&, lngMenuIndex2&)
Call sendmessagebynum(ProgramHWND&, WM_COMMAND, MenuItemId&, 0&)
End Sub
Sub Pause(a)
' pause program
Dim b As Long
b = Timer
Do While Timer - b < a
DoEvents
Loop
End Sub
Private Sub StayOnTop(a As Form)
Call SetWindowPos(a.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub
Private Sub StayNotOnTop(a As Form)
Call SetWindowPos(a.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
End Sub
Public Sub WaitForWindowByTitle(ParentWindow As Long, WindowText As String)
    Dim FindThisWindow As Long
    Do: DoEvents
        FindThisWindow& = Findchildbytitle(ParentWindow&, WindowText$)
    Loop Until FindThisWindow& <> 0&
End Sub
Public Sub WaitForWindowByClass(ParentWindow As Long, ClassWindow As String)
    Dim FindThisWindow As Long
    Do: DoEvents
        FindThisWindow& = FindChildByClass(ParentWindow&, ClassWindow$)
    Loop Until FindThisWindow& <> 0&
End Sub
Public Sub RunMenuByString(ParentWindow As Long, StringToGet As String)
    Dim MenuHandle As Long, MenuItemCount As Long, NextItem As Long
    Dim SubMenu As Long, NextMenuItemCount As Long, MenuItemId As Long
    Dim NextNextItem As Long, NextMenuItemId As Long, FixedString As String
    MenuHandle& = GetMenu(ParentWindow&)
    MenuItemCount& = GetMenuItemCount(MenuHandle&)
    For NextItem& = 0& To MenuItemCount& - 1
        SubMenu& = GetSubMenu(MenuHandle&, NextItem&)
        NextMenuItemCount& = GetMenuItemCount(SubMenu&)
        For NextNextItem& = 0& To NextMenuItemCount& - 1
             NextMenuItemId& = GetMenuItemID(SubMenu&, NextNextItem&)
             FixedString$ = String(100, " ")
             Call GetMenuString(SubMenu&, NextMenuItemId&, FixedString$, 100, 1)
             If InStr(LCase(FixedString$), LCase(StringToGet$)) Then
                  Call SendMessageLong(ParentWindow&, WM_COMMAND, NextMenuItemId&, 0&)
                  Exit Sub
             End If
        Next NextNextItem&
    Next NextItem&
End Sub
Public Function NextOfClassByCount(ParentWindow As Long, ClassWindow As String, ByCount As Long) As Long
    Dim NextOfClass As Long, NextWindow As Long
    If ByCount& > ClassInstance(ParentWindow&, ClassWindow$) Then Exit Function
    If FindWindowEx(ParentWindow&, 0&, ClassWindow$, vbNullString) = 0& Then Exit Function
    For NextOfClass& = 1 To ByCount&
        NextWindow& = FindWindowEx(ParentWindow&, NextWindow&, ClassWindow$, vbNullString)
    Next NextOfClass&
    NextOfClassByCount& = NextWindow&
End Function
Public Function Findchildbytitle(ParentWindow As Long, WindowText As String) As Long
'NOT RECOMENDED FOR USE, THE 32 BIT API IN THIS MODULE IS MORE ACCURATE AND FASTER
    Dim GetChild As Long, GetNextChild As Long, PrepLong As Long
    GetChild& = GetWindow(ParentWindow&, 5)
    If UCase(GetCaption(GetChild&)) Like UCase(WindowText$) Then Findchildbytitle& = GetChild&
    GetChild& = GetWindow(ParentWindow&, 5)
    While ParentWindow&
        GetNextChild& = GetWindow(ParentWindow&, 5)
        If UCase(GetCaption(GetNextChild&)) Like UCase(WindowText$) & "*" Then Findchildbytitle& = GetChild&
           GetChild& = GetWindow(ParentWindow&, 5)
        If UCase(GetCaption(GetChild&)) Like UCase(WindowText$) & "*" Then Findchildbytitle& = GetChild&
    Wend
    Findchildbytitle& = 0&
End Function
Public Function FindChildByClass(ParentWindow As Long, ClassWindow As String) As Long
'For those who dont do 32 bit api i converted these oldschool 16 bit methods for you
    FindChildByClass& = FindWindowEx(ParentWindow&, 0&, ClassWindow$, vbNullString)
End Function

Public Function FindChildByTitleEx(ParentWindow As Long, WindowText As String) As Long
'For those who dont do 32 bit api i converted these oldschool 16 bit methods for you, with a twist
    FindChildByTitleEx& = FindWindowEx(ParentWindow&, 0&, vbNullString, WindowText$)
End Function
Public Function ClassInstance(ParentWindow As Long, ClassWindow As String) As Long
     Dim OnInstance As Long, CurrentCount As Long
     If FindWindowEx(ParentWindow&, 0&, ClassWindow$, vbNullString) = 0& Then Exit Function
     ClassInstance& = 0&
     Do: DoEvents
         OnInstance& = FindWindowEx(ParentWindow&, OnInstance&, ClassWindow$, vbNullString)
         If OnInstance& <> 0& Then
             CurrentCount& = CurrentCount& + 1
            Else
             Exit Do
         End If
     Loop
     ClassInstance& = CurrentCount&
End Function

Public Function ChildInstance(ParentWindow As Long) As Long
     Dim OnInstance As Long, CurrentCount As Long
     If IsWindow(ParentWindow&) = 0& Then Exit Function
     OnInstance& = GetWindow(ParentWindow&, 5)
     If OnInstance& <> 0& Then ChildInstance& = 1
     Do: DoEvents
         OnInstance& = GetWindow(OnInstance&, 2)
         If OnInstance& <> 0& Then CurrentCount& = CurrentCount& + 1
     Loop Until OnInstance& = 0&
     ChildInstance& = CurrentCount& + 1
End Function
Public Function GetCaption(WinHandle As Long) As String
    Dim buffer As String, TextLen As Long
    TextLen& = GetWindowTextLength(WinHandle&)
    buffer$ = String(TextLen&, 0&)
    Call GetWindowText(WinHandle&, buffer$, TextLen& + 1)
    GetCaption$ = buffer$
End Function




